from pathlib import Path
import json
from typing import Dict, Any, List, Optional
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timezone

import pandas as pd


# UEQ-S dimensions present in the logs (kept for later use)
# Order here matches the standard UEQ-S item numbering (1..8):
# 1: obstructive–supportive
# 2: complicated–easy
# 3: inefficient–efficient
# 4: confusing–clear
# 5: boring–exciting
# 6: not interesting–interesting
# 7: conventional–inventive
# 8: usual–leading edge
UEQ_KEYS = [
    "ueqsResults.obstructive-supportive",          # item 1
    "ueqsResults.complicated-easy",                # item 2
    "ueqsResults.inefficient-efficient",           # item 3
    "ueqsResults.confusing-clear",                 # item 4
    "ueqsResults.boring-exciting",                 # item 5
    "ueqsResults.not-interesting-interesting",     # item 6
    "ueqsResults.conventional-inventive",          # item 7
    "ueqsResults.usual-leading",                   # item 8
]

# Tasks go from index 0 to 10 (inclusive)
MIN_TASK_INDEX = 0
MAX_TASK_INDEX = 10

# Drop any sessions whose events are all before this cutoff
CUTOFF_MS = int(
    datetime(2025, 11, 4, 13, 30, tzinfo=timezone.utc).timestamp() * 1000
)


def iter_session_dirs(input_root: Path) -> List[Path]:
    """Return all session directories under input/."""
    return [p for p in input_root.iterdir() if p.is_dir()]


def load_session_events(session_dir: Path) -> List[Dict[str, Any]]:
    """Load all JSON logs for a single session as Python dicts."""
    events: List[Dict[str, Any]] = []
    for json_path in session_dir.glob("*.json"):
        with json_path.open("r", encoding="utf-8") as f:
            try:
                events.append(json.load(f))
            except json.JSONDecodeError:
                # Keep it simple: just skip broken files
                continue
    # Sort by timestamp if present, to make "session start" well-defined
    events.sort(key=lambda e: e.get("timestamp", 0))
    return events


def extract_cohort(events: List[Dict[str, Any]]) -> Optional[str]:
    """Get cohort parameter from the consent event URL, if present."""
    for ev in events:
        if ev.get("type") != "informedConsentCollected":
            continue
        page_url = ev.get("pageUrl")
        if not isinstance(page_url, str):
            continue
        parsed = urlparse(page_url)
        params = parse_qs(parsed.query)
        cohort_vals = params.get("cohort")
        if cohort_vals:
            return cohort_vals[0]
    return None


def extract_survey_data(events: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Extract UEQ-S and other questionnaire responses from survey_completion event."""
    survey_events = [ev for ev in events if ev.get("type") == "survey_completion"]
    if not survey_events:
        return {}

    # Use the last survey completion event in time
    survey_events.sort(key=lambda e: e.get("timestamp", 0))
    ev = survey_events[-1]

    out: Dict[str, Any] = {}

    # UEQ-S items (explicit list), converted from -3..3 to 1..7 scale
    for key in UEQ_KEYS:
        if key in ev:
            raw_val = ev[key]
            if isinstance(raw_val, (int, float)):
                out[key] = int(raw_val) + 4
            else:
                out[key] = None

    # Other questionnaire fields (e.g. experience, feedback)
    skip_keys = {
        "sessionId",
        "type",
        "timestamp",
        "lastConsentSessionId",
        *UEQ_KEYS,
    }
    for k, v in ev.items():
        if k in skip_keys:
            continue
        out[k] = v

    return out


def build_task_summary_for_session(
    session_name: str, events: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Build one wide-row with task times and skip reasons for a single session.

    Duration = time (in seconds) from first event of the session
    to the first 'task_fulfilled' with that taskIndex.
    """
    row: Dict[str, Any] = {}
    row["numLogs"] = len(events)
    row["cohort"] = extract_cohort(events)

    if not events:
        # Still create all task columns as None
        for ti in range(MIN_TASK_INDEX, MAX_TASK_INDEX + 1):
            row[f"task{ti}_time_s"] = None
            row[f"task{ti}_skipReason"] = None
        return row

    session_start = events[0].get("timestamp", 0)

    # For each task we want the *first* fulfillment or skip
    first_fulfilled: Dict[int, Dict[str, Any]] = {}
    first_skipped: Dict[int, Dict[str, Any]] = {}

    for ev in events:
        ev_type = ev.get("type")
        task_index = ev.get("taskIndex")
        if not isinstance(task_index, int):
            continue

        if ev_type == "task_fulfilled":
            if task_index not in first_fulfilled:
                first_fulfilled[task_index] = ev
        elif ev_type == "task_skipped":
            if task_index not in first_skipped:
                first_skipped[task_index] = ev

    # Per-task timings and skip reasons
    for ti in range(MIN_TASK_INDEX, MAX_TASK_INDEX + 1):
        fulfilled_ev = first_fulfilled.get(ti)
        skipped_ev = first_skipped.get(ti)

        # Duration (seconds) from session start to first fulfillment
        if fulfilled_ev is not None and "timestamp" in fulfilled_ev:
            dt_ms = fulfilled_ev["timestamp"] - session_start
            row[f"task{ti}_time_s"] = round(dt_ms / 1000.0, 2)
        else:
            row[f"task{ti}_time_s"] = None

        # Skip reason from first skip event, if any
        if skipped_ev is not None:
            row[f"task{ti}_skipReason"] = skipped_ev.get("skipReason")
        else:
            row[f"task{ti}_skipReason"] = None

    # Total completion time for the whole session (first to last event)
    last_ts = events[-1].get("timestamp", session_start)
    if isinstance(last_ts, (int, float)) and isinstance(session_start, (int, float)):
        row["total_time_s"] = round((last_ts - session_start) / 1000.0, 2)
    else:
        row["total_time_s"] = None

    # Total completed tasks (count distinct fulfilled task indices in our range)
    row["total_completed_tasks"] = sum(
        1
        for ti in range(MIN_TASK_INDEX, MAX_TASK_INDEX + 1)
        if ti in first_fulfilled
    )

    # Questionnaire data (UEQ-S and extra fields like experience, feedback)
    row.update(extract_survey_data(events))

    return row


def write_task_summary_xlsx(input_root: Path, output_path: Path) -> None:
    """Create a wide XLSX table with one row per session."""
    session_dirs = iter_session_dirs(input_root)
    rows: List[Dict[str, Any]] = []

    for session_dir in session_dirs:
        events = load_session_events(session_dir)
        if not events:
            continue

        # Skip sessions that are entirely before the cutoff timestamp
        if not any(
            isinstance(ev.get("timestamp"), (int, float))
            and ev["timestamp"] >= CUTOFF_MS
            for ev in events
        ):
            continue

        row = build_task_summary_for_session(session_dir.name, events)

        # Optionally drop specific cohorts (e.g. 'Testing' or missing/empty)
        cohort = row.get("cohort")
        if cohort is None or cohort == "" or cohort == "Testing":
            continue

        rows.append(row)

    if not rows:
        print("No sessions found. Nothing to write.")
        return

    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False)
    print(f"Wrote task summary XLSX to: {output_path}")


def main() -> None:
    project_root = Path(__file__).parent
    input_root = project_root / "input"
    output_root = project_root / "output"
    output_root.mkdir(parents=True, exist_ok=True)
    task_summary_xlsx = output_root / "task_summary.xlsx"

    if not input_root.exists():
        print(f"Input folder not found: {input_root}")
        return

    # First: build the wide per-session task table as XLSX
    write_task_summary_xlsx(input_root, task_summary_xlsx)


if __name__ == "__main__":
    main()


