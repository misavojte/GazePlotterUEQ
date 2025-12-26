"""
Generate a high-quality academic Dot Plot visualization for UEQ-S results.
Reads ueqs_sheet.xlsx and creates ueq_results_figure.pdf.
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as path_effects
import numpy as np

# Configuration
EXCEL_FILE = Path("ueqs_sheet.xlsx")
OUTPUT_FILE = Path("ueq_results_figure.pdf")
DPI = 600  # High resolution for line art (journal requirement)

# UEQ scale range
X_MIN = -3.0
X_MAX = 3.0

# Q1 Journal-Optimized Color Palette
# Colorblind-friendly sequential palette (ColorBrewer-inspired YlGnBu)
# Works excellently in both color and grayscale printing
# Optimized for high contrast while maintaining readability
BENCHMARK_COLORS = {
    "Bad": "#f0f9ff",           # Very light blue-white (ensures visibility)
    "Below Average": "#cce5f2", # Light blue-gray (colorblind safe)
    "Above Average": "#85c1e2", # Medium blue (clear distinction)
    "Good": "#4a90c2",          # Medium-dark blue (professional)
    "Excellent": "#2c5f8d",     # Dark blue (strong contrast, readable)
}

# Annotation font styling (used for threshold labels and legend titles)
ANNOTATION_FONT = {
    "fontsize": 9.5,
    "fontstyle": "italic",
    "color": "#333333",
}


def load_confidence_intervals_data(excel_path: Path) -> dict:
    """
    Extract Mean, Confidence, and N for Pragmatic Quality, Hedonic Quality, and Overall
    from the Confidence_Intervals sheet.
    
    Returns a dictionary with keys: 'Pragmatic Quality', 'Hedonic Quality', 'Overall'
    Each value is a dict with 'mean', 'confidence', 'n'
    """
    # Read the Confidence_Intervals sheet
    df = pd.read_excel(excel_path, sheet_name="Confidence_Intervals", header=None)
    
    # Based on the Excel structure:
    # Row 2 (index 2): Contains "Confidence intervals (p=0.05) per scale" in column 8
    # Row 3 (index 3): Contains column headers: Scale (col 8), Mean (col 9), Std. Dev. (col 10), 
    #                  N (col 11), Confidence (col 12), Confidence interval (col 13)
    # Row 4+ (index 4+): Contains the actual data
    
    # Column indices for the scales section (starts at column 8)
    SCALE_COL = 8
    MEAN_COL = 9
    STD_DEV_COL = 10
    N_COL = 11
    CONFIDENCE_COL = 12
    
    # Data starts at row 4 (index 4)
    data_start_row = 4
    
    # Extract data for the three scales we need
    results = {}
    target_scales = ["Pragmatic Quality", "Hedonic Quality", "Overall"]
    
    for row_idx in range(data_start_row, min(data_start_row + 10, df.shape[0])):
        scale_cell = df.iloc[row_idx, SCALE_COL]
        
        # Skip if cell is NaN or empty
        if pd.isna(scale_cell):
            continue
        
        scale_name = str(scale_cell).strip()
        
        if scale_name in target_scales:
            try:
                mean_val = float(df.iloc[row_idx, MEAN_COL])
                n_val = int(float(df.iloc[row_idx, N_COL]))
                confidence_val = float(df.iloc[row_idx, CONFIDENCE_COL])
                
                results[scale_name] = {
                    'mean': mean_val,
                    'confidence': confidence_val,
                    'n': n_val
                }
            except (ValueError, TypeError) as e:
                print(f"Warning: Could not parse row for {scale_name}: {e}")
                continue
    
    # Verify we got all three scales
    missing = set(target_scales) - set(results.keys())
    if missing:
        raise ValueError(f"Missing data for scales: {missing}")
    
    return results


def load_benchmark_thresholds(excel_path: Path) -> dict:
    """
    Load benchmark category boundaries for each scale from the Benchmark sheet.

    The Benchmark sheet encodes quantile-based borders in the table labelled
    'Benchmark borders (purely technical, please ignore)'.

    We interpret them as:
        Bad          : [X_MIN, q25]
        Below Average: (q25, q50]
        Above Average: (q50, q75]
        Good         : (q75, q90]
        Excellent    : (q90, X_MAX]

    Returns:
        Dict[str, Dict[str, tuple]] mapping normalized scale name to
        category -> (low, high) interval.
    """
    df = pd.read_excel(excel_path, sheet_name="Benchmark", header=None)

    # Find the row that contains the 'Benchmark borders' label
    anchor_row = None
    for idx, row in df.iterrows():
        for val in row.values:
            if isinstance(val, str) and "Benchmark borders" in val:
                anchor_row = idx
                break
        if anchor_row is not None:
            break

    if anchor_row is None:
        raise ValueError("Could not find 'Benchmark borders' section in Benchmark sheet.")

    header_row = anchor_row + 1  # row with: Scale, 0.25, 0.5, 0.75, 0.9
    first_data_row = header_row + 1

    # Column indices for quantile borders
    SCALE_COL = 0
    Q25_COL = 1
    Q50_COL = 2
    Q75_COL = 3
    Q90_COL = 4

    # Map short names in Benchmark sheet to the full names we use elsewhere
    name_map = {
        "Pragmatic Q.": "Pragmatic Quality",
        "Hedonic Q.": "Hedonic Quality",
        "Overall": "Overall",
    }

    thresholds: dict = {}

    for row_idx in range(first_data_row, first_data_row + 10):
        if row_idx >= df.shape[0]:
            break

        raw_name = df.iloc[row_idx, SCALE_COL]
        if pd.isna(raw_name):
            continue

        raw_name = str(raw_name).strip()
        if raw_name not in name_map:
            continue

        norm_name = name_map[raw_name]

        try:
            q25 = float(df.iloc[row_idx, Q25_COL])
            q50 = float(df.iloc[row_idx, Q50_COL])
            q75 = float(df.iloc[row_idx, Q75_COL])
            q90 = float(df.iloc[row_idx, Q90_COL])
        except (TypeError, ValueError):
            raise ValueError(f"Could not parse benchmark borders for scale '{raw_name}'")

        thresholds[norm_name] = {
            "Bad": (X_MIN, q25),
            "Below Average": (q25, q50),
            "Above Average": (q50, q75),
            "Good": (q75, q90),
            "Excellent": (q90, X_MAX),
        }

    # Ensure we got thresholds for all three scales we plot
    missing = {s for s in ["Pragmatic Quality", "Hedonic Quality", "Overall"]} - set(
        thresholds.keys()
    )
    if missing:
        raise ValueError(f"Missing benchmark thresholds for scales: {missing}")

    return thresholds


def create_dot_plot(data: dict, benchmark_thresholds: dict, output_path: Path, dpi: int = 300):
    """
    Create a horizontal dot plot with error bars and color-coded background zones based on benchmarks.
    Optimized for Q1 academic journal publication standards.
    
    Args:
        data: Dictionary with scale names as keys and dicts with 'mean', 'confidence', 'n' as values
        output_path: Path to save the figure
        dpi: Resolution for the output figure
    """
    # Set publication-quality font settings BEFORE creating figure
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Arial", "Helvetica", "DejaVu Sans", "Liberation Sans"]
    plt.rcParams["font.size"] = 10
    plt.rcParams["axes.labelsize"] = 11
    plt.rcParams["axes.titlesize"] = 11
    plt.rcParams["xtick.labelsize"] = 10
    plt.rcParams["ytick.labelsize"] = 10.5
    plt.rcParams["legend.fontsize"] = 9
    plt.rcParams["figure.dpi"] = dpi
    
    # Extract data in the order we want to display (top to bottom: Pragmatic, Hedonic, Overall)
    scales = ["Pragmatic Quality", "Hedonic Quality", "Overall"]
    # Reverse for display so Pragmatic is at top (y=2), Overall at bottom (y=0)
    scales = scales[::-1]
    means = [data[scale]["mean"] for scale in scales]
    confidences = [data[scale]["confidence"] for scale in scales]
    # Get sample size (n) - should be the same for all scales
    n = data["Overall"]["n"]

    # Create figure with publication-standard dimensions (single column: ~3.5", double: ~7")
    # Using wider format suitable for modern journal layouts
    fig, ax = plt.subplots(figsize=(8.5, 4.5))

    # Draw per-scale background zones using thresholds from Benchmark sheet
    zone_order = ["Bad", "Below Average", "Above Average", "Good", "Excellent"]
    bar_height = 0.8  # height of the background band per scale

    for i, scale in enumerate(scales):
        y_bottom = i - bar_height / 2
        zones = benchmark_thresholds[scale]

        for zone_name in zone_order:
            x_start, x_end = zones[zone_name]
            color = BENCHMARK_COLORS[zone_name]
            rect = mpatches.Rectangle(
                (x_start, y_bottom),
                x_end - x_start,
                bar_height,
                facecolor=color,
                edgecolor="none",
                alpha=0.7,  # Slightly more opaque for better visibility in print
                zorder=0,
            )
            ax.add_patch(rect)

    # Y positions for the three scales
    y_positions = np.arange(len(scales))

    # Create horizontal dot plot with error bars (publication-quality styling)
    ax.errorbar(
        means,
        y_positions,
        xerr=confidences,
        fmt="o",
        capsize=5,           # Slightly larger caps for clarity
        capthick=1.8,        # Thicker caps for visibility
        markersize=10,       # Larger markers for better visibility
        color="#1a1a1a",     # Near-black for crisp printing
        markerfacecolor="white",
        markeredgewidth=2.0, # Clear marker edges
        linewidth=2.0,       # Clear error bars
        zorder=3,
        elinewidth=2.0,      # Explicit error bar line width
    )

    # Add labels just above each dot with ± symbol showing error margins
    # Using halo effect (white outline) for publication-quality readability
    for i, (mean, conf) in enumerate(zip(means, confidences)):
        # Position label slightly higher to avoid collision with data points
        label_y = i + 0.14
        label_text = f"{mean:.2f} ± {conf:.2f}"
        ax.text(
            mean,
            label_y,
            label_text,
            va="bottom",
            ha="center",
            fontsize=10,           # Slightly larger for print clarity
            fontweight="medium",   # Medium weight for better visibility
            color="#1a1a1a",       # Near-black for crisp printing
            path_effects=[
                path_effects.withStroke(linewidth=1.8, foreground="white", alpha=0.85)
            ],
            zorder=4,
        )
    
    # Set y-axis labels with multi-line formatting (reversed to match scales order)
    y_labels = ["Pragmatic\nQuality", "Hedonic\nQuality", "Overall"]
    y_labels = y_labels[::-1]  # Reverse to match reversed scales
    ax.set_yticks(y_positions)
    ax.set_yticklabels(y_labels, fontsize=10.5)

    # Set x-axis range and labels
    ax.set_xlim(X_MIN, X_MAX)
    ax.set_xlabel("UEQ-S Score", fontsize=11, fontweight="normal")
    
    # Remove top and right spines for academic style
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # Publication-quality axis lines and tick styling
    ax.spines["left"].set_linewidth(1.2)
    ax.spines["bottom"].set_linewidth(1.2)
    ax.spines["left"].set_color("#333333")
    ax.spines["bottom"].set_color("#333333")
    ax.tick_params(
        axis="both",
        which="both",
        direction="out",
        length=5,           # Slightly longer ticks for clarity
        width=1.2,          # Thicker ticks for print
        labelsize=10,
        color="#333333",
    )

    # Add vertical dashed line at threshold of positive evaluation (0.8)
    threshold = 0.8
    ax.axvline(
        x=threshold,
        color="#333333",      # Dark gray for better visibility
        linestyle="--",
        linewidth=1.2,        # Slightly thicker for print
        alpha=0.5,            # More visible
        zorder=1,
        dashes=(5, 3),        # Explicit dash pattern for consistency
    )

    # Add label above the threshold line
    top_y = 0.98  # Positioned just above the top of the plot area, close to the dashed line
    ax.text(
        threshold,
        top_y,
        "Positive Evaluation (> 0.8)",
        ha="center",
        va="bottom",
        **ANNOTATION_FONT,      # Use shared annotation font styling
        transform=ax.get_xaxis_transform(),
        clip_on=False,
    )

    # Add subtle grid for better readability (publication-quality styling)
    ax.grid(True, axis="x", linestyle=":", alpha=0.3, linewidth=0.8, zorder=0, color="#cccccc")
    ax.set_axisbelow(True)
    
    # FIX: Remove the old legend code and use this "Ribbon" style
    
    # 1. Create the legend handles (colored rectangles)
    handles = []
    labels = ["Bad", "Below Average", "Above Average", "Good", "Excellent"]
    
    for label in labels:
        # Use subtle border for lighter colors, darker border for darker colors
        edge_color = '#b0b0b0' if label in ["Bad", "Below Average"] else '#808080'
        handles.append(mpatches.Patch(
            facecolor=BENCHMARK_COLORS[label], 
            edgecolor=edge_color,  # Adaptive border for optimal contrast
            linewidth=0.6,          # Slightly thicker for print clarity
            label=label
        ))

    # 2. Place it perfectly below the X-axis title
    # bbox_to_anchor=(0.5, -0.25) moves it down to clear the axis labels
    legend = ax.legend(
        handles, 
        labels, 
        loc='upper center', 
        bbox_to_anchor=(0.5, -0.25), 
        ncol=5,                 # Force all items into 1 row
        frameon=False,          # No box border (looks cleaner)
        fontsize=9,             # Readable text size
        handlelength=1.5,       # Wider color swatches (looks more like a key)
        handleheight=0.8,
        columnspacing=1.5,      # Space between items
        title="UEQ-S Benchmark Dataset",
        title_fontsize=ANNOTATION_FONT["fontsize"]
    )
    # Apply shared annotation font styling to legend title
    title = legend.get_title()
    title.set_style(ANNOTATION_FONT["fontstyle"])
    title.set_color(ANNOTATION_FONT["color"])
    
    # CRITICAL: Adjust layout padding so the legend isn't cut off
    # The 'rect' parameter reserves space at the bottom (0.15 = 15% height)
    plt.tight_layout(rect=[0, 0.15, 1, 1])
    
    # Save figure at publication-quality resolution
    # PDF format is preferred for journal submissions (vector graphics)
    fig.savefig(
        output_path, 
        dpi=dpi, 
        bbox_inches="tight", 
        pad_inches=0.1,      # Small padding for professional appearance
        facecolor="white",
        edgecolor="none",
        format="pdf",
        transparent=False,
    )
    print(f"Figure saved to: {output_path}")

    plt.close()


def main():
    """Main function to generate the UEQ figure."""
    excel_path = Path(EXCEL_FILE)
    
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    
    print(f"Loading data from {excel_path}...")
    data = load_confidence_intervals_data(excel_path)
    benchmark_thresholds = load_benchmark_thresholds(excel_path)
    
    print("Extracted data:")
    for scale, values in data.items():
        print(f"  {scale}: Mean={values['mean']:.3f}, Confidence={values['confidence']:.3f}, N={values['n']}")
    
    print(f"\nGenerating dot plot...")
    create_dot_plot(data, benchmark_thresholds, OUTPUT_FILE, dpi=DPI)
    
    print("Done!")


if __name__ == "__main__":
    main()

