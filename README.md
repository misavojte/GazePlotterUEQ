# GazePlotterUEQ

GazePlotterUEQ is a specialized toolset designed to process, analyze, and visualize User Experience Questionnaire (Short Version - UEQ-S) data. It was developed to evaluate the user experience of GazePlotter, an eye-tracking visualization tool, by converting raw interaction logs into publication-ready academic figures.

## Key Features

* Automated Log Processing: Converts raw JSON session logs into a structured Excel summary including task timings, skip reasons, and questionnaire responses.
* UEQ-S Statistical Analysis: Utilizes a comprehensive analysis framework to calculate scale means for Pragmatic Quality, Hedonic Quality, and Overall user experience.
* Academic-Grade Visualization: Generates high-resolution (600 DPI), colorblind-friendly dot plots with confidence intervals and benchmark background zones, optimized for Q1 journal requirements.
* Benchmarking: Automatically compares results against a benchmark dataset of 21,175 persons from 468 studies.
* Data Validation: Includes heuristics to detect suspicious or inconsistent data patterns, such as random answering or middle-category bias.

## Project Structure

| File | Description |
| :--- | :--- |
| main.py | The entry point for data extraction. Processes raw JSON logs from the /input folder and generates task_summary.xlsx. |
| generate_ueq_figure.py | Python script that reads analyzed Excel data to create the final academic visualization. |
| ueqs_sheet.xlsx | The core analytical engine (represented via CSV exports) that handles all UEQ-S calculations, confidence intervals, and benchmarking. |
| feedback_coded.xlsx | Contains qualitatively coded user feedback, categorizing open-ended responses into themes like Visuals, Speed, and Data Context. |
| requirements.txt | Lists necessary Python dependencies like numpy, pandas, and matplotlib. |

## Installation and Setup

1. Clone the repository and navigate to the project root.
2. Install dependencies: `pip install -r requirements.txt`.
3. Prepare Input: Place raw JSON session logs into an /input directory.

## Workflow

1. Extract Data: Run `python main.py` to aggregate raw logs. This script maps interaction data and converts UEQ-S responses from the logged -3 to +3 scale to the 1 to 7 scale required for the analysis sheet.
2. Analyze Statistics: Import the resulting data into `ueqs_sheet.xlsx`. The sheet automatically calculates scale consistency (Cronbach's Alpha) and confidence intervals (p=0.05).
3. Generate Figures: Run `python generate_ueq_figure.py` to produce `ueq_results_figure.pdf`. This visualization provides an immediate quantified look at the tool's performance relative to industry standards.

## Data Quantification

The project quantifies User Experience across dimensions using a scale from -3 (extremely bad) to +3 (extremely good):

* Pragmatic Quality: Measures efficiency, perspicuity, and dependability (Task-oriented).
* Hedonic Quality: Measures stimulation and novelty (Non-task oriented).
* Overall: Represents the total combined score of the user experience.

Results are categorized into benchmark tiers: Bad, Below Average, Above Average, Good, and Excellent. For example, current results show an Excellent rating for Hedonic Quality (1.77 +/- 0.28) and an Above Average rating for Pragmatic Quality (1.39 +/- 0.35).