# Charlotte Grosjean - 22.04.2026 - UNIL Master Thesis

"""
OVERLAP PROCESSING PIPELINE
This script automates the processing of overlapping periods between different data sources for pluviometer stations. It includes functions to:
- Load and clean station data files
- Identify overlapping time periods between sources
- Save overlapping data for further analysis
- Align overlapping data series using various methods (exact, tolerance-based, interpolation)
- Compute comparison metrics (mean difference, RMSE, correlation)
- Generate visualizations of aligned series 

Main functions are included in overlap_utils.py. The script is structured in three main sections:
1. Identify overlaps: Scans processed station files, identifies overlapping periods, and saves the overlapping data to CSV files.
2. Compare overlaps: Loads the overlapping data, performs alignments, computes metrics, and saves results to a summary CSV file.
3. Visualize results: Generates line plots and comparative scatter plots for the overlapping periods.
"""


#%% Imports and configuration

import os
import pandas as pd
import numpy as np
from glob import glob
from collections import defaultdict
import re
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.io as pio
from scipy import signal
from A_data_processing_utils import *
from C_overlap_utils import *  # helper functions for identify_overlaps.py and compare_overlaps.py
import E_plot_config as cfg  # centralized plot styling configuration
from I_paths_config import PROCESSED_DIRECTORY, OVERLAPS_DIRECTORY

#%% 1 - identify overlaps

base_directory_processed = PROCESSED_DIRECTORY
output_directory = OVERLAPS_DIRECTORY
os.makedirs(output_directory, exist_ok=True)

# Run overlap identification for all stations
def process_overlaps(directory):
    """Process overlaps for a station.
    1. Load all source files
    2. Find overlapping periods
    3. Save overlapping data for each source
    """
    station_name, sources = load_station_files(directory)

    if len(sources) < 2:
        print(f"Station {station_name}: Not enough sources to check overlaps.")
        return

    print(f"\nProcessing overlaps for station: {station_name}")

    overlaps = find_overlaps(sources)

    if not overlaps:
        print(f"No overlaps found for station {station_name}.")
        return

    # Create station subfolder
    station_output_dir = os.path.join(output_directory, station_name)
    os.makedirs(station_output_dir, exist_ok=True)

    for source1, source2, start_overlap, end_overlap in overlaps:
        df1 = next(df for src, df in sources if src == source1)
        df2 = next(df for src, df in sources if src == source2)
        save_overlap_data(station_name, source1, df1, source2, df2, start_overlap, end_overlap, station_output_dir)

# Run all stations
if __name__ == "__main__":
    for folder_name in os.listdir(base_directory_processed):
        folder_path = os.path.join(base_directory_processed, folder_name)
        if os.path.isdir(folder_path):
            process_overlaps(folder_path)

    print("Overlap identification completed.")

#%% 2 - compare overlaps

# Temporal alignment
ALIGNMENT_TOLERANCE_MINUTES = 1.5
LAG_RANGE = range(-60, 61)  # ±60 minutes

# Validation thresholds
MIN_OVERLAP_DURATION_HOURS = 2    # Minimum overlap period

# Base folder containing station subfolders and overlap CSV files
BASE_OVERLAP_DIR = OVERLAPS_DIRECTORY
# Output summary file that aggregates comparison metrics
OUTPUT_SUMMARY = os.path.join(BASE_OVERLAP_DIR, "overlap_comparison_summary.csv")
# Use a flexible filename pattern to detect overlap groups from file names
FILE_PATTERN = re.compile(r"_overlap_([0-9_]+)\.csv$")

# Process overlaps for all stations and save comparison metrics to a summary CSV
def process_overlap_group(file_paths, station_name, overlap_key):
    """Process one overlap group and return comparison metrics."""
    if len(file_paths) != 2:
        print(f"    -> Skipped: {len(file_paths)} files (need exactly 2)")
        return None

    try:
        df1 = load_and_prepare_df(file_paths[0])
        df2 = load_and_prepare_df(file_paths[1])

        # Print file sizes to help diagnose missing overlap
        print(f"    -> File 1: {os.path.basename(file_paths[0])} - {len(df1)} rows")
        print(f"    -> File 2: {os.path.basename(file_paths[1])} - {len(df2)} rows")

        if df1.empty:
            print(f"    -> Failed: First file is empty or invalid")
            return None
        if df2.empty:
            print(f"    -> Failed: Second file is empty or invalid")
            return None

        MIN_POINTS_ALIGNED = min(len(df1), len(df2))

        source_1_name = os.path.basename(file_paths[0]).replace('.csv', '')
        source_2_name = os.path.basename(file_paths[1]).replace('.csv', '')
        source_1_name = source_1_name.split('_overlap')[0]
        source_2_name = source_2_name.split('_overlap')[0]

        # Test exact and tolerance-based alignments and choose the one with more aligned points.
        exact_aligned = align_data(df1, df2, method='exact')
        tolerance_aligned = align_data(df1, df2, method='tolerance', tolerance_minutes=ALIGNMENT_TOLERANCE_MINUTES)
        print(f"    -> Exact alignment: {len(exact_aligned)} common points")
        print(f"    -> Tolerance alignment: {len(tolerance_aligned)} points")

        if len(tolerance_aligned) > len(exact_aligned):
            aligned = tolerance_aligned
            alignment_method = f'tolerance (±{ALIGNMENT_TOLERANCE_MINUTES} min)'
        else:
            aligned = exact_aligned
            alignment_method = 'exact'

        if len(aligned) < MIN_POINTS_ALIGNED:
            print(f"    -> Trying nearest neighbor alignment...")
            nearest_aligned = align_data(df1, df2, method='nearest')
            print(f"    -> Nearest alignment: {len(nearest_aligned)} points")
            if len(nearest_aligned) > len(aligned):
                aligned = nearest_aligned
                alignment_method = f'nearest neighbor 3min grid)'

        if aligned.empty:
            print(f"    -> Failed: No temporal overlap between files")
            return None

        metrics = compute_metrics(aligned)
        print(f"    -> Success: {metrics['n_points']} points compared")

        # Create interactive aligned series plot whenever aligned data exists.
        interactive_dir = os.path.join(os.path.dirname(file_paths[0]), 'interactive_plots')
        os.makedirs(interactive_dir, exist_ok=True)
        try:
            aligned_html_path = plot_aligned_series_interactive(aligned, station_name, overlap_key, interactive_dir, source_1_name, source_2_name)
            print(f"    -> Aligned series HTML interactive saved to: {os.path.basename(aligned_html_path)}")
        except Exception as e:
            print(f"    -> Warning: aligned series HTML interactive not saved: {e}")

        # Compute cross-correlation
        lags_hours, corr_values, best_lag_info = compute_cross_correlation(aligned, max_lag_hours=24)
        if lags_hours is not None:
            best_lag_hours, best_corr = best_lag_info
            print(f"    -> Cross-correlation: best lag = {best_lag_hours:.2f}h (corr = {best_corr:.3f})")

            # Create cross-correlation plots
            cross_corr_dir = os.path.join(os.path.dirname(file_paths[0]), 'cross_correlation_plots')
            os.makedirs(cross_corr_dir, exist_ok=True)
            plot_path = plot_cross_correlation(lags_hours, corr_values, best_lag_info, station_name, overlap_key, source_1_name, source_2_name, cross_corr_dir)
            print(f"    -> Cross-correlation PNG saved to: {os.path.basename(plot_path)}")
            interactive_cross_path = plot_cross_correlation_interactive(lags_hours, corr_values, best_lag_info, station_name, overlap_key, source_1_name, source_2_name,cross_corr_dir)
            print(f"    -> Cross-correlation HTML interactive saved to: {os.path.basename(interactive_cross_path)}")
        else:
            best_lag_hours, best_corr = np.nan, np.nan
            print(f"    -> Cross-correlation: not enough data points")

        # Save aligned data to a CSV file with metadata header
        cross_corr_info = (best_lag_hours, best_corr) if lags_hours is not None else None
        aligned_output_path = save_aligned_data(aligned, station_name, overlap_key, file_paths, os.path.dirname(file_paths[0]), alignment_method, metrics, cross_corr_info)
        print(f"    -> Aligned data saved to: {os.path.basename(aligned_output_path)}")

        return {
            'station': station_name,
            'overlap_period': overlap_key,
            'source_1': source_1_name,
            'source_2': source_2_name,
            'n_rows_file_1': len(df1),
            'n_rows_file_2': len(df2),
            'alignment_method': alignment_method,
            **metrics,
            'best_lag_hours': best_lag_hours,
            'best_correlation': best_corr
        }

    except Exception as e:
        print(f"    -> Error: {e}")
        return None

# Main processing function to loop over all stations and overlap groups
def main():
    """Main processing function."""
    if not os.path.isdir(BASE_OVERLAP_DIR):
        raise FileNotFoundError(f"Directory not found: {BASE_OVERLAP_DIR}")

    results = []

    # Loop over each station folder in the overlap directory
    station_dirs = [d for d in glob(os.path.join(BASE_OVERLAP_DIR, "*")) if os.path.isdir(d)]
    for station_dir in sorted(station_dirs):
        station_name = os.path.basename(station_dir)
        print(f"\nProcessing station: {station_name}")

        # Group files by the overlap key extracted from file names
        groups = defaultdict(list)
        csv_files = glob(os.path.join(station_dir, "*.csv"))
        print(f"Found {len(csv_files)} CSV files")

        for file_path in csv_files:
            filename = os.path.basename(file_path)
            key = get_overlap_key(filename)
            if key:
                groups[key].append(file_path)
                print(f"  Grouped {filename} -> key: {key}")
            else:
                print(f"  Skipped {filename} (no overlap key found)")

        print(f"Found {len(groups)} overlap groups")

        # Process each overlap group and collect metrics
        for overlap_key, file_paths in groups.items():
            print(f"  Processing group {overlap_key} with {len(file_paths)} files")
            result = process_overlap_group(file_paths, station_name, overlap_key)
            if result:
                results.append(result)
                print(f"    -> Success: {result['n_points']} points compared")
            else:
                print(f"    -> Failed or no valid data")

    # Save a CSV summary of all comparisons
    if results:
        df_summary = pd.DataFrame(results)
        df_summary.to_csv(OUTPUT_SUMMARY, index=False)
        print(f"\nComparison summary saved to: {OUTPUT_SUMMARY}")
        print(f"Processed {len(results)} overlap comparisons")
    else:
        print("\nNo valid overlap comparisons found")


if __name__ == "__main__":
    main()

#%% 3 - Visualize results overlaps 

# Configuration
BASE_OVERLAP_DIR = OVERLAPS_DIRECTORY
PLOT_FILENAME_SUFFIX = "_overlap_plot.png"
MODE = "single"  # Set to "single" or "all"

# Optional high-quality static zoom export for report figures.
# Set GENERATE_REPORT_ZOOM to True and choose the time window to export.
GENERATE_REPORT_ZOOM = True
REPORT_ZOOM_START = "2026-03-14 00:00"
REPORT_ZOOM_END = "2026-03-15 00:00"
REPORT_ZOOM_FORMATS = ("pdf", "png")
REPORT_ZOOM_DPI = 600

# For single mode, set the two CSV files to compare
SINGLE_FILE_1 = '/Users/charlottegrosjean/Library/Mobile Documents/com~apple~CloudDocs/2UNIL/Master/Master thesis/3_Data/2_Overlaps/BETH/BETH_20260401_txt_overlap_20260305_20260401.csv'
SINGLE_FILE_2 = '/Users/charlottegrosjean/Library/Mobile Documents/com~apple~CloudDocs/2UNIL/Master/Master thesis/3_Data/2_Overlaps/BETH/BETH_20260402_csv_overlap_20260305_20260401.csv'
SAVE_SINGLE_PLOT = True
SHOW_SINGLE_PLOT = True

# Pattern used to group overlap files by overlap period
FILE_PATTERN = re.compile(r"_overlap_([0-9_]+)\.csv$")

def process_station(station_dir):
    station_name = os.path.basename(station_dir)
    groups = build_groups(station_dir)                              # function defined above to group files by overlap key
    if not groups:
        print(f"Station {station_name}: no overlap files found.")
        return

    plots_dir = os.path.join(station_dir, "plots")
    os.makedirs(plots_dir, exist_ok=True)

    for overlap_key, file_paths in groups.items():
        if len(file_paths) < 2:
            print(f"Station {station_name} - period {overlap_key}: less than 2 files, cannot plot.")
            continue

        # Standard line plot
        output_path_line = os.path.join(plots_dir, f"{station_name}_overlap_{overlap_key}{PLOT_FILENAME_SUFFIX}")
        try:
            plot_group(file_paths, output_path_line, station_name, overlap_key)
            print(f"Generated line plot: {output_path_line}")
        except Exception as exc:
            print(f"Error station {station_name}, overlap {overlap_key} (line): {exc}")

        if GENERATE_REPORT_ZOOM:
            report_dir = os.path.join(plots_dir, "report_exports")
            try:
                report_paths = plot_group_report_zoom(
                    file_paths,
                    report_dir,
                    station_name,
                    overlap_key,
                    REPORT_ZOOM_START,
                    REPORT_ZOOM_END,
                    formats=REPORT_ZOOM_FORMATS,
                    dpi=REPORT_ZOOM_DPI
                )
                print(f"Generated report zoom exports: {', '.join(report_paths)}")
            except Exception as exc:
                print(f"Info: No report zoom generated for {station_name}, overlap {overlap_key}: {exc}")
        
        # Comparative scatter plot (Using aligned data generated by compare_overlaps.py script)
        aligned_dir = os.path.join(station_dir, "aligned_data")
        aligned_file = os.path.join(aligned_dir, f"{station_name}_overlap_{overlap_key}_aligned.csv")
        if os.path.exists(aligned_file):
            output_path_comparison = os.path.join(plots_dir, f"{station_name}_overlap_{overlap_key}_comparison{PLOT_FILENAME_SUFFIX}")
            try:
                plot_group_scatter_comparison(aligned_file, output_path_comparison, station_name, overlap_key)
                print(f"Generated comparative scatter plot: {output_path_comparison}")
            except Exception as exc:
                print(f"Error station {station_name}, overlap {overlap_key} (comparison): {exc}")
        else:
            print(f"Info: No aligned data found for comparative scatter plot")

# Execution 

def main():
    if MODE == 'single':
        if not os.path.exists(SINGLE_FILE_1) or not os.path.exists(SINGLE_FILE_2):
            raise FileNotFoundError('One or both files specified for single mode are not found.')
        
        visualize_single_overlap(SINGLE_FILE_1, SINGLE_FILE_2, save_plot=SAVE_SINGLE_PLOT, show_plot=SHOW_SINGLE_PLOT)
        
        # Try to find and use aligned data for comparative scatter plot
        station_dir = os.path.dirname(SINGLE_FILE_1)
        station_name = os.path.basename(station_dir)
        
        # Extract overlap key from the file name
        overlap_key = get_overlap_key(os.path.basename(SINGLE_FILE_1))

        if GENERATE_REPORT_ZOOM:
            report_dir = os.path.join(station_dir, "plots", "report_exports")
            try:
                report_paths = plot_group_report_zoom(
                    [SINGLE_FILE_1, SINGLE_FILE_2],
                    report_dir,
                    station_name,
                    overlap_key or "single",
                    REPORT_ZOOM_START,
                    REPORT_ZOOM_END,
                    formats=REPORT_ZOOM_FORMATS,
                    dpi=REPORT_ZOOM_DPI
                )
                print(f"Generated report zoom exports: {', '.join(report_paths)}")
            except Exception as exc:
                print(f"Info: No report zoom generated for single mode: {exc}")
        
        if overlap_key:
            aligned_file = os.path.join(station_dir, 'aligned_data', f"{station_name}_overlap_{overlap_key}_aligned.csv")
            visualize_single_overlap_scatter_comparison(SINGLE_FILE_1, SINGLE_FILE_2, 
                                                      save_plot=SAVE_SINGLE_PLOT, show_plot=SHOW_SINGLE_PLOT, 
                                                      aligned_file_path=aligned_file)
        else:
            # If no aligned file is found, still create comparison plot from direct alignment
            visualize_single_overlap_scatter_comparison(SINGLE_FILE_1, SINGLE_FILE_2, 
                                                      save_plot=SAVE_SINGLE_PLOT, show_plot=SHOW_SINGLE_PLOT, 
                                                      aligned_file_path=None)
        return

    if MODE == 'all':
        if not os.path.isdir(BASE_OVERLAP_DIR):
            raise FileNotFoundError(f"Directory not found: {BASE_OVERLAP_DIR}")

        station_dirs = [d for d in glob(os.path.join(BASE_OVERLAP_DIR, "*")) if os.path.isdir(d)]
        for station_dir in sorted(station_dirs):
            process_station(station_dir)
        return

    raise ValueError("Invalid MODE: choose 'single' or 'all'.")

if __name__ == '__main__':
    main()

# %%
