# Charlotte Grosjean - 26.05.2026 - UNIL Master Thesis
"""
STATION GRIDS: AVAILABILITY + PRECIPITATION

This script builds and exports station data grids from consolidated measurement files.

Process:
1. Load all consolidated station CSV files from 2_Consolidated_Stations
2. Build a shared 3-minute time index spanning all stations and periods
3. Distribute station measurements to the shared grid (accounting for variable source resolution)
4. Generate availability grid (binary: data present or gap)
5. Generate precipitation grid (accumulated mm per 3-minute bucket)
6. Generate gap-code grid (marks -1=inter-file gaps, -2=intra-file gaps)
7. Export to parquet/CSV and generate interactive heatmaps

Outputs:
- 4_Grids/4.1_Availability: binary availability grid, per-station summary, daily heatmap
- 4_Grids/4.2_Precipitation: precipitation grid, gap codes, MeteoSwiss data, daily/30-min/3-min heatmaps
"""

from __future__ import annotations

import os

import pandas as pd

from G_station_grids_utils import (
    load_consolidated_station_data,
    build_station_grids,
    save_availability_outputs,
    build_daily_availability_heatmap,
    save_heatmap_html,
    save_rain_outputs,
    build_daily_rain_heatmap,
    build_30min_heatmap,
    add_meteosuisse_to_grid,
    save_heatmap_png,
    save_qc_long_outputs,
    save_monthly_qc_excel,
    save_period_summaries,
)
from I_paths_config import (
    CONSOLIDATED_DIRECTORY,
    AVAILABILITY_DIRECTORY,
    PRECIPITATION_DIRECTORY,
    GRID_STEP_MINUTES,
    DEFAULT_SOURCE_STEP_MINUTES,
    METEOSWISS_STEP_MINUTES,
    RESAMPLE_LABEL,
    RESAMPLE_CLOSED,
    ensure_output_directories,
)


# %% 0 - Configuration parameters

# Input/output directories
BASE_DIRECTORY_CONSOLIDATED = CONSOLIDATED_DIRECTORY
OUTPUT_DIRECTORY_AVAILABILITY = AVAILABILITY_DIRECTORY
OUTPUT_DIRECTORY_RAIN = PRECIPITATION_DIRECTORY

# Time resolution parameters (3-minute grid throughout)
# - Pluvimate data: 3-minute resolution
# - MeteoSwiss data: 10-minute resolution (redistributed to 3-minute grid)
# GRID_STEP_MINUTES, DEFAULT_SOURCE_STEP_MINUTES, METEOSWISS_STEP_MINUTES imported above

# Time-binning convention: right-labelled, right-closed
# - Timestamp 12:03 represents the interval (12:00, 12:03]
# - Each precipitation value is an accumulation ending at that timestamp
# RESAMPLE_LABEL, RESAMPLE_CLOSED imported above
TIME_ALIGNMENT_MESSAGE = (
    f"[INFO] Time alignment: {GRID_STEP_MINUTES}-minute grid, "
    f"source data default={DEFAULT_SOURCE_STEP_MINUTES} min, "
    f"MeteoSwiss={METEOSWISS_STEP_MINUTES} min, "
    f"label='{RESAMPLE_LABEL}', closed='{RESAMPLE_CLOSED}'."
)

# Ensure output directories exist
ensure_output_directories()


# %% 1 - Main execution: Availability grid

def run_availability_grid(availability_df: pd.DataFrame | None) -> None:
    """
    Build and export the binary availability grid.

    Output directory: 3_Data/4_Grids/4.1_Availability
    """
    print("\n" + "=" * 80)
    print("SECTION 1: AVAILABILITY GRID")
    print("=" * 80)

    if availability_df is None:
        print("[ERROR] No data available to build availability grid.")
        return

    print("[INFO] Exporting availability grid and summaries...")
    parquet_path, excel_path = save_availability_outputs(availability_df)
    save_period_summaries(availability_df, grid_type="availability")

    print("[INFO] Generating daily availability heatmap...")
    daily_availability_df = build_daily_availability_heatmap(availability_df)
    html_path = save_heatmap_html(
        daily_availability_df,
        is_daily=True,
        resolution_label="daily",
        output_subdir="availability",
    )
    png_path = save_heatmap_png(
        daily_availability_df,
        output_subdir="availability",
        filename="Availability_grid.png",
        is_availability=True,
    )

    print("\nAvailability outputs:")
    if parquet_path:
        print(f"  Grid                 : {os.path.basename(parquet_path)}")
    if excel_path:
        print(f"  Per-station summary  : {os.path.basename(excel_path)}")
    print("  Period summaries     : station_availability_period_summary.xlsx")
    if html_path:
        print(f"  Daily heatmap        : {os.path.basename(html_path)}")
    if png_path:
        print(f"  Static heatmap       : {os.path.basename(png_path)}")


# %% 2 - Main execution: Precipitation grid

def run_rain_grid(
    precip_df: pd.DataFrame | None,
    gap_code_df: pd.DataFrame | None,
    global_index: pd.DatetimeIndex | None,
) -> None:
    """
    Build and export the precipitation grid.

    Output directory: 3_Data/4_Grids/4.2_Precipitation
    """
    print("\n" + "=" * 80)
    print("SECTION 2: RAIN GRID")
    print("=" * 80)

    if precip_df is None or gap_code_df is None or global_index is None:
        print("[ERROR] No data available to build precipitation grid.")
        return

    precip_df, gap_code_df = add_meteosuisse_to_grid(
        precip_df,
        gap_code_df,
        global_index,
    )

    print("[INFO] Exporting precipitation grid and summaries...")
    wide_path, excel_path = save_rain_outputs(precip_df, gap_code_df)
    long_parquet_path, long_csv_path = save_qc_long_outputs(precip_df, gap_code_df)
    monthly_qc_path = save_monthly_qc_excel(precip_df, gap_code_df)
    save_period_summaries(precip_df, grid_type="precip")

    print("[INFO] Generating precipitation heatmaps...")
    html_path_3min = save_heatmap_html(
        precip_df,
        is_daily=False,
        resolution_label="3min",
        output_subdir="rain",
    )
    daily_rain_df = build_daily_rain_heatmap(precip_df)
    html_path_daily = save_heatmap_html(
        daily_rain_df,
        is_daily=True,
        resolution_label="daily",
        output_subdir="rain",
    )
    png_path_daily = save_heatmap_png(
        daily_rain_df,
        output_subdir="rain",
        filename="Precip_grid.png",
        is_availability=False,
    )
    precip_30min = build_30min_heatmap(precip_df)
    html_path_30min = save_heatmap_html(
        precip_30min,
        is_daily=False,
        resolution_label="30min",
        output_subdir="rain",
    )

    print("\nRain-grid outputs:")
    if wide_path:
        print(f"  Wide grid            : {os.path.basename(wide_path)}")
    if excel_path:
        print(f"  Per-station summary  : {os.path.basename(excel_path)}")
    if long_parquet_path:
        print(f"  Long parquet table   : {os.path.basename(long_parquet_path)}")
    if long_csv_path:
        print(f"  Long CSV table       : {os.path.basename(long_csv_path)}")
    if monthly_qc_path:
        print(f"  Monthly QC workbook  : {os.path.basename(monthly_qc_path)}")
    print("  Period summaries     : station_precip_period_summary.xlsx")
    if html_path_3min:
        print(f"  3-min heatmap        : {os.path.basename(html_path_3min)}")
    if html_path_daily:
        print(f"  Daily heatmap        : {os.path.basename(html_path_daily)}")
    if html_path_30min:
        print(f"  30-min heatmap       : {os.path.basename(html_path_30min)}")
    if png_path_daily:
        print(f"  Static daily heatmap : {os.path.basename(png_path_daily)}")


# %% 3 - Main orchestration

def main() -> None:
    """Orchestrate the complete station grid building and export process."""
    print("\n" + "=" * 80)
    print("STATION GRIDS: AVAILABILITY + PRECIPITATION")
    print("=" * 80)
    print(TIME_ALIGNMENT_MESSAGE)
    
    print("[INFO] Loading consolidated station data...")
    station_data = load_consolidated_station_data()
    
    print("[INFO] Building shared 3-minute grids with overlap-based redistribution...")
    availability_df, precip_df, gap_code_df, global_index = build_station_grids(station_data)
    
    run_availability_grid(availability_df)
    run_rain_grid(precip_df, gap_code_df, global_index)

    print("\n" + "=" * 80)
    print("STATION GRID PROCESSING COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    main()
