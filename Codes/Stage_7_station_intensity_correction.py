# Charlotte Grosjean - 07.05.2026 - UNIL Master Thesis
"""
Calculate and analyze intensity-based correction factors for processed pluviometer station files.

For each file in 1_Stations_Processed, the script:
- loads the processed station rainfall series;
- infers the file time step and regularizes the station series;
- uses the MeteoSwiss station with the best Stage 6 correlation for this file;
- distributes that MeteoSwiss 10-minute station onto the same time grid;
- calculates diagnostics for that selected MeteoSwiss station;
- identifies the common rainy timesteps (p > 0) between station and MeteoSwiss;
- compares the mean intensity ONLY ON THESE COMMON TIMESTEPS;
- calculates the multiplicative correction factor;
- compares station and MeteoSwiss total precipitation over the same common period;
- classifies the correction factor by category (<10%, 10-20%, 20-50%, >50%);
- saves the analysis results to a summary file.

NOTE: This script ONLY CALCULATES factors - it does NOT apply corrections or modify any files.
"""

from __future__ import annotations

import os
import shutil
from glob import glob
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from B_excel_utils import format_excel_report
from H_cross_correlation_utils import (
    load_meteosuisse_source,
    load_processed_file,
    infer_timestep_minutes,
    regularize_station_series,
    distribute_meteo_to_target_grid,
)
from I_paths_config import (
    PROCESSED_DIRECTORY,
    METEOSUISSE_FILES,
    CROSS_CORRELATION_DIRECTORY,
    CORRECTION_FACTORS_FILE,
    DATA_DIRECTORY,
)


STAGE6_OUTPUT_DIRECTORY = CROSS_CORRELATION_DIRECTORY / "Cross_correlation_MeteoSuisse"

SUPPORTED_EXTENSIONS = {".txt", ".xlsx", ".csv"}

# Below this relative difference, the factor is considered negligible.
MIN_RELATIVE_DIFFERENCE_TO_CORRECT = 0.10

# Guardrails: outside this range, the factor is still reported as extreme.
# These cases should be inspected manually because they often indicate a dry
# period, clogging, a time-shift issue, or a strong local rainfall gradient.
MIN_APPLIED_FACTOR = 0.50
MAX_APPLIED_FACTOR = 2.5

# Minimum number of common rainy timesteps required to compute a correction factor.
# If fewer than this many timesteps have rain in both sources, the correction
# is not calculated (status: not_enough_common_rainy_timesteps).
MIN_COMMON_RAINY_TIMESTEPS = 10

AVAILABLE_DATA_FILE = DATA_DIRECTORY / "Available_Data.xlsx"
CORRECTION_FACTOR_COLUMN = "Raw correction factor"
BEST_CORRECTION_SOURCE_COLUMN = "Best correction source"
CORRECTION_ANALYSIS_FILE = CORRECTION_FACTORS_FILE  # Output summary file for correction factor analysis
STAGE6_CORRELATION_SUMMARY_FILE = STAGE6_OUTPUT_DIRECTORY / "lausanne_cross_correlation_by_file.csv"


def mean_positive(values: pd.Series | np.ndarray) -> tuple[float, int]:
    """Return mean and count for strictly positive finite values."""
    series = pd.Series(values, dtype=float)
    positives = series[np.isfinite(series) & (series > 0)]
    if positives.empty:
        return np.nan, 0
    return float(positives.mean()), int(len(positives))


def common_period_total_statistics(
    station_grid: pd.Series,
    meteo_grid: pd.Series,
) -> dict[str, object]:
    """Compare station and MeteoSwiss totals over exactly the same timestamps."""
    common = pd.DataFrame(
        {
            "station": pd.to_numeric(station_grid, errors="coerce"),
            "meteo": pd.to_numeric(meteo_grid, errors="coerce"),
        }
    ).dropna()
    common = common[(common["station"] >= 0) & (common["meteo"] >= 0)]

    if common.empty:
        return {
            "Common_period_start": pd.NaT,
            "Common_period_end": pd.NaT,
            "Common_period_timesteps": 0,
            "Station_total_common_period_mm": np.nan,
            "MeteoSwiss_total_common_period_mm": np.nan,
            "Total_difference_common_period_mm": np.nan,
            "Total_absolute_difference_common_period_mm": np.nan,
            "Total_relative_bias_common_period": np.nan,
            "Total_ratio_station_to_meteoswiss": np.nan,
        }

    station_total = float(common["station"].sum())
    meteo_total = float(common["meteo"].sum())
    difference = station_total - meteo_total
    relative_bias = difference / meteo_total if meteo_total > 0 else np.nan
    ratio = station_total / meteo_total if meteo_total > 0 else np.nan

    return {
        "Common_period_start": common.index.min(),
        "Common_period_end": common.index.max(),
        "Common_period_timesteps": int(len(common)),
        "Station_total_common_period_mm": station_total,
        "MeteoSwiss_total_common_period_mm": meteo_total,
        "Total_difference_common_period_mm": difference,
        "Total_absolute_difference_common_period_mm": abs(difference),
        "Total_relative_bias_common_period": relative_bias,
        "Total_ratio_station_to_meteoswiss": ratio,
    }


def list_processed_files(processed_directory: Path) -> list[Path]:
    """Return supported processed files, ignoring temporary Excel files."""
    files: list[Path] = []
    for file_path in glob(str(processed_directory / "*" / "*")):
        path = Path(file_path)
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(path)
    return sorted(files, key=lambda p: (p.parent.name, p.name))


def load_stage6_best_meteosuisse_lookup(
    summary_file: Path,
) -> dict[tuple[str, str], str]:
    """
    Return the MeteoSwiss station with the best Stage 6 correlation per file.

    This mirrors the Stage 8 validation logic: for each processed file, the
    selected source is the row with the highest finite best_correlation.
    """
    if not summary_file.exists():
        raise FileNotFoundError(
            f"Stage 6 correlation summary not found: {summary_file}"
        )

    if summary_file.suffix.lower() == ".csv":
        summary_df = pd.read_csv(summary_file)
    else:
        sheets = pd.read_excel(summary_file, sheet_name=None)
        summary_df = pd.concat(sheets.values(), ignore_index=True)

    required_columns = {"MeteoSwiss_Station", "Station", "File", "best_correlation"}
    missing = required_columns.difference(summary_df.columns)
    if missing:
        raise ValueError(
            f"Stage 6 correlation summary is missing columns: {sorted(missing)}"
        )

    summary_df = summary_df.copy()
    summary_df["best_correlation"] = pd.to_numeric(
        summary_df["best_correlation"],
        errors="coerce",
    )

    lookup: dict[tuple[str, str], str] = {}
    for (station, file_name), group in summary_df.groupby(["Station", "File"], dropna=False):
        finite = group[np.isfinite(group["best_correlation"])]
        if finite.empty:
            continue

        best_row = finite.loc[finite["best_correlation"].idxmax()]
        lookup[(str(station), str(file_name))] = str(best_row["MeteoSwiss_Station"])

    return lookup


def calculate_meteoswiss_correction(
    station_grid: pd.Series,
    timestep_minutes: float,
    meteo_station: str,
    meteo_df: pd.DataFrame,
) -> dict[str, object]:
    """Calculate diagnostics against one MeteoSwiss station."""
    meteo_grid = distribute_meteo_to_target_grid(
        meteo_df,
        station_grid.index,
        timestep_minutes,
    )

    common_rainy_mask = (station_grid > 0) & (meteo_grid > 0)
    common_rainy_timesteps = common_rainy_mask.sum()
    common_total_stats = common_period_total_statistics(station_grid, meteo_grid)

    # Check if we have enough common rainy timesteps
    if common_rainy_timesteps < MIN_COMMON_RAINY_TIMESTEPS:
        station_mean_all, station_count_all = mean_positive(station_grid)
        meteo_mean_all, meteo_count_all = mean_positive(meteo_grid)
        
        return {
            "MeteoSwiss_Station": meteo_station,
            "Status": "not_enough_common_rainy_timesteps",
            "Raw_factor": np.nan,
            "Time_step_min": timestep_minutes,
            "Station_positive_count_all": station_count_all,
            "MeteoSwiss_positive_count_all": meteo_count_all,
            "Common_rainy_timesteps": int(common_rainy_timesteps),
            **common_total_stats,
        }

    # Calculate mean intensity ONLY on common rainy timesteps
    station_mean_common = float(station_grid[common_rainy_mask].mean())
    meteo_mean_common = float(meteo_grid[common_rainy_mask].mean())

    # Also keep track of overall statistics for reporting
    station_mean_all, station_count_all = mean_positive(station_grid)
    meteo_mean_all, meteo_count_all = mean_positive(meteo_grid)

    raw_factor = meteo_mean_common / station_mean_common
    relative_difference = abs(raw_factor - 1.0)

    # Diagnostic status only: no correction is applied to the source files here.
    if relative_difference < MIN_RELATIVE_DIFFERENCE_TO_CORRECT:
        status = "factor_within_10_percent"
    elif raw_factor < MIN_APPLIED_FACTOR or raw_factor > MAX_APPLIED_FACTOR:
        status = "factor_outside_guardrails"
    else:
        status = "factor_calculated"

    return {
        "MeteoSwiss_Station": meteo_station,
        "Status": status,
        "Raw_factor": float(raw_factor),
        "Relative_difference": float(relative_difference),
        "Time_step_min": timestep_minutes,
        # Common rainy timesteps statistics
        "Common_rainy_timesteps": int(common_rainy_timesteps),
        "Station_mean_common_mm": station_mean_common,
        "MeteoSwiss_mean_common_mm": meteo_mean_common,
        # Overall statistics (for reference)
        "Station_positive_count_all": station_count_all,
        "MeteoSwiss_positive_count_all": meteo_count_all,
        "Station_mean_all_mm": station_mean_all,
        "MeteoSwiss_mean_all_mm": meteo_mean_all,
        **common_total_stats,
    }


def calculate_file_corrections(
    file_path: Path,
    meteosuisse_data: dict[str, pd.DataFrame],
    selected_meteosuisse_station: str | None,
) -> list[dict[str, object]]:
    """
    Calculate the correction-factor diagnostic for the selected MeteoSwiss station.
    
    The correction is based on the mean intensity ONLY on timesteps where
    both the station and the selected MeteoSwiss station have positive rainfall.
    """
    if not selected_meteosuisse_station:
        return [
            {
                "MeteoSwiss_Station": "",
                "Status": "no_stage6_best_meteosuisse_station",
                "Raw_factor": np.nan,
            }
        ]

    if selected_meteosuisse_station not in meteosuisse_data:
        return [
            {
                "MeteoSwiss_Station": selected_meteosuisse_station,
                "Status": "stage6_best_meteosuisse_station_not_loaded",
                "Raw_factor": np.nan,
            }
        ]

    selected_meteosuisse_data = {
        selected_meteosuisse_station: meteosuisse_data[selected_meteosuisse_station]
    }

    station_df = load_processed_file(str(file_path))
    if station_df.empty:
        return [{"MeteoSwiss_Station": station, "Status": "no_station_data", "Raw_factor": np.nan}
                for station in selected_meteosuisse_data]

    timestep_minutes = infer_timestep_minutes(station_df)
    if timestep_minutes is None:
        return [{"MeteoSwiss_Station": station, "Status": "no_valid_timestep", "Raw_factor": np.nan}
                for station in selected_meteosuisse_data]

    station_grid = regularize_station_series(station_df, timestep_minutes)
    if not selected_meteosuisse_data:
        return [{"MeteoSwiss_Station": "", "Status": "no_meteosuisse_data", "Raw_factor": np.nan}]

    return [
        calculate_meteoswiss_correction(
            station_grid=station_grid,
            timestep_minutes=timestep_minutes,
            meteo_station=meteo_station,
            meteo_df=meteo_df,
        )
        for meteo_station, meteo_df in selected_meteosuisse_data.items()
    ]


def classify_correction_factor(raw_factor: float | None) -> str:
    """
    Classify a raw correction factor by its relative difference from 1.0.
    
    Categories:
    - <10%: No significant correction
    - 10-20%: Small correction
    - 20-50%: Moderate correction
    - >50%: Large correction
    - N/A: Factor not calculated
    """
    if not np.isfinite(raw_factor):
        return "N/A"
    
    relative_diff = abs(raw_factor - 1.0) * 100
    
    if relative_diff < 10:
        return "<10%"
    elif relative_diff < 20:
        return "10-20%"
    elif relative_diff < 50:
        return "20-50%"
    else:
        return ">50%"


def update_available_data(available_data_file: Path, results_df: pd.DataFrame) -> None:
    """Add/update the best raw correction factor column in Available_Data.xlsx."""
    if not available_data_file.exists():
        print(f"Warning: {available_data_file} not found. Skipping update.", flush=True)
        return
    
    available_df = pd.read_excel(available_data_file)
    if CORRECTION_FACTOR_COLUMN not in available_df.columns:
        available_df[CORRECTION_FACTOR_COLUMN] = np.nan
    if BEST_CORRECTION_SOURCE_COLUMN not in available_df.columns:
        available_df[BEST_CORRECTION_SOURCE_COLUMN] = ""

    best_rows = select_best_correction_rows(results_df)
    factor_lookup = {}
    source_lookup = {}
    for _, row in best_rows.iterrows():
        key = (row["Station"], row["File name"], row["File type"])
        factor_lookup[key] = row["Raw_factor"]
        source_lookup[key] = row.get("MeteoSwiss_Station", "")

    for idx, row in available_df.iterrows():
        key = (row["Station"], row["File name"], row["File type"])
        if key in factor_lookup:
            available_df.at[idx, CORRECTION_FACTOR_COLUMN] = factor_lookup[key]
            available_df.at[idx, BEST_CORRECTION_SOURCE_COLUMN] = source_lookup.get(key, "")

    available_df.to_excel(available_data_file, index=False)
    format_excel_report(str(available_data_file))


def select_best_correction_rows(results_df: pd.DataFrame) -> pd.DataFrame:
    """Keep the MeteoSwiss station whose raw factor is closest to 1 for each file."""
    if results_df.empty:
        return results_df.copy()

    df = results_df.copy()
    df["_abs_factor_difference"] = pd.to_numeric(df["Raw_factor"], errors="coerce").sub(1).abs()
    best_indices = []

    for _, group in df.groupby(["Station", "File name", "File type"], dropna=False):
        finite = group[np.isfinite(group["_abs_factor_difference"])]
        if not finite.empty:
            best_indices.append(finite["_abs_factor_difference"].idxmin())
        else:
            best_indices.append(group.index[0])

    return df.loc[best_indices].drop(columns=["_abs_factor_difference"])


def main() -> None:
    meteosuisse_data = {
        station: load_meteosuisse_source(file_path)
        for station, file_path in METEOSUISSE_FILES.items()
    }
    best_meteosuisse_by_file = load_stage6_best_meteosuisse_lookup(
        STAGE6_CORRELATION_SUMMARY_FILE
    )
    processed_files = list_processed_files(PROCESSED_DIRECTORY)

    results: list[dict[str, object]] = []
    for index, source_path in enumerate(processed_files, start=1):
        station = source_path.parent.name

        base_result = {
            "Station": station,
            "File name": source_path.name,
            "File type": source_path.suffix.lower()
        }
        selected_meteosuisse_station = best_meteosuisse_by_file.get(
            (station, source_path.name)
        )
        try:
            file_results = [
                {**base_result, **correction_result}
                for correction_result in calculate_file_corrections(
                    source_path,
                    meteosuisse_data,
                    selected_meteosuisse_station,
                )
            ]
        except Exception as exc:
            file_results = [
                {
                    **base_result,
                    "MeteoSwiss_Station": selected_meteosuisse_station or "",
                    "Status": "error",
                    "Raw_factor": np.nan,
                    "Error": f"{type(exc).__name__}: {exc}",
                }
            ]

        results.extend(file_results)
        status_summary = ", ".join(
            f"{row.get('MeteoSwiss_Station', '')}={row.get('Status', '')}"
            for row in file_results
        )
        best_row = select_best_correction_rows(pd.DataFrame(file_results)).iloc[0]
        print(
            f"[{index:03d}/{len(processed_files):03d}] "
            f"{station} / {source_path.name}: {status_summary} "
            f"| best={best_row.get('MeteoSwiss_Station', '')} "
            f"(raw_factor={best_row.get('Raw_factor', np.nan):.4g}, "
            f"common_rainy={best_row.get('Common_rainy_timesteps', 'N/A')})",
            flush=True,
        )

    results_df = pd.DataFrame(results)
    
    # Add classification column for correction factors
    results_df["Factor_category"] = results_df["Raw_factor"].apply(classify_correction_factor)
    
    # Add a column for relative difference percentage
    relative_difference = results_df.get(
        "Relative_difference",
        pd.Series(np.nan, index=results_df.index),
    )
    results_df["Relative_difference_percent"] = relative_difference.apply(
        lambda x: f"{x*100:.2f}%" if np.isfinite(x) else "N/A"
    )
    
    total_relative_bias = results_df.get(
        "Total_relative_bias_common_period",
        pd.Series(np.nan, index=results_df.index),
    )
    results_df["Total_relative_bias_common_period_percent"] = total_relative_bias.apply(
        lambda x: f"{x*100:.2f}%" if np.isfinite(x) else "N/A"
    )

    useful_columns = [
        "Station",
        "File name",
        "File type",
        "MeteoSwiss_Station",
        "Status",
        "Raw_factor",
        "Relative_difference_percent",
        "Factor_category",
        "Common_rainy_timesteps",
        "Common_period_start",
        "Common_period_end",
        "Common_period_timesteps",
        "Station_total_common_period_mm",
        "MeteoSwiss_total_common_period_mm",
        "Total_difference_common_period_mm",
        "Total_absolute_difference_common_period_mm",
        "Total_relative_bias_common_period_percent",
        "Total_ratio_station_to_meteoswiss",
    ]
    results_export = results_df[[col for col in useful_columns if col in results_df.columns]]
    best_results_export = select_best_correction_rows(results_export)
    
    summary_excel = CORRECTION_ANALYSIS_FILE
    with pd.ExcelWriter(summary_excel, engine="openpyxl") as writer:
        results_export.to_excel(writer, sheet_name="selected_meteoswiss_source", index=False)
        best_results_export.to_excel(writer, sheet_name="best_by_file", index=False)
    format_excel_report(str(summary_excel))
    
    # Available_Data.xlsx is kept as an availability table. QC diagnostics are
    # exported here and later gathered in Stage_8 Quality_Control_report.xlsx.

    # Print summary statistics
    print(f"\n{'='*80}", flush=True)
    print(f"CORRECTION FACTORS ANALYSIS (NO CORRECTIONS APPLIED)", flush=True)
    print(f"{'='*80}", flush=True)
    n_files_analyzed = results_df[["Station", "File name", "File type"]].drop_duplicates().shape[0]
    print(f"Total files analyzed: {n_files_analyzed}", flush=True)
    print(f"Total selected-source comparisons: {len(results_df)}", flush=True)
    print(f"\nStatus distribution:", flush=True)
    status_counts = results_df["Status"].value_counts()
    for status, count in status_counts.items():
        percentage = (count / len(results_df)) * 100
        print(f"  {status}: {count} ({percentage:.1f}%)", flush=True)
    
    print(f"\nCorrection factor categories distribution:", flush=True)
    factor_counts = results_df["Factor_category"].value_counts()
    for category in ["<10%", "10-20%", "20-50%", ">50%", "N/A"]:
        count = factor_counts.get(category, 0)
        percentage = (count / len(results_df)) * 100 if count > 0 else 0
        print(f"  {category}: {count} ({percentage:.1f}%)", flush=True)
    
    # Statistics on common rainy timesteps
    valid_results = results_df[results_df["Status"].isin(["factor_calculated", "factor_within_10_percent", "factor_outside_guardrails"])]
    if not valid_results.empty:
        print(f"\nCommon rainy timesteps statistics (valid files):", flush=True)
        print(f"  Mean: {valid_results['Common_rainy_timesteps'].mean():.1f}", flush=True)
        print(f"  Median: {valid_results['Common_rainy_timesteps'].median():.1f}", flush=True)
        print(f"  Min: {int(valid_results['Common_rainy_timesteps'].min())}", flush=True)
        print(f"  Max: {int(valid_results['Common_rainy_timesteps'].max())}", flush=True)
    
    # Raw factors summary (selected Stage 6 sources, including those not applied)
    all_factors = results_df[np.isfinite(results_df["Raw_factor"])]["Raw_factor"]
    if not all_factors.empty:
        print(f"\nSelected raw correction factors summary ({len(all_factors)} files with calculated factors):", flush=True)
        print(f"  Mean factor: {all_factors.mean():.3f}", flush=True)
        print(f"  Median factor: {all_factors.median():.3f}", flush=True)
        print(f"  Min factor: {all_factors.min():.3f}", flush=True)
        print(f"  Max factor: {all_factors.max():.3f}", flush=True)

    best_results = select_best_correction_rows(results_df)
    best_factors = best_results[np.isfinite(best_results["Raw_factor"])]["Raw_factor"]
    if not best_factors.empty:
        print(f"\nBest raw correction factors summary ({len(best_factors)} files with calculated factors):", flush=True)
        print(f"  Mean factor: {best_factors.mean():.3f}", flush=True)
        print(f"  Median factor: {best_factors.median():.3f}", flush=True)
        print(f"  Min factor: {best_factors.min():.3f}", flush=True)
        print(f"  Max factor: {best_factors.max():.3f}", flush=True)
    
    print(f"\n{'='*80}", flush=True)
    print(f"Analysis results saved in: {summary_excel}", flush=True)
    print(f"Available data not modified by this stage.", flush=True)
    print(f"{'='*80}\n", flush=True)


if __name__ == "__main__":
    main()
