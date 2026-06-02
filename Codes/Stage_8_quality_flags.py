# Charlotte Grosjean - 22.05.2026 - UNIL Master Thesis
"""
Assign final quality flags to consolidated station files from:
- the cross-correlation summary file;
- the correction-factor analysis file;
- a daily MeteoSwiss consistency check.

The script updates Available_Data.xlsx with:
- Flag correlation;
- Flag correction factor;
- compact final quality-control results.

Detailed diagnostics and statistics are written separately to:
- Quality_Control_report.xlsx;
- Quality_Control_statistics.xlsx.

It applies numeric per-timestep flags directly to the station consolidated CSVs.
The processed source files are not modified.

Thresholds are intentionally grouped at the top of the file so they can be
adjusted quickly while testing different quality criteria.
"""

from __future__ import annotations

import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from B_excel_utils import format_excel_report
from H_cross_correlation_utils import (
    distribute_meteo_to_target_grid,
    load_meteosuisse_source,
    load_processed_file,
    parse_station_datetime,
    regularize_station_series,
)
from I_paths_config import (
    DATA_DIRECTORY,
    AVAILABLE_DATA_FILE,
    PROCESSED_DIRECTORY,
    CONSOLIDATED_DIRECTORY,
    METEOSUISSE_FILES,
    CORRELATION_SUMMARY_FILE,
    FALLBACK_CORRELATION_SUMMARY_FILE,
    CORRECTION_FACTORS_FILE,
    QUALITY_CONTROL_DECISIONS_FILE,
    QUALITY_CONTROL_STATISTICS_FILE,
    DAILY_REPORTS_DIRECTORY,
)


# ===== FILES =====


# ===== FLAGS =====
FLAG_VALID = "valid"
FLAG_WEAK_KEEP = "weak but keep"
FLAG_REJECTED = "rejected"
FLAG_MISSING = "not calculated"


# ===== CORRELATION THRESHOLDS =====
# The best correlation is the maximum best_correlation across MeteoSwiss
# stations in Summary.xlsx.
#
# Default interpretation:
#   best correlation >= 0.70 -> valid
#   0.35 <= best correlation < 0.70 -> weak but keep
#   best correlation < 0.35 -> rejected
#
# To keep files from 0.10 upward, set CORRELATION_KEEP_MIN = 0.10.
CORRELATION_VALID_MIN = 0.5
CORRELATION_KEEP_MIN = 0.35


# ===== CORRECTION FACTOR THRESHOLDS =====
# The correction factor is classified by its relative difference from 1.0:
#   relative difference = abs(Raw_factor - 1.0)
#
# Default interpretation:
#   difference <= 10% -> valid
#   10% < difference <= 50% -> weak but keep
#   difference > 50% -> rejected
CORRECTION_VALID_MAX_RELATIVE_DIFF = 0.20
CORRECTION_KEEP_MAX_RELATIVE_DIFF = 0.50


# ===== OUTPUT COLUMNS IN Available_Data.xlsx =====
CORRELATION_FLAG_COLUMN = "Flag correlation"
CORRECTION_FLAG_COLUMN = "Flag correction factor"

# Detailed metrics are kept in the dedicated decision report. Available_Data.xlsx
# keeps the main per-file QC results for reporting.
WRITE_SUPPORTING_METRICS_TO_AVAILABLE = False
BEST_CORRELATION_COLUMN = "Best correlation"
BEST_CORRELATION_SOURCE_COLUMN = "Best correlation source"
REFERENCE_METEOSWISS_COLUMN = "Reference MeteoSwiss station"
BEST_LAG_STEPS_COLUMN = "Best lag steps"
BEST_LAG_MINUTES_COLUMN = "Best lag minutes"
RAW_CORRECTION_FACTOR_COLUMN = "Raw correction factor"
BEST_CORRECTION_SOURCE_COLUMN = "Best correction source"
FINAL_PERCENT_COLUMNS = [f"Final percent flag {flag}" for flag in [0, 1, 2, 3, 4]]
AVAILABLE_REPORT_QC_COLUMNS = [
    REFERENCE_METEOSWISS_COLUMN,
    BEST_CORRELATION_COLUMN,
    CORRELATION_FLAG_COLUMN,
    RAW_CORRECTION_FACTOR_COLUMN,
    CORRECTION_FLAG_COLUMN,
    "Initial quality flag",
    "Final file decision",
    *FINAL_PERCENT_COLUMNS,
]
CORRECTION_DETAIL_COLUMNS = [
    BEST_CORRECTION_SOURCE_COLUMN,
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
SUPPORTING_METRIC_COLUMNS = [
    BEST_CORRELATION_COLUMN,
    BEST_CORRELATION_SOURCE_COLUMN,
    BEST_LAG_STEPS_COLUMN,
    BEST_LAG_MINUTES_COLUMN,
    RAW_CORRECTION_FACTOR_COLUMN,
    *CORRECTION_DETAIL_COLUMNS,
]
AVAILABLE_QC_COLUMNS = [
    REFERENCE_METEOSWISS_COLUMN,
    CORRELATION_FLAG_COLUMN,
    CORRECTION_FLAG_COLUMN,
    "Initial quality flag",
    "N days pass",
    "N days fail",
    "N days skip",
    "Final points flag 0",
    "Final points flag 1",
    "Final points flag 2",
    "Final points flag 3",
    "Final points flag 4",
    "Daily check status",
    "Final file decision",
    *FINAL_PERCENT_COLUMNS,
    *SUPPORTING_METRIC_COLUMNS,
]


# ===== PER-TIMESTEP QUALITY FLAG IN CONSOLIDATED FILES =====
# Adds a numeric quality flag to each station file in 2_Consolidated_Stations:
#   0 = valid globally or valid daily
#   1 = weak globally, daily check skipped
#   2 = rejected globally by both correlation and correction factor
#   3 = rejected daily
#   4 = outlier precipitation value (> 15 mm / 3 minutes)
UPDATE_CONSOLIDATED_FILES = True
QUALITY_FLAG_COLUMN = "quality flag"
CONSOLIDATED_QUALITY_FLAG_COLUMN = "quality_flag"
OUTLIER_RAIN_THRESHOLD_MM_PER_TIMESTEP = 15.0

# Legacy processed-file writing is kept in helper functions below for reference,
# but Stage 8 now writes only to consolidated CSV files.
UPDATE_PROCESSED_FILES = False
BACKUP_PROCESSED_FILES = False
PROCESSED_BACKUP_DIRECTORY_PREFIX = "1_Stations_Processed_quality_flag_backup"

# Faster XLSX writing rewrites the workbook values in streaming mode instead of
# editing cells one by one. This is much faster for large files, but formatting
# from the original processed XLSX is not preserved.
FAST_XLSX_REWRITE = True


# ===== DAILY CHECK THRESHOLDS =====
TIME_STEP_MINUTES = 3
DAILY_MIN_VALID_POINTS = 40
DAILY_MIN_TOTAL_RAIN_MM = 0.5
DAILY_DRY_TOTAL_DIFF_MAX_MM = 0.2
DAILY_TOTAL_DIFF_MAX_MM = 0.5
DAILY_CORRELATION_MIN = 0.40
DAILY_BIAS_MAX = 0.30

# Consolidated files are backed up into a timestamped mirror folder before editing.
BACKUP_CONSOLIDATED_FILES = True
CONSOLIDATED_BACKUP_DIRECTORY_PREFIX = "2_Consolidated_Stations_quality_flag_backup"
DAILY_REPORT_DIRECTORY = DAILY_REPORTS_DIRECTORY
METEOSUISSE_SOURCE_ALIASES = {
    "lausanne": "LSN",
    "lsn": "LSN",
    "lsn_complete": "LSN",
    "pully": "PUY",
    "puy": "PUY",
    "puy_complete": "PUY",
    "vit": "VIT",
    "vit_complete": "VIT",
}


def normalize_key(value: object) -> str:
    """Normalize station/file keys for stable Excel joins."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_meteosuisse_source(value: object) -> str:
    """Normalize MeteoSwiss source labels from old and new summaries."""
    key = normalize_key(value).lower().replace(" ", "_")
    return METEOSUISSE_SOURCE_ALIASES.get(key, normalize_key(value))


def classify_correlation(best_correlation: float | None) -> str:
    """Return the quality flag for the best correlation value."""
    best_correlation = pd.to_numeric(best_correlation, errors="coerce")
    if not np.isfinite(best_correlation):
        return FLAG_MISSING
    if best_correlation >= CORRELATION_VALID_MIN:
        return FLAG_VALID
    if best_correlation >= CORRELATION_KEEP_MIN:
        return FLAG_WEAK_KEEP
    return FLAG_REJECTED


def classify_correction_factor(raw_factor: float | None) -> str:
    """Return the quality flag for the raw correction factor."""
    raw_factor = pd.to_numeric(raw_factor, errors="coerce")
    if not np.isfinite(raw_factor) or raw_factor <= 0:
        return FLAG_MISSING

    relative_difference = abs(raw_factor - 1.0)
    if relative_difference <= CORRECTION_VALID_MAX_RELATIVE_DIFF:
        return FLAG_VALID
    if relative_difference <= CORRECTION_KEEP_MAX_RELATIVE_DIFF:
        return FLAG_WEAK_KEEP
    return FLAG_REJECTED


def is_bad_global_flag(flag: object) -> bool:
    """Return True when a global text flag should count as rejected."""
    return normalize_key(flag) in {FLAG_REJECTED, FLAG_MISSING}


def make_initial_quality_flag(correlation_flag: object, correction_flag: object) -> int:
    """
    Combine global correlation and correction-factor flags into one numeric flag.

    0 = both tests valid
    1 = one or both tests weak, but not both rejected
    2 = both tests rejected or missing
    """
    corr = normalize_key(correlation_flag)
    corr_factor = normalize_key(correction_flag)

    if is_bad_global_flag(corr) and is_bad_global_flag(corr_factor):
        return 2
    if corr == FLAG_VALID and corr_factor == FLAG_VALID:
        return 0
    return 1


def resolve_correlation_summary_file(summary_file: Path) -> Path:
    """Use the configured Stage 6 summary file, or the legacy fallback."""
    if summary_file.exists():
        return summary_file
    if FALLBACK_CORRELATION_SUMMARY_FILE.exists():
        return FALLBACK_CORRELATION_SUMMARY_FILE
    raise FileNotFoundError(
        f"Correlation summary not found: {summary_file} "
        f"or {FALLBACK_CORRELATION_SUMMARY_FILE}"
    )


def read_long_correlation_summary(summary_file: Path) -> pd.DataFrame:
    """Read the current Stage 6 long-format XLSX/CSV summary."""
    if summary_file.suffix.lower() == ".csv":
        summary_df = pd.read_csv(summary_file)
    else:
        sheets = pd.read_excel(summary_file, sheet_name=None)
        summary_df = pd.concat(sheets.values(), ignore_index=True)

    required = {"MeteoSwiss_Station", "Station", "File", "best_correlation"}
    missing = required.difference(summary_df.columns)
    if missing:
        return read_correlation_flags_from_decision_report(QUALITY_CONTROL_DECISIONS_FILE)

    summary_df = summary_df.copy()
    summary_df["Station"] = summary_df["Station"].map(normalize_key)
    summary_df["File name"] = summary_df["File"].map(normalize_key)
    summary_df["MeteoSwiss_Station"] = summary_df["MeteoSwiss_Station"].map(normalize_key)
    summary_df[BEST_CORRELATION_COLUMN] = pd.to_numeric(
        summary_df["best_correlation"], errors="coerce"
    )
    if "best_lag_steps" in summary_df.columns:
        summary_df[BEST_LAG_STEPS_COLUMN] = pd.to_numeric(
            summary_df["best_lag_steps"], errors="coerce"
        )
    else:
        summary_df[BEST_LAG_STEPS_COLUMN] = np.nan
    if "best_lag_minutes" in summary_df.columns:
        summary_df[BEST_LAG_MINUTES_COLUMN] = pd.to_numeric(
            summary_df["best_lag_minutes"], errors="coerce"
        )
    else:
        summary_df[BEST_LAG_MINUTES_COLUMN] = np.nan

    records: list[dict[str, object]] = []
    for (station, file_name), group in summary_df.groupby(["Station", "File name"]):
        correlations = group[BEST_CORRELATION_COLUMN]
        if correlations.notna().any():
            best_idx = correlations.idxmax()
            best_row = summary_df.loc[best_idx]
            best_correlation = float(best_row[BEST_CORRELATION_COLUMN])
            best_source = normalize_key(best_row["MeteoSwiss_Station"])
            best_lag_steps = best_row[BEST_LAG_STEPS_COLUMN]
            best_lag_minutes = best_row[BEST_LAG_MINUTES_COLUMN]
        else:
            best_correlation = np.nan
            best_source = ""
            best_lag_steps = np.nan
            best_lag_minutes = np.nan

        records.append(
            {
                "Station": station,
                "File name": file_name,
                BEST_CORRELATION_COLUMN: best_correlation,
                BEST_CORRELATION_SOURCE_COLUMN: best_source,
                BEST_LAG_STEPS_COLUMN: best_lag_steps,
                BEST_LAG_MINUTES_COLUMN: best_lag_minutes,
                CORRELATION_FLAG_COLUMN: classify_correlation(best_correlation),
            }
        )

    return pd.DataFrame(records)


def read_correlation_flags_from_decision_report(report_file: Path) -> pd.DataFrame:
    """Fallback: recover per-file correlation diagnostics from the Stage 8 report."""
    if not report_file.exists():
        raise FileNotFoundError(
            f"Correlation summary is incomplete and decision report was not found: {report_file}"
        )

    report_df = pd.read_excel(report_file, sheet_name="selected_file_decisions")
    required = {"Station", "File name", BEST_CORRELATION_COLUMN}
    missing = required.difference(report_df.columns)
    if missing:
        raise ValueError(
            f"Decision report is missing correlation columns: {sorted(missing)}"
        )

    report_df = report_df.copy()
    report_df["Station"] = report_df["Station"].map(normalize_key)
    report_df["File name"] = report_df["File name"].map(normalize_key)
    report_df[BEST_CORRELATION_COLUMN] = pd.to_numeric(
        report_df[BEST_CORRELATION_COLUMN],
        errors="coerce",
    )
    if BEST_CORRELATION_SOURCE_COLUMN not in report_df.columns:
        report_df[BEST_CORRELATION_SOURCE_COLUMN] = ""
    if BEST_LAG_STEPS_COLUMN not in report_df.columns:
        report_df[BEST_LAG_STEPS_COLUMN] = np.nan
    if BEST_LAG_MINUTES_COLUMN not in report_df.columns:
        report_df[BEST_LAG_MINUTES_COLUMN] = np.nan

    report_df[CORRELATION_FLAG_COLUMN] = report_df[BEST_CORRELATION_COLUMN].apply(
        classify_correlation
    )
    columns = [
        "Station",
        "File name",
        BEST_CORRELATION_COLUMN,
        BEST_CORRELATION_SOURCE_COLUMN,
        BEST_LAG_STEPS_COLUMN,
        BEST_LAG_MINUTES_COLUMN,
        CORRELATION_FLAG_COLUMN,
    ]
    return report_df[columns].drop_duplicates(
        subset=["Station", "File name"],
        keep="first",
    )


def read_best_correlations(summary_file: Path) -> pd.DataFrame:
    """
    Read Summary.xlsx and keep the best correlation across MeteoSwiss stations.

    Summary.xlsx has two header rows:
    - row 1: MeteoSwiss station names;
    - row 2: metrics such as best_lag_minutes and best_correlation.
    """
    summary_file = resolve_correlation_summary_file(summary_file)

    if summary_file.name != "Summary.xlsx":
        return read_long_correlation_summary(summary_file)

    summary_df = pd.read_excel(summary_file, header=[0, 1])

    station_col = next(
        col for col in summary_df.columns if str(col[1]).strip().lower() == "station"
    )
    file_col = next(
        col for col in summary_df.columns if str(col[1]).strip().lower() == "file"
    )
    correlation_cols = [
        col
        for col in summary_df.columns
        if str(col[1]).strip().lower() == "best_correlation"
    ]
    if not correlation_cols:
        raise ValueError("No best_correlation columns found in correlation summary.")

    records: list[dict[str, object]] = []
    for _, row in summary_df.iterrows():
        station = normalize_key(row[station_col])
        file_name = normalize_key(row[file_col])
        if not station or not file_name:
            continue

        correlations = pd.to_numeric(row[correlation_cols], errors="coerce")
        if correlations.notna().any():
            best_col = correlations.idxmax()
            best_correlation = float(correlations.loc[best_col])
            best_source = str(best_col[0]).strip()
            lag_steps_col = (best_col[0], "best_lag_steps")
            lag_minutes_col = (best_col[0], "best_lag_minutes")
            best_lag_steps = row.get(lag_steps_col, np.nan)
            best_lag_minutes = row.get(lag_minutes_col, np.nan)
        else:
            best_correlation = np.nan
            best_source = ""
            best_lag_steps = np.nan
            best_lag_minutes = np.nan

        records.append(
            {
                "Station": station,
                "File name": file_name,
                BEST_CORRELATION_COLUMN: best_correlation,
                BEST_CORRELATION_SOURCE_COLUMN: best_source,
                BEST_LAG_STEPS_COLUMN: best_lag_steps,
                BEST_LAG_MINUTES_COLUMN: best_lag_minutes,
                CORRELATION_FLAG_COLUMN: classify_correlation(best_correlation),
            }
        )

    return pd.DataFrame(records)


def read_correction_flags(correction_file: Path) -> pd.DataFrame:
    """Read correction_factors_analysis.xlsx produced from Stage 6 selected sources."""
    if not correction_file.exists():
        raise FileNotFoundError(f"Correction-factor file not found: {correction_file}")

    correction_df = pd.read_excel(correction_file)
    required_columns = {"Station", "File name", "Raw_factor"}
    missing = required_columns.difference(correction_df.columns)
    if missing:
        raise ValueError(
            f"Correction-factor file is missing columns: {sorted(missing)}"
        )

    correction_df = correction_df.copy()
    correction_df["Station"] = correction_df["Station"].map(normalize_key)
    correction_df["File name"] = correction_df["File name"].map(normalize_key)
    if "MeteoSwiss_Station" in correction_df.columns:
        correction_df["MeteoSwiss_Station"] = correction_df["MeteoSwiss_Station"].map(
            normalize_meteosuisse_source
        )
    else:
        correction_df["MeteoSwiss_Station"] = ""

    correction_df[RAW_CORRECTION_FACTOR_COLUMN] = pd.to_numeric(
        correction_df["Raw_factor"], errors="coerce"
    )

    correction_df = correction_df.drop_duplicates(
        subset=["Station", "File name"],
        keep="first",
    ).copy()
    correction_df[BEST_CORRECTION_SOURCE_COLUMN] = correction_df["MeteoSwiss_Station"]
    correction_df[CORRECTION_FLAG_COLUMN] = correction_df[
        RAW_CORRECTION_FACTOR_COLUMN
    ].apply(classify_correction_factor)

    output_columns = [
        "Station",
        "File name",
        RAW_CORRECTION_FACTOR_COLUMN,
        BEST_CORRECTION_SOURCE_COLUMN,
        CORRECTION_FLAG_COLUMN,
    ]
    output_columns.extend(
        column
        for column in CORRECTION_DETAIL_COLUMNS
        if column in correction_df.columns and column not in output_columns
    )
    return correction_df[output_columns]


def make_backup(file_path: Path) -> Path:
    """Create a timestamped backup before updating Available_Data.xlsx."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = file_path.with_name(f"{file_path.stem}_backup_{timestamp}{file_path.suffix}")
    shutil.copy2(file_path, backup_path)
    return backup_path


def format_workbook(file_path: Path) -> None:
    """Apply simple readable formatting to all sheets in an Excel workbook."""
    workbook = load_workbook(file_path)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for worksheet in workbook.worksheets:
        if worksheet.max_row == 0:
            continue
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for column_cells in worksheet.columns:
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in column_cells
            )
            column_letter = get_column_letter(column_cells[0].column)
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 55)
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=False)
    workbook.save(file_path)


def final_file_decision(row: pd.Series) -> str:
    """Summarize final per-timestep flags into one compact file-level label."""
    counts = {}
    for flag in [0, 1, 2, 3, 4]:
        value = pd.to_numeric(row.get(f"Final points flag {flag}", 0), errors="coerce")
        counts[flag] = 0 if pd.isna(value) else int(value)
    measured_total = sum(counts.values())
    if measured_total == 0:
        return "no_measurement_rows"
    if counts[2] == measured_total:
        return "rejected_file"
    if counts[3] > 0 or counts[4] > 0:
        return "partly_rejected"
    if counts[1] > 0:
        return "uncertain_kept"
    if counts[0] > 0:
        return "valid"
    return "unknown"


def backup_processed_file(file_path: Path, backup_root: Path | None) -> None:
    """Mirror one processed file into the backup folder before editing."""
    if backup_root is None:
        return

    relative_path = file_path.relative_to(PROCESSED_DIRECTORY)
    backup_path = backup_root / relative_path
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(file_path, backup_path)


def pearson_corr(series_a: pd.Series, series_b: pd.Series) -> float:
    """Compute Pearson correlation with NaN filtering."""
    x = pd.to_numeric(series_a, errors="coerce").to_numpy(dtype=float)
    y = pd.to_numeric(series_b, errors="coerce").to_numpy(dtype=float)
    mask = np.isfinite(x) & np.isfinite(y)
    if mask.sum() < DAILY_MIN_VALID_POINTS:
        return np.nan

    x = x[mask]
    y = y[mask]
    x_centered = x - x.mean()
    y_centered = y - y.mean()
    denominator = np.sqrt(np.sum(x_centered ** 2) * np.sum(y_centered ** 2))
    if denominator == 0:
        return np.nan
    return float(np.sum(x_centered * y_centered) / denominator)


def analyze_daily_quality(
    file_path: Path,
    meteo_df: pd.DataFrame,
    best_lag_steps: int = 0,
) -> tuple[dict[object, int], dict[str, object], pd.DataFrame]:
    """Run the daily correlation/bias check for one processed file."""
    station_df = load_processed_file(str(file_path))
    if station_df.empty:
        return {}, {
            "N days pass": 0,
            "N days fail": 0,
            "N days skip": 0,
            "Daily check status": "no_station_data",
        }, pd.DataFrame()

    station_series = regularize_station_series(station_df, TIME_STEP_MINUTES)
    meteo_series = distribute_meteo_to_target_grid(
        meteo_df,
        station_series.index,
        TIME_STEP_MINUTES,
    )
    if best_lag_steps:
        meteo_series = meteo_series.shift(best_lag_steps)
    common = pd.DataFrame(
        {
            "station": station_series,
            "meteo": meteo_series,
        }
    ).dropna()

    if common.empty:
        return {}, {
            "N days pass": 0,
            "N days fail": 0,
            "N days skip": 0,
            "Daily check status": "no_overlap",
        }, pd.DataFrame()

    common["Date"] = common.index.date
    daily_flags: dict[object, int] = {}
    daily_records: list[dict[str, object]] = []

    for date_value, day_data in common.groupby("Date"):
        n_points = int(len(day_data))
        station_total = float(day_data["station"].sum())
        meteo_total = float(day_data["meteo"].sum())
        reasons: list[str] = []
        corr = np.nan
        bias = np.nan

        if n_points < DAILY_MIN_VALID_POINTS:
            status = "SKIP"
            reasons.append(f"valid points {n_points} < {DAILY_MIN_VALID_POINTS}")
        elif max(station_total, meteo_total) < DAILY_MIN_TOTAL_RAIN_MM:
            total_diff = abs(station_total - meteo_total)
            if total_diff <= DAILY_DRY_TOTAL_DIFF_MAX_MM:
                status = "PASS"
                reasons.append(
                    f"dry day totals close: diff {total_diff:.3f} "
                    f"<= {DAILY_DRY_TOTAL_DIFF_MAX_MM}"
                )
            else:
                status = "FAIL"
                reasons.append(
                    f"dry day totals differ: diff {total_diff:.3f} "
                    f"> {DAILY_DRY_TOTAL_DIFF_MAX_MM}"
                )
        else:
            corr = pearson_corr(day_data["station"], day_data["meteo"])
            total_diff = abs(station_total - meteo_total)
            mean_meteo = float(day_data["meteo"].mean())
            bias = (
                (float(day_data["station"].mean()) - mean_meteo) / mean_meteo
                if mean_meteo != 0
                else np.nan
            )

            if not np.isfinite(corr) or corr < DAILY_CORRELATION_MIN:
                corr_text = "N/A" if not np.isfinite(corr) else f"{corr:.3f}"
                reasons.append(f"corr {corr_text} < {DAILY_CORRELATION_MIN}")
            if (
                (not np.isfinite(bias) or abs(bias) > DAILY_BIAS_MAX)
                and total_diff > DAILY_TOTAL_DIFF_MAX_MM
            ):
                bias_text = "N/A" if not np.isfinite(bias) else f"{bias:.3f}"
                reasons.append(
                    f"bias {bias_text} outside +/-{DAILY_BIAS_MAX} "
                    f"and total diff {total_diff:.3f} > {DAILY_TOTAL_DIFF_MAX_MM}"
                )

            status = "FAIL" if reasons else "PASS"

        if status == "PASS":
            daily_flags[date_value] = 0
        elif status == "FAIL":
            daily_flags[date_value] = 3

        daily_records.append(
            {
                "Date": date_value,
                "Daily status": status,
                "Daily reason": "; ".join(reasons) if reasons else "OK",
                "Daily valid points": n_points,
                "Daily station rain mm": round(station_total, 4),
                "Daily MeteoSwiss rain mm": round(meteo_total, 4),
                "Daily correlation": (
                    round(corr, 4)
                    if np.isfinite(corr)
                    else np.nan
                ),
                "Daily bias": (
                    round(bias, 4)
                    if np.isfinite(bias)
                    else np.nan
                ),
            }
        )

    daily_report = pd.DataFrame(daily_records)
    counts = daily_report["Daily status"].value_counts()
    summary = {
        "N days pass": int(counts.get("PASS", 0)),
        "N days fail": int(counts.get("FAIL", 0)),
        "N days skip": int(counts.get("SKIP", 0)),
        "Daily check status": "ok",
    }
    return daily_flags, summary, daily_report


def analyze_daily_quality_from_consolidated_rows(
    measurement_df: pd.DataFrame,
    meteo_df: pd.DataFrame,
    best_lag_steps: int = 0,
) -> tuple[dict[object, int], dict[str, object], pd.DataFrame]:
    """Run the daily correlation/bias check from consolidated measurement rows."""
    if measurement_df.empty:
        return {}, {
            "N days pass": 0,
            "N days fail": 0,
            "N days skip": 0,
            "Daily check status": "no_station_data",
        }, pd.DataFrame()

    valid_measurement_df = measurement_df[
        measurement_df["P_mm"].le(OUTLIER_RAIN_THRESHOLD_MM_PER_TIMESTEP)
    ]

    station_series = (
        valid_measurement_df[["DateTime", "P_mm"]]
        .dropna(subset=["DateTime", "P_mm"])
        .sort_values("DateTime")
        .groupby("DateTime")["P_mm"]
        .sum()
        .astype(float)
    )
    station_series = station_series[station_series >= 0]
    if station_series.empty:
        return {}, {
            "N days pass": 0,
            "N days fail": 0,
            "N days skip": 0,
            "Daily check status": "no_station_data",
        }, pd.DataFrame()

    meteo_series = distribute_meteo_to_target_grid(
        meteo_df,
        station_series.index,
        TIME_STEP_MINUTES,
    )
    if best_lag_steps:
        meteo_series = meteo_series.shift(best_lag_steps)

    common = pd.DataFrame(
        {
            "station": station_series,
            "meteo": meteo_series,
        }
    ).dropna()

    if common.empty:
        return {}, {
            "N days pass": 0,
            "N days fail": 0,
            "N days skip": 0,
            "Daily check status": "no_overlap",
        }, pd.DataFrame()

    common["Date"] = common.index.date
    daily_flags: dict[object, int] = {}
    daily_records: list[dict[str, object]] = []

    for date_value, day_data in common.groupby("Date"):
        n_points = int(len(day_data))
        station_total = float(day_data["station"].sum())
        meteo_total = float(day_data["meteo"].sum())
        reasons: list[str] = []
        corr = np.nan
        bias = np.nan

        if n_points < DAILY_MIN_VALID_POINTS:
            status = "SKIP"
            reasons.append(f"valid points {n_points} < {DAILY_MIN_VALID_POINTS}")
        elif max(station_total, meteo_total) < DAILY_MIN_TOTAL_RAIN_MM:
            total_diff = abs(station_total - meteo_total)
            if total_diff <= DAILY_DRY_TOTAL_DIFF_MAX_MM:
                status = "PASS"
                reasons.append(
                    f"dry day totals close: diff {total_diff:.3f} "
                    f"<= {DAILY_DRY_TOTAL_DIFF_MAX_MM}"
                )
            else:
                status = "FAIL"
                reasons.append(
                    f"dry day totals differ: diff {total_diff:.3f} "
                    f"> {DAILY_DRY_TOTAL_DIFF_MAX_MM}"
                )
        else:
            corr = pearson_corr(day_data["station"], day_data["meteo"])
            total_diff = abs(station_total - meteo_total)
            mean_meteo = float(day_data["meteo"].mean())
            bias = (
                (float(day_data["station"].mean()) - mean_meteo) / mean_meteo
                if mean_meteo != 0
                else np.nan
            )

            if not np.isfinite(corr) or corr < DAILY_CORRELATION_MIN:
                corr_text = "N/A" if not np.isfinite(corr) else f"{corr:.3f}"
                reasons.append(f"corr {corr_text} < {DAILY_CORRELATION_MIN}")
            if (
                (not np.isfinite(bias) or abs(bias) > DAILY_BIAS_MAX)
                and total_diff > DAILY_TOTAL_DIFF_MAX_MM
            ):
                bias_text = "N/A" if not np.isfinite(bias) else f"{bias:.3f}"
                reasons.append(
                    f"bias {bias_text} outside +/-{DAILY_BIAS_MAX} "
                    f"and total diff {total_diff:.3f} > {DAILY_TOTAL_DIFF_MAX_MM}"
                )

            status = "FAIL" if reasons else "PASS"

        if status == "PASS":
            daily_flags[date_value] = 0
        elif status == "FAIL":
            daily_flags[date_value] = 3

        daily_records.append(
            {
                "Date": date_value,
                "Daily status": status,
                "Daily reason": "; ".join(reasons) if reasons else "OK",
                "Daily valid points": n_points,
                "Daily station rain mm": round(station_total, 4),
                "Daily MeteoSwiss rain mm": round(meteo_total, 4),
                "Daily correlation": round(corr, 4) if np.isfinite(corr) else np.nan,
                "Daily bias": round(bias, 4) if np.isfinite(bias) else np.nan,
            }
        )

    daily_report = pd.DataFrame(daily_records)
    counts = daily_report["Daily status"].value_counts()
    summary = {
        "N days pass": int(counts.get("PASS", 0)),
        "N days fail": int(counts.get("FAIL", 0)),
        "N days skip": int(counts.get("SKIP", 0)),
        "Daily check status": "ok",
    }
    return daily_flags, summary, daily_report


def build_quality_values_for_dates(
    dates: pd.Series,
    initial_quality_value: int,
    daily_flags: dict[object, int],
) -> list[int]:
    """Map row dates to final numeric quality values."""
    fallback_quality_value = 1 if initial_quality_value in {0, 1} else initial_quality_value
    return [
        daily_flags.get(date_value, fallback_quality_value)
        for date_value in dates
    ]


def read_processed_row_dates(file_path: Path) -> pd.Series:
    """Read one processed file and return one date per original measurement row."""
    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(file_path)
        if "created_at" in df.columns:
            datetimes = pd.to_datetime(df["created_at"], utc=True, errors="coerce").dt.tz_localize(None)
        elif "DateTime" in df.columns:
            datetimes = pd.to_datetime(df["DateTime"], errors="coerce")
        else:
            raise ValueError(f"CSV has no recognized datetime column: {file_path}")
        return datetimes.dt.date

    if suffix == ".txt":
        df = pd.read_csv(
            file_path,
            sep="\t",
            skiprows=5,
            header=None,
            usecols=[1],
            names=["DateTime"],
        )
        datetimes = df["DateTime"].apply(parse_station_datetime)
        return datetimes.dt.date

    if suffix == ".xlsx":
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        worksheet = workbook.active
        header_row = 7
        headers = [
            worksheet.cell(header_row, column).value
            for column in range(1, worksheet.max_column + 1)
        ]
        lowered = [normalize_key(header).lower() for header in headers]
        datetime_col = next(
            (
                index + 1
                for index, header in enumerate(lowered)
                if "date" in header or "time" in header
            ),
            1,
        )

        values = []
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            has_measurement = any(
                worksheet.cell(row_number, column).value is not None
                for column in range(1, worksheet.max_column + 1)
            )
            if has_measurement:
                values.append(worksheet.cell(row_number, datetime_col).value)
        workbook.close()
        datetimes = pd.Series(values).apply(parse_station_datetime)
        return datetimes.dt.date

    raise ValueError(f"Unsupported processed file type: {file_path}")


def update_csv_quality_flag(file_path: Path, quality_values: list[int]) -> None:
    """Add or update the quality flag column in a CSV processed file."""
    df = pd.read_csv(file_path)
    if len(quality_values) != len(df):
        raise ValueError(
            f"Expected {len(df)} quality values for {file_path}, got {len(quality_values)}"
        )
    df[QUALITY_FLAG_COLUMN] = quality_values
    df.to_csv(file_path, index=False)


def update_txt_quality_flag(file_path: Path, quality_values: list[int]) -> None:
    """
    Add or update the quality flag column in a TXT processed file.

    The first five lines are metadata. Measurement rows are tab-separated and
    do not have a header, so the numeric quality flag is written as a fourth
    column on each measurement row.
    """
    with open(file_path, "r", encoding="utf-8", errors="replace") as input_file:
        lines = input_file.readlines()

    output_lines: list[str] = []
    value_index = 0

    for line_number, line in enumerate(lines, start=1):
        line_ending = "\n" if line.endswith("\n") else ""
        content = line[:-1] if line_ending else line

        if line_number <= 5 or not content.strip():
            output_lines.append(line)
            continue

        parts = content.split("\t")
        if value_index >= len(quality_values):
            raise ValueError(f"Not enough quality values for {file_path}")
        quality_text = str(quality_values[value_index])
        value_index += 1

        if len(parts) >= 4:
            parts[3] = quality_text
        else:
            while len(parts) < 3:
                parts.append("")
            parts.append(quality_text)
        output_lines.append("\t".join(parts) + line_ending)

    with open(file_path, "w", encoding="utf-8") as output_file:
        output_file.writelines(output_lines)

    if value_index != len(quality_values):
        raise ValueError(
            f"Expected to write {len(quality_values)} quality values for {file_path}, "
            f"wrote {value_index}"
        )


def update_xlsx_quality_flag(file_path: Path, quality_values: list[int]) -> None:
    """Add or update the quality flag column in an XLSX processed file."""
    if FAST_XLSX_REWRITE:
        update_xlsx_quality_flag_fast(file_path, quality_values)
        return

    workbook = load_workbook(file_path)
    worksheet = workbook.active
    header_row = 7

    headers = [
        worksheet.cell(header_row, column).value
        for column in range(1, worksheet.max_column + 1)
    ]
    quality_col = next(
        (
            index + 1
            for index, header in enumerate(headers)
            if normalize_key(header).lower() == QUALITY_FLAG_COLUMN
        ),
        None,
    )
    if quality_col is None:
        non_empty_header_cols = [
            index + 1
            for index, header in enumerate(headers)
            if normalize_key(header)
        ]
        quality_col = (max(non_empty_header_cols) if non_empty_header_cols else 0) + 1
        worksheet.cell(header_row, quality_col).value = QUALITY_FLAG_COLUMN

    value_index = 0
    for row_number in range(header_row + 1, worksheet.max_row + 1):
        has_measurement = any(
            worksheet.cell(row_number, column).value is not None
            for column in range(1, quality_col)
        )
        if has_measurement:
            if value_index >= len(quality_values):
                workbook.close()
                raise ValueError(f"Not enough quality values for {file_path}")
            worksheet.cell(row_number, quality_col).value = quality_values[value_index]
            value_index += 1

    workbook.save(file_path)
    workbook.close()

    if value_index != len(quality_values):
        raise ValueError(
            f"Expected to write {len(quality_values)} quality values for {file_path}, "
            f"wrote {value_index}"
        )


def update_xlsx_quality_flag_fast(file_path: Path, quality_values: list[int]) -> None:
    """
    Add/update the quality flag column in an XLSX file using streaming I/O.

    This preserves cell values but not workbook formatting. It is intended for
    processed measurement files that are read as tabular data by Stage 7.
    """
    header_row = 7
    source_workbook = load_workbook(file_path, read_only=True, data_only=True)
    source_sheet = source_workbook.active
    output_workbook = Workbook(write_only=True)
    output_sheet = output_workbook.create_sheet(title=source_sheet.title)

    quality_col_index: int | None = None
    value_index = 0

    for row_number, row in enumerate(source_sheet.iter_rows(values_only=True), start=1):
        values = list(row)

        if row_number == header_row:
            headers = [normalize_key(value).lower() for value in values]
            quality_col_index = next(
                (
                    index
                    for index, header in enumerate(headers)
                    if header == QUALITY_FLAG_COLUMN
                ),
                None,
            )
            if quality_col_index is None:
                values.append(QUALITY_FLAG_COLUMN)
                quality_col_index = len(values) - 1
            output_sheet.append(values)
            continue

        if row_number <= header_row:
            output_sheet.append(values)
            continue

        has_measurement = any(value is not None for value in values)
        if has_measurement:
            if value_index >= len(quality_values):
                source_workbook.close()
                raise ValueError(f"Not enough quality values for {file_path}")
            if quality_col_index is None:
                raise ValueError(f"Header row not found in {file_path}")
            while len(values) <= quality_col_index:
                values.append(None)
            values[quality_col_index] = quality_values[value_index]
            value_index += 1

        output_sheet.append(values)

    source_workbook.close()

    if value_index != len(quality_values):
        raise ValueError(
            f"Expected to write {len(quality_values)} quality values for {file_path}, "
            f"wrote {value_index}"
        )

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False,
            dir=file_path.parent,
        ) as temp_file:
            temp_path = Path(temp_file.name)
        output_workbook.save(temp_path)
        shutil.move(str(temp_path), str(file_path))
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


def update_one_processed_file(file_path: Path, quality_values: list[int]) -> None:
    """Dispatch quality-flag writing according to processed file type."""
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        update_csv_quality_flag(file_path, quality_values)
    elif suffix == ".txt":
        update_txt_quality_flag(file_path, quality_values)
    elif suffix == ".xlsx":
        update_xlsx_quality_flag(file_path, quality_values)
    else:
        raise ValueError(f"Unsupported processed file type: {file_path}")


def update_processed_files_quality_flags(output_df: pd.DataFrame) -> dict[str, int]:
    """
    Add/update quality flag values in 1_Stations_Processed files.

    Global flags are assigned first. Files with initial flag 0 or 1 then get a
    daily MeteoSwiss check; daily PASS rows become 0 and daily FAIL rows become 3.
    """
    if not UPDATE_PROCESSED_FILES:
        return {"updated": 0, "missing": 0, "errors": 0, "daily_checked": 0}

    if not PROCESSED_DIRECTORY.exists():
        raise FileNotFoundError(f"Processed directory not found: {PROCESSED_DIRECTORY}")

    meteosuisse_data = {
        station: load_meteosuisse_source(file_path)
        for station, file_path in METEOSUISSE_FILES.items()
    }
    DAILY_REPORT_DIRECTORY.mkdir(parents=True, exist_ok=True)

    backup_root = None
    if BACKUP_PROCESSED_FILES:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_root = DATA_DIRECTORY / f"{PROCESSED_BACKUP_DIRECTORY_PREFIX}_{timestamp}"
        backup_root.mkdir(parents=True, exist_ok=True)

    stats = {"updated": 0, "missing": 0, "errors": 0, "daily_checked": 0}
    total_files = len(output_df)
    for position, (idx, row) in enumerate(output_df.iterrows(), start=1):
        station = normalize_key(row["Station"])
        file_name = normalize_key(row["File name"])
        file_path = PROCESSED_DIRECTORY / station / file_name
        initial_quality_value = int(row["Initial quality flag"])
        print(
            f"[{position:03d}/{total_files:03d}] Processing {station} / {file_name}",
            flush=True,
        )

        if not file_path.exists():
            stats["missing"] += 1
            print(f"    Missing processed file: {file_path}", flush=True)
            continue

        try:
            daily_flags: dict[object, int] = {}
            daily_summary = {
                "N days pass": 0,
                "N days fail": 0,
                "N days skip": 0,
                "Daily check status": "not_run_global_rejected",
            }

            if initial_quality_value in {0, 1}:
                best_source = normalize_meteosuisse_source(
                    row.get(BEST_CORRELATION_SOURCE_COLUMN, "")
                )
                meteo_df = meteosuisse_data.get(best_source)
                if meteo_df is None:
                    daily_summary["Daily check status"] = "missing_meteoswiss_source"
                else:
                    best_lag_steps = pd.to_numeric(
                        row.get(BEST_LAG_STEPS_COLUMN, 0),
                        errors="coerce",
                    )
                    best_lag_steps = 0 if pd.isna(best_lag_steps) else int(best_lag_steps)
                    daily_flags, daily_summary, daily_report = analyze_daily_quality(
                        file_path,
                        meteo_df,
                        best_lag_steps,
                    )
                    stats["daily_checked"] += 1

                    if not daily_report.empty:
                        report_path = (
                            DAILY_REPORT_DIRECTORY
                            / f"{station}_{file_path.stem}_daily_quality.csv"
                        )
                        daily_report.to_csv(report_path, index=False)

            for column, value in daily_summary.items():
                output_df.at[idx, column] = value

            row_dates = read_processed_row_dates(file_path)
            quality_values = build_quality_values_for_dates(
                row_dates,
                initial_quality_value,
                daily_flags,
            )
            backup_processed_file(file_path, backup_root)
            update_one_processed_file(file_path, quality_values)

            final_counts = pd.Series(quality_values).value_counts()
            for value in [0, 1, 2, 3, 4]:
                output_df.at[idx, f"Final points flag {value}"] = int(final_counts.get(value, 0))

            stats["updated"] += 1
            print(
                "    OK - "
                f"initial={initial_quality_value}, "
                f"daily={daily_summary['Daily check status']}, "
                f"pass={daily_summary['N days pass']}, "
                f"fail={daily_summary['N days fail']}, "
                f"skip={daily_summary['N days skip']}",
                flush=True,
            )
        except Exception as exc:
            stats["errors"] += 1
            print(
                f"    Error while updating {file_path}: {type(exc).__name__}: {exc}",
                flush=True,
            )

    if backup_root is not None:
        print(f"\nProcessed files backup: {backup_root}")

    return stats


def backup_consolidated_file(file_path: Path, backup_root: Path | None) -> None:
    """Mirror one consolidated file into the backup folder before editing."""
    if backup_root is None:
        return

    backup_path = backup_root / file_path.name
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(file_path, backup_path)


def update_consolidated_files_quality_flags(output_df: pd.DataFrame) -> dict[str, int]:
    """
    Add/update final quality flags directly in consolidated station CSV files.

    Gap rows are preserved:
    - Source_File == "No data" is not modified;
    - P_mm < 0 keeps its existing -1/-2 quality flag.
    """
    if not UPDATE_CONSOLIDATED_FILES:
        return {"updated": 0, "missing": 0, "errors": 0, "daily_checked": 0}

    if not CONSOLIDATED_DIRECTORY.exists():
        raise FileNotFoundError(
            f"Consolidated directory not found: {CONSOLIDATED_DIRECTORY}"
        )

    meteosuisse_data = {
        station: load_meteosuisse_source(file_path)
        for station, file_path in METEOSUISSE_FILES.items()
    }
    DAILY_REPORT_DIRECTORY.mkdir(parents=True, exist_ok=True)

    backup_root = None
    if BACKUP_CONSOLIDATED_FILES:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_root = DATA_DIRECTORY / f"{CONSOLIDATED_BACKUP_DIRECTORY_PREFIX}_{timestamp}"
        backup_root.mkdir(parents=True, exist_ok=True)

    stats = {"updated": 0, "missing": 0, "errors": 0, "daily_checked": 0}
    rows_by_station = {
        normalize_key(station): group.copy()
        for station, group in output_df.groupby("Station", dropna=False)
    }

    total_stations = len(rows_by_station)
    for position, (station, station_rows) in enumerate(
        sorted(rows_by_station.items(), key=lambda item: item[0]),
        start=1,
    ):
        consolidated_path = (
            CONSOLIDATED_DIRECTORY / f"{station}_Consolidated_Complete.csv"
        )
        print(
            f"[{position:03d}/{total_stations:03d}] Processing consolidated {station}",
            flush=True,
        )

        if not consolidated_path.exists():
            stats["missing"] += len(station_rows)
            print(f"    Missing consolidated file: {consolidated_path}", flush=True)
            continue

        try:
            df = pd.read_csv(consolidated_path)
            required_columns = {"Source_File", "DateTime", "P_mm"}
            missing_columns = required_columns.difference(df.columns)
            if missing_columns:
                raise ValueError(
                    f"Consolidated file is missing columns: {sorted(missing_columns)}"
                )

            df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
            df["P_mm"] = pd.to_numeric(df["P_mm"], errors="coerce")
            if "quality_flag" not in df.columns:
                df["quality_flag"] = np.where(df["P_mm"] < 0, df["P_mm"], 0)
            df["quality_flag"] = pd.to_numeric(
                df["quality_flag"], errors="coerce"
            ).fillna(0).astype(int)
            df.loc[df["Source_File"].map(normalize_key).eq("No data"), "quality_flag"] = -1
            df.loc[df["P_mm"].eq(-2), "quality_flag"] = -2
            df.loc[df["P_mm"].eq(-1), "quality_flag"] = -1

            measurement_mask = (
                df["Source_File"].map(normalize_key).ne("No data")
                & df["P_mm"].ge(0)
                & df["DateTime"].notna()
            )

            for row_idx, row in station_rows.iterrows():
                file_name = normalize_key(row["File name"])
                source_mask = measurement_mask & (
                    df["Source_File"].map(normalize_key) == file_name
                )
                n_source_rows = int(source_mask.sum())
                initial_quality_value = int(row["Initial quality flag"])

                if n_source_rows == 0:
                    stats["missing"] += 1
                    print(f"    Missing source rows: {file_name}", flush=True)
                    output_df.at[row_idx, "Daily check status"] = "missing_source_rows"
                    continue

                daily_flags: dict[object, int] = {}
                daily_summary = {
                    "N days pass": 0,
                    "N days fail": 0,
                    "N days skip": 0,
                    "Daily check status": "not_run_global_rejected",
                }

                if initial_quality_value in {0, 1}:
                    best_source = normalize_meteosuisse_source(
                        row.get(BEST_CORRELATION_SOURCE_COLUMN, "")
                    )
                    meteo_df = meteosuisse_data.get(best_source)
                    if meteo_df is None:
                        daily_summary["Daily check status"] = "missing_meteoswiss_source"
                    else:
                        best_lag_steps = pd.to_numeric(
                            row.get(BEST_LAG_STEPS_COLUMN, 0),
                            errors="coerce",
                        )
                        best_lag_steps = (
                            0 if pd.isna(best_lag_steps) else int(best_lag_steps)
                        )
                        daily_flags, daily_summary, daily_report = (
                            analyze_daily_quality_from_consolidated_rows(
                                df.loc[source_mask, ["DateTime", "P_mm"]],
                                meteo_df,
                                best_lag_steps,
                            )
                        )
                        stats["daily_checked"] += 1

                        if not daily_report.empty:
                            report_path = (
                                DAILY_REPORT_DIRECTORY
                                / f"{station}_{Path(file_name).stem}_daily_quality.csv"
                            )
                            daily_report.to_csv(report_path, index=False)

                for column, value in daily_summary.items():
                    output_df.at[row_idx, column] = value

                source_dates = df.loc[source_mask, "DateTime"].dt.date
                final_values = build_quality_values_for_dates(
                    source_dates,
                    initial_quality_value,
                    daily_flags,
                )
                final_series = pd.Series(final_values, index=df.index[source_mask])
                outlier_mask = (
                    source_mask
                    & df["P_mm"].gt(OUTLIER_RAIN_THRESHOLD_MM_PER_TIMESTEP)
                )
                final_series.loc[outlier_mask[outlier_mask].index] = 4
                df.loc[source_mask, "quality_flag"] = final_series.astype(int).to_numpy()

                final_counts = final_series.value_counts()
                for value in [0, 1, 2, 3, 4]:
                    output_df.at[row_idx, f"Final points flag {value}"] = int(
                        final_counts.get(value, 0)
                    )

                stats["updated"] += 1

            backup_consolidated_file(consolidated_path, backup_root)
            df.to_csv(consolidated_path, index=False)
            print(
                f"    OK - source files updated: {len(station_rows)}",
                flush=True,
            )

        except Exception as exc:
            stats["errors"] += 1
            print(
                f"    Error while updating {consolidated_path}: "
                f"{type(exc).__name__}: {exc}",
                flush=True,
            )

    if backup_root is not None:
        print(f"\nConsolidated files backup: {backup_root}")

    return stats


def update_available_data(
    available_data_file: Path,
    correlation_flags: pd.DataFrame,
    correction_flags: pd.DataFrame,
) -> pd.DataFrame:
    """Merge the quality flags into Available_Data.xlsx and save the workbook."""
    if not available_data_file.exists():
        raise FileNotFoundError(f"Available data file not found: {available_data_file}")

    available_df = pd.read_excel(available_data_file)
    required_columns = {"Station", "File name"}
    missing = required_columns.difference(available_df.columns)
    if missing:
        raise ValueError(f"Available data file is missing columns: {sorted(missing)}")

    for column in [
        CORRELATION_FLAG_COLUMN,
        CORRECTION_FLAG_COLUMN,
        BEST_CORRELATION_COLUMN,
        BEST_CORRELATION_SOURCE_COLUMN,
        BEST_LAG_STEPS_COLUMN,
        BEST_LAG_MINUTES_COLUMN,
        RAW_CORRECTION_FACTOR_COLUMN,
        REFERENCE_METEOSWISS_COLUMN,
        "Initial quality flag",
        "N days pass",
        "N days fail",
        "N days skip",
        "Daily check status",
        "Final points flag 0",
        "Final points flag 1",
        "Final points flag 2",
        "Final points flag 3",
        "Final points flag 4",
        *FINAL_PERCENT_COLUMNS,
        "Final file decision",
    ]:
        if column in available_df.columns:
            available_df = available_df.drop(columns=[column])
    available_df = available_df.drop(columns=AVAILABLE_QC_COLUMNS, errors="ignore")

    available_df["Station"] = available_df["Station"].map(normalize_key)
    available_df["File name"] = available_df["File name"].map(normalize_key)

    output_df = available_df.merge(
        correlation_flags,
        how="left",
        on=["Station", "File name"],
    )
    output_df = output_df.merge(
        correction_flags,
        how="left",
        on=["Station", "File name"],
    )

    output_df[CORRELATION_FLAG_COLUMN] = output_df[CORRELATION_FLAG_COLUMN].fillna(
        FLAG_MISSING
    )
    output_df[CORRECTION_FLAG_COLUMN] = output_df[CORRECTION_FLAG_COLUMN].fillna(
        FLAG_MISSING
    )
    output_df[REFERENCE_METEOSWISS_COLUMN] = output_df[
        BEST_CORRELATION_SOURCE_COLUMN
    ].fillna("")
    missing_reference = output_df[REFERENCE_METEOSWISS_COLUMN].eq("")
    if BEST_CORRECTION_SOURCE_COLUMN in output_df.columns:
        output_df.loc[missing_reference, REFERENCE_METEOSWISS_COLUMN] = output_df.loc[
            missing_reference, BEST_CORRECTION_SOURCE_COLUMN
        ].fillna("")
    output_df["Initial quality flag"] = output_df.apply(
        lambda row: make_initial_quality_flag(
            row[CORRELATION_FLAG_COLUMN],
            row[CORRECTION_FLAG_COLUMN],
        ),
        axis=1,
    )
    for column in [
        "N days pass",
        "N days fail",
        "N days skip",
        "Final points flag 0",
        "Final points flag 1",
        "Final points flag 2",
        "Final points flag 3",
        "Final points flag 4",
    ]:
        output_df[column] = 0
    for column in FINAL_PERCENT_COLUMNS:
        output_df[column] = 0.0
    output_df["Daily check status"] = "not_run"

    make_backup(available_data_file)
    compact_available_output(output_df).to_excel(available_data_file, index=False)
    format_excel_report(str(available_data_file))
    return output_df


def compact_available_output(output_df: pd.DataFrame) -> pd.DataFrame:
    """Return the availability table with compact QC columns for reporting."""
    available_output_df = output_df.copy()

    final_count_columns = [f"Final points flag {flag}" for flag in [0, 1, 2, 3, 4]]
    existing_count_columns = [
        column for column in final_count_columns if column in available_output_df.columns
    ]
    if existing_count_columns:
        measured_total = available_output_df[existing_count_columns].sum(axis=1)
        for flag in [0, 1, 2, 3, 4]:
            count_column = f"Final points flag {flag}"
            percent_column = f"Final percent flag {flag}"
            if count_column in available_output_df.columns:
                available_output_df[percent_column] = np.where(
                    measured_total > 0,
                    available_output_df[count_column] / measured_total * 100,
                    0.0,
                ).round(1)

    if REFERENCE_METEOSWISS_COLUMN not in available_output_df.columns:
        available_output_df[REFERENCE_METEOSWISS_COLUMN] = available_output_df.get(
            BEST_CORRELATION_SOURCE_COLUMN,
            "",
        )

    detail_columns_to_drop = [
        column
        for column in AVAILABLE_QC_COLUMNS
        if column not in AVAILABLE_REPORT_QC_COLUMNS
    ]
    available_output_df = available_output_df.drop(
        columns=detail_columns_to_drop,
        errors="ignore",
    )

    report_columns = [
        column
        for column in AVAILABLE_REPORT_QC_COLUMNS
        if column in available_output_df.columns
    ]
    base_columns = [
        column
        for column in available_output_df.columns
        if column not in report_columns
    ]

    comment_index = next(
        (
            idx
            for idx, column in enumerate(base_columns)
            if normalize_key(column).lower() in {"logbook comments", "comments"}
        ),
        len(base_columns),
    )
    ordered_columns = (
        base_columns[:comment_index]
        + report_columns
        + base_columns[comment_index:]
    )
    return available_output_df[ordered_columns]


def add_final_file_decisions(output_df: pd.DataFrame) -> pd.DataFrame:
    """Add the compact final decision label after final timestep counts exist."""
    output_df["Final file decision"] = output_df.apply(final_file_decision, axis=1)
    return output_df


def write_quality_control_decision_report(output_df: pd.DataFrame) -> None:
    """Write one readable workbook with all QC diagnostics by file."""
    decision_columns = [
        "Station",
        "File name",
        "File type",
        BEST_CORRELATION_COLUMN,
        BEST_CORRELATION_SOURCE_COLUMN,
        BEST_LAG_STEPS_COLUMN,
        BEST_LAG_MINUTES_COLUMN,
        CORRELATION_FLAG_COLUMN,
        RAW_CORRECTION_FACTOR_COLUMN,
        *CORRECTION_DETAIL_COLUMNS,
        CORRECTION_FLAG_COLUMN,
        "Initial quality flag",
        "N days pass",
        "N days fail",
        "N days skip",
        "Daily check status",
        "Final points flag 0",
        "Final points flag 1",
        "Final points flag 2",
        "Final points flag 3",
        "Final points flag 4",
        "Final file decision",
    ]
    per_file = output_df[
        [column for column in decision_columns if column in output_df.columns]
    ].copy()

    try:
        intensity_all_sources = pd.read_excel(CORRECTION_FACTORS_FILE)
    except Exception:
        intensity_all_sources = pd.DataFrame()

    step_summary = pd.DataFrame(
        [
            {
                "Stage": "Stage 6 - Cross-correlation",
                "Purpose": "Check temporal consistency between station rainfall and MeteoSwiss.",
                "Main metrics": "Best correlation, best lag, best MeteoSwiss source.",
                "Decision columns": CORRELATION_FLAG_COLUMN,
                "Decision rule": (
                    f"valid if r >= {CORRELATION_VALID_MIN}; weak/kept if "
                    f"{CORRELATION_KEEP_MIN} <= r < {CORRELATION_VALID_MIN}; rejected below."
                ),
            },
            {
                "Stage": "Stage 7 - Intensity correction",
                "Purpose": "Check intensity and total-volume consistency over matched periods.",
                "Main metrics": (
                    "Raw correction factor, common rainy timesteps, selected MeteoSwiss "
                    "source, and station/MeteoSwiss totals over the common period."
                ),
                "Decision columns": CORRECTION_FLAG_COLUMN,
                "Decision rule": (
                    "Uses the Stage 6 best-correlation MeteoSwiss source. "
                    f"valid if |factor-1| <= {CORRECTION_VALID_MAX_RELATIVE_DIFF}; "
                    f"weak/kept if <= {CORRECTION_KEEP_MAX_RELATIVE_DIFF}; rejected above."
                ),
            },
            {
                "Stage": "Stage 8 - Quality flags",
                "Purpose": "Combine global diagnostics, run daily checks, and write final timestep flags.",
                "Main metrics": "Initial quality flag, daily PASS/FAIL/SKIP, final flag counts.",
                "Decision columns": "Initial quality flag, Daily check status, Final file decision",
                "Decision rule": (
                    "0 = valid; 1 = uncertain kept; 2 = globally rejected file; "
                    "3 = rejected day; 4 = outlier > 15 mm/3 min."
                ),
            },
        ]
    )

    with pd.ExcelWriter(QUALITY_CONTROL_DECISIONS_FILE, engine="openpyxl") as writer:
        step_summary.to_excel(writer, sheet_name="stage_rules", index=False)
        per_file.to_excel(writer, sheet_name="selected_file_decisions", index=False)
        if not intensity_all_sources.empty:
            intensity_all_sources.to_excel(writer, sheet_name="intensity_all_sources", index=False)
    format_workbook(QUALITY_CONTROL_DECISIONS_FILE)


def read_consolidated_flag_distribution() -> pd.DataFrame:
    """Read final quality flags from consolidated files and count by station."""
    records: list[dict[str, object]] = []
    for file_path in sorted(CONSOLIDATED_DIRECTORY.glob("*_Consolidated_Complete.csv")):
        station = file_path.name.split("_")[0]
        try:
            df = pd.read_csv(file_path, usecols=lambda col: col in {"quality_flag", "P_mm"})
        except Exception:
            continue
        if "quality_flag" not in df.columns:
            continue
        flags = pd.to_numeric(df["quality_flag"], errors="coerce").dropna().astype(int)
        counts = flags.value_counts()
        row = {"Station": station, "Total timesteps": int(counts.sum())}
        for flag in [-2, -1, 0, 1, 2, 3, 4]:
            row[f"Flag {flag} timesteps"] = int(counts.get(flag, 0))
        gap_total = row["Flag -2 timesteps"] + row["Flag -1 timesteps"]
        row["Gap timesteps"] = gap_total
        row["Gap percent"] = gap_total / row["Total timesteps"] * 100 if row["Total timesteps"] else 0
        records.append(row)
    return pd.DataFrame(records)


def build_transition_statistics(output_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate initial-file flag to final-timestep flag transitions."""
    records: list[dict[str, object]] = []
    for initial_flag in [0, 1, 2]:
        subset = output_df[output_df["Initial quality flag"] == initial_flag]
        for final_flag in [0, 1, 2, 3, 4]:
            column = f"Final points flag {final_flag}"
            n_points = int(subset[column].sum()) if column in subset.columns else 0
            records.append(
                {
                    "Initial file flag": initial_flag,
                    "Final timestep flag": final_flag,
                    "Timesteps": n_points,
                }
            )
    transition_df = pd.DataFrame(records)
    total = transition_df["Timesteps"].sum()
    transition_df["Percent of measured timesteps"] = (
        transition_df["Timesteps"] / total * 100 if total else 0
    )
    return transition_df


def build_key_transitions_by_station(output_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate key daily check transitions (0→0, 0→3, 1→0, 1→3) by station."""
    records: list[dict[str, object]] = []
    for station, group in output_df.groupby("Station", dropna=False):
        transitions = {
            "0→0 (valid kept valid)": 0,
            "0→3 (valid rejected daily)": 0,
            "1→0 (weak became valid)": 0,
            "1→3 (weak rejected daily)": 0,
        }
        
        # 0→0: initial flag 0 with final flag 0
        flag_0_idx = group["Initial quality flag"] == 0
        transitions["0→0 (valid kept valid)"] = int(
            group.loc[flag_0_idx, "Final points flag 0"].sum()
        )
        
        # 0→3: initial flag 0 with final flag 3
        transitions["0→3 (valid rejected daily)"] = int(
            group.loc[flag_0_idx, "Final points flag 3"].sum()
        )
        
        # 1→0: initial flag 1 with final flag 0
        flag_1_idx = group["Initial quality flag"] == 1
        transitions["1→0 (weak became valid)"] = int(
            group.loc[flag_1_idx, "Final points flag 0"].sum()
        )
        
        # 1→3: initial flag 1 with final flag 3
        transitions["1→3 (weak rejected daily)"] = int(
            group.loc[flag_1_idx, "Final points flag 3"].sum()
        )
        
        row = {"Station": station}
        row.update(transitions)
        row["Total transitions"] = sum(transitions.values())
        records.append(row)
    
    return pd.DataFrame(records)


def build_initial_flag_statistics_by_station(
    output_df: pd.DataFrame,
    gap_stats: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Summarize file, measured-data, and gap shares by initial flag per station."""
    final_columns = [f"Final points flag {flag}" for flag in [0, 1, 2, 3, 4]]
    data_df = output_df.copy()
    existing_final_columns = [
        column for column in final_columns if column in data_df.columns
    ]
    if existing_final_columns:
        data_df["Measured data points"] = data_df[existing_final_columns].sum(axis=1)
    else:
        data_df["Measured data points"] = 0

    records: list[dict[str, object]] = []
    for station, group in data_df.groupby("Station", dropna=False):
        total_files = len(group)
        total_data_points = int(group["Measured data points"].sum())
        row: dict[str, object] = {
            "Station": station,
            "Total files": total_files,
            "Total measured data points": total_data_points,
        }
        for initial_flag in [0, 1, 2]:
            flag_subset = group[group["Initial quality flag"] == initial_flag]
            file_count = len(flag_subset)
            data_count = int(flag_subset["Measured data points"].sum())
            row[f"Initial flag {initial_flag} files"] = file_count
            row[f"Initial flag {initial_flag} files percent"] = (
                file_count / total_files * 100 if total_files else 0
            )
            row[f"Initial flag {initial_flag} data points"] = data_count
            row[f"Initial flag {initial_flag} data percent"] = (
                data_count / total_data_points * 100 if total_data_points else 0
            )
        records.append(row)

    summary = pd.DataFrame(records)
    if gap_stats is None or gap_stats.empty:
        return summary

    gap_columns = [
        "Station",
        "Total timesteps",
        "Flag -2 timesteps",
        "Flag -1 timesteps",
        "Gap timesteps",
        "Gap percent",
    ]
    available_gap_columns = [
        column for column in gap_columns if column in gap_stats.columns
    ]
    summary = summary.merge(
        gap_stats[available_gap_columns],
        how="left",
        on="Station",
    )
    summary = summary.rename(
        columns={
            "Total timesteps": "Total timesteps including gaps",
            "Flag -2 timesteps": "Initial intra-file gap timesteps (-2)",
            "Flag -1 timesteps": "Initial inter-file gap timesteps (-1)",
            "Gap timesteps": "Total gap timesteps (-1 and -2)",
            "Gap percent": "Total gap percent",
        }
    )
    for column in [
        "Initial intra-file gap timesteps (-2)",
        "Initial inter-file gap timesteps (-1)",
    ]:
        percent_column = column.replace("timesteps", "percent")
        summary[percent_column] = np.where(
            summary["Total timesteps including gaps"] > 0,
            summary[column] / summary["Total timesteps including gaps"] * 100,
            0,
        )
    summary["Measured data percent"] = np.where(
        summary["Total timesteps including gaps"] > 0,
        summary["Total measured data points"] / summary["Total timesteps including gaps"] * 100,
        0,
    )
    percent_columns = [
        column for column in summary.columns if "percent" in column.lower()
    ]
    summary[percent_columns] = summary[percent_columns].round(1)
    return summary


def build_final_flag_statistics_by_station(
    output_df: pd.DataFrame,
    gap_stats: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Summarize final measured-data flags and gaps by station."""
    final_columns = [f"Final points flag {flag}" for flag in [0, 1, 2, 3, 4]]
    existing_final_columns = [
        column for column in final_columns if column in output_df.columns
    ]

    records: list[dict[str, object]] = []
    for station, group in output_df.groupby("Station", dropna=False):
        total_files = len(group)
        total_data_points = int(group[existing_final_columns].sum().sum()) if existing_final_columns else 0
        row: dict[str, object] = {
            "Station": station,
            "Total files": total_files,
            "Total measured data points": total_data_points,
        }
        for final_flag in [0, 1, 2, 3, 4]:
            column = f"Final points flag {final_flag}"
            data_count = int(group[column].sum()) if column in group.columns else 0
            row[f"Final flag {final_flag} data points"] = data_count
            row[f"Final flag {final_flag} data percent"] = (
                data_count / total_data_points * 100 if total_data_points else 0
            )
        records.append(row)

    summary = pd.DataFrame(records)
    if gap_stats is None or gap_stats.empty:
        percent_columns = [
            column for column in summary.columns if "percent" in column.lower()
        ]
        summary[percent_columns] = summary[percent_columns].round(1)
        return summary

    gap_columns = [
        "Station",
        "Total timesteps",
        "Flag -2 timesteps",
        "Flag -1 timesteps",
        "Gap timesteps",
        "Gap percent",
    ]
    available_gap_columns = [
        column for column in gap_columns if column in gap_stats.columns
    ]
    summary = summary.merge(
        gap_stats[available_gap_columns],
        how="left",
        on="Station",
    )
    summary = summary.rename(
        columns={
            "Total timesteps": "Total timesteps including gaps",
            "Flag -2 timesteps": "Final intra-file gap timesteps (-2)",
            "Flag -1 timesteps": "Final inter-file gap timesteps (-1)",
            "Gap timesteps": "Total gap timesteps (-1 and -2)",
            "Gap percent": "Total gap percent",
        }
    )
    for column in [
        "Final intra-file gap timesteps (-2)",
        "Final inter-file gap timesteps (-1)",
    ]:
        percent_column = column.replace("timesteps", "percent")
        summary[percent_column] = np.where(
            summary["Total timesteps including gaps"] > 0,
            summary[column] / summary["Total timesteps including gaps"] * 100,
            0,
        )
    summary["Measured data percent"] = np.where(
        summary["Total timesteps including gaps"] > 0,
        summary["Total measured data points"] / summary["Total timesteps including gaps"] * 100,
        0,
    )
    percent_columns = [
        column for column in summary.columns if "percent" in column.lower()
    ]
    summary[percent_columns] = summary[percent_columns].round(1)
    return summary


def normalize_file_type_for_summary(value: object) -> str:
    """Return a compact lower-case file extension for summary tables."""
    key = normalize_key(value).lower()
    if key in {"", "nan", "none", "unknown"}:
        return "unknown"
    if key.startswith("."):
        return key
    if key in {"txt", "xlsx", "csv"}:
        return f".{key}"
    suffix = Path(key).suffix.lower()
    return suffix if suffix else key


def read_intra_file_gap_counts_by_file_type() -> pd.DataFrame:
    """Count final intra-file gap timesteps (-2) by source file extension."""
    records: list[dict[str, object]] = []
    for file_path in sorted(CONSOLIDATED_DIRECTORY.glob("*_Consolidated_Complete.csv")):
        try:
            df = pd.read_csv(
                file_path,
                usecols=lambda col: col in {"Source_File", "quality_flag", "P_mm"},
            )
        except Exception:
            continue
        if "Source_File" not in df.columns:
            continue
        if "quality_flag" in df.columns:
            flags = pd.to_numeric(df["quality_flag"], errors="coerce")
            intra_file_mask = flags.eq(-2)
        elif "P_mm" in df.columns:
            intra_file_mask = pd.to_numeric(df["P_mm"], errors="coerce").eq(-2)
        else:
            continue
        if not intra_file_mask.any():
            continue
        gap_df = df.loc[intra_file_mask, ["Source_File"]].copy()
        gap_df["File type"] = gap_df["Source_File"].map(normalize_file_type_for_summary)
        counts = gap_df["File type"].value_counts(dropna=False)
        for file_type, count in counts.items():
            records.append(
                {
                    "File type": file_type,
                    "Final flag -2 timesteps": int(count),
                }
            )
    if not records:
        return pd.DataFrame(columns=["File type", "Final flag -2 timesteps"])
    return (
        pd.DataFrame(records)
        .groupby("File type", as_index=False)["Final flag -2 timesteps"]
        .sum()
    )


def build_final_flag_statistics_by_file_type(output_df: pd.DataFrame) -> pd.DataFrame:
    """Summarize final quality-flag shares by source file extension."""
    final_columns = [f"Final points flag {flag}" for flag in [0, 1, 2, 3, 4]]
    existing_final_columns = [
        column for column in final_columns if column in output_df.columns
    ]
    if not existing_final_columns:
        return pd.DataFrame()

    data_df = output_df.copy()
    if "File type" in data_df.columns:
        data_df["File type summary"] = data_df["File type"].map(
            normalize_file_type_for_summary
        )
    else:
        data_df["File type summary"] = (
            data_df["File name"]
            .map(normalize_file_type_for_summary)
        )
    data_df["Measured timesteps"] = data_df[existing_final_columns].sum(axis=1)

    records: list[dict[str, object]] = []
    for file_type, group in data_df.groupby("File type summary", dropna=False):
        measured_total = int(group["Measured timesteps"].sum())
        row: dict[str, object] = {
            "File type": file_type,
            "N files": int(len(group)),
            "Measured timesteps": measured_total,
        }
        for flag in [0, 1, 2, 3, 4]:
            column = f"Final points flag {flag}"
            count = int(group[column].sum()) if column in group.columns else 0
            row[f"Final flag {flag} timesteps"] = count
            row[f"Final flag {flag} percent"] = (
                count / measured_total * 100 if measured_total else 0
            )
        records.append(row)

    summary = pd.DataFrame(records).sort_values("File type").reset_index(drop=True)
    intra_file_gaps = read_intra_file_gap_counts_by_file_type()
    summary = summary.merge(intra_file_gaps, how="left", on="File type")
    summary["Final flag -2 timesteps"] = (
        summary["Final flag -2 timesteps"].fillna(0).astype(int)
    )
    summary["Timesteps including intra-file gaps"] = (
        summary["Measured timesteps"] + summary["Final flag -2 timesteps"]
    )
    summary["Final flag -2 percent"] = np.where(
        summary["Timesteps including intra-file gaps"] > 0,
        summary["Final flag -2 timesteps"]
        / summary["Timesteps including intra-file gaps"]
        * 100,
        0,
    )
    for flag in [0, 1, 2, 3, 4]:
        count_column = f"Final flag {flag} timesteps"
        percent_column = f"Final flag {flag} percent"
        summary[percent_column] = np.where(
            summary["Timesteps including intra-file gaps"] > 0,
            summary[count_column] / summary["Timesteps including intra-file gaps"] * 100,
            0,
        )
    ordered_columns = [
        "File type",
        "N files",
        "Measured timesteps",
        "Final flag -2 timesteps",
        "Timesteps including intra-file gaps",
        "Final flag -2 percent",
    ]
    for flag in [0, 1, 2, 3, 4]:
        ordered_columns.extend(
            [f"Final flag {flag} timesteps", f"Final flag {flag} percent"]
        )
    summary = summary[[column for column in ordered_columns if column in summary.columns]]
    percent_columns = [
        column for column in summary.columns if "percent" in column.lower()
    ]
    summary[percent_columns] = summary[percent_columns].round(1)
    return summary


def build_file_type_share_by_final_flag(
    final_by_file_type: pd.DataFrame,
) -> pd.DataFrame:
    """Summarize the file-type composition within each final quality flag."""
    if final_by_file_type.empty or "File type" not in final_by_file_type.columns:
        return pd.DataFrame()

    records: list[dict[str, object]] = []
    file_types = final_by_file_type["File type"].astype(str).tolist()
    for flag in [-2, 0, 1, 2, 3, 4]:
        count_column = f"Final flag {flag} timesteps"
        if count_column not in final_by_file_type.columns:
            continue
        total_timesteps = int(final_by_file_type[count_column].sum())
        row: dict[str, object] = {
            "Final quality flag": flag,
            "Total timesteps": total_timesteps,
        }
        for file_type in file_types:
            value = int(
                final_by_file_type.loc[
                    final_by_file_type["File type"].astype(str).eq(file_type),
                    count_column,
                ].sum()
            )
            row[f"{file_type} timesteps"] = value
            row[f"{file_type} percent"] = (
                value / total_timesteps * 100 if total_timesteps else 0
            )
        records.append(row)

    summary = pd.DataFrame(records)
    percent_columns = [
        column for column in summary.columns if "percent" in column.lower()
    ]
    summary[percent_columns] = summary[percent_columns].round(1)
    return summary


def write_quality_control_statistics(output_df: pd.DataFrame) -> None:
    """Write overall QC statistics in a separate workbook."""
    final_columns = [f"Final points flag {flag}" for flag in [0, 1, 2, 3, 4]]
    final_counts = {
        flag: int(output_df.get(f"Final points flag {flag}", pd.Series(dtype=float)).sum())
        for flag in [0, 1, 2, 3, 4]
    }
    measured_total = sum(final_counts.values())
    overall_rows = [
        {"Metric": "Files in Available_Data", "Value": len(output_df), "Percent": 100.0},
        {"Metric": "Daily PASS days", "Value": int(output_df["N days pass"].sum()), "Percent": np.nan},
        {"Metric": "Daily FAIL days", "Value": int(output_df["N days fail"].sum()), "Percent": np.nan},
        {"Metric": "Daily SKIP days", "Value": int(output_df["N days skip"].sum()), "Percent": np.nan},
    ]
    for flag, count in final_counts.items():
        overall_rows.append(
            {
                "Metric": f"Measured timesteps final flag {flag}",
                "Value": count,
                "Percent": count / measured_total * 100 if measured_total else 0,
            }
        )
    overall = pd.DataFrame(overall_rows)

    correlation_distribution = (
        output_df[CORRELATION_FLAG_COLUMN]
        .value_counts(dropna=False)
        .rename_axis("Flag correlation")
        .reset_index(name="Files")
    )
    correction_distribution = (
        output_df[CORRECTION_FLAG_COLUMN]
        .value_counts(dropna=False)
        .rename_axis("Flag correction factor")
        .reset_index(name="Files")
    )
    initial_distribution = (
        output_df["Initial quality flag"]
        .value_counts(dropna=False)
        .sort_index()
        .rename_axis("Initial quality flag")
        .reset_index(name="Files")
    )
    final_decision_distribution = (
        output_df["Final file decision"]
        .value_counts(dropna=False)
        .rename_axis("Final file decision")
        .reset_index(name="Files")
    )

    station_stats = output_df.groupby("Station", dropna=False)[final_columns].sum().reset_index()
    station_stats["Measured timesteps"] = station_stats[final_columns].sum(axis=1)
    for flag in [0, 1, 2, 3, 4]:
        station_stats[f"Flag {flag} percent"] = np.where(
            station_stats["Measured timesteps"] > 0,
            station_stats[f"Final points flag {flag}"] / station_stats["Measured timesteps"] * 100,
            0,
        )

    transition_stats = build_transition_statistics(output_df)
    key_transitions = build_key_transitions_by_station(output_df)
    gap_stats = read_consolidated_flag_distribution()
    initial_by_station = build_initial_flag_statistics_by_station(output_df, gap_stats)
    final_by_station = build_final_flag_statistics_by_station(output_df, gap_stats)
    final_by_file_type = build_final_flag_statistics_by_file_type(output_df)
    file_type_share_by_flag = build_file_type_share_by_final_flag(final_by_file_type)

    with pd.ExcelWriter(QUALITY_CONTROL_STATISTICS_FILE, engine="openpyxl") as writer:
        overall.to_excel(writer, sheet_name="overall", index=False)
        correlation_distribution.to_excel(writer, sheet_name="correlation_flags", index=False)
        correction_distribution.to_excel(writer, sheet_name="correction_flags", index=False)
        initial_distribution.to_excel(writer, sheet_name="initial_flags", index=False)
        initial_by_station.to_excel(writer, sheet_name="initial_flags_by_station", index=False)
        final_by_station.to_excel(writer, sheet_name="final_flags_by_station", index=False)
        final_by_file_type.to_excel(writer, sheet_name="final_flags_by_file_type", index=False)
        file_type_share_by_flag.to_excel(writer, sheet_name="file_type_share_by_flag", index=False)
        final_decision_distribution.to_excel(writer, sheet_name="final_file_decisions", index=False)
        station_stats.to_excel(writer, sheet_name="station_measured_flags", index=False)
        gap_stats.to_excel(writer, sheet_name="station_gaps_all_flags", index=False)
        transition_stats.to_excel(writer, sheet_name="initial_to_final", index=False)
        key_transitions.to_excel(writer, sheet_name="daily_flag_transitions", index=False)
    format_workbook(QUALITY_CONTROL_STATISTICS_FILE)


def print_flag_summary(output_df: pd.DataFrame) -> None:
    """Print a compact summary of the assigned flags."""
    print("\nQuality flags written to Available_Data.xlsx")
    print("=" * 80)

    for column in [CORRELATION_FLAG_COLUMN, CORRECTION_FLAG_COLUMN]:
        print(f"\n{column}:")
        counts = output_df[column].value_counts(dropna=False)
        for flag, count in counts.items():
            percentage = count / len(output_df) * 100 if len(output_df) else 0
            print(f"  {flag}: {count} ({percentage:.1f}%)")

    if SUPPORTING_METRIC_COLUMNS:
        matched_correlations = output_df[BEST_CORRELATION_COLUMN].notna().sum()
        matched_factors = output_df[RAW_CORRECTION_FACTOR_COLUMN].notna().sum()
        print("\nMatched metrics:")
        print(f"  Correlations: {matched_correlations}/{len(output_df)}")
        print(f"  Correction factors: {matched_factors}/{len(output_df)}")

    print("\nInitial numeric quality flag:")
    counts = output_df["Initial quality flag"].value_counts(dropna=False).sort_index()
    for flag, count in counts.items():
        percentage = count / len(output_df) * 100 if len(output_df) else 0
        print(f"  {flag}: {count} ({percentage:.1f}%)")

    print("\nDaily check:")
    print(f"  PASS days: {int(output_df['N days pass'].sum())}")
    print(f"  FAIL days: {int(output_df['N days fail'].sum())}")
    print(f"  SKIP days: {int(output_df['N days skip'].sum())}")

    final_flag_columns = [
        column
        for column in output_df.columns
        if column.startswith("Final points flag ")
    ]
    if final_flag_columns:
        print("\nFinal timestep quality flags:")
        for column in sorted(final_flag_columns, key=lambda value: int(value.rsplit(" ", 1)[-1])):
            print(f"  {column.replace('Final points flag ', '')}: {int(output_df[column].sum())}")


def main() -> None:
    correlation_flags = read_best_correlations(CORRELATION_SUMMARY_FILE)
    correction_flags = read_correction_flags(CORRECTION_FACTORS_FILE)
    output_df = update_available_data(
        AVAILABLE_DATA_FILE,
        correlation_flags,
        correction_flags,
    )
    consolidated_stats = update_consolidated_files_quality_flags(output_df)
    output_df = add_final_file_decisions(output_df)
    compact_available_output(output_df).to_excel(AVAILABLE_DATA_FILE, index=False)
    format_excel_report(str(AVAILABLE_DATA_FILE))
    write_quality_control_decision_report(output_df)
    write_quality_control_statistics(output_df)
    print_flag_summary(output_df)
    if UPDATE_CONSOLIDATED_FILES:
        print("\nConsolidated files quality flag update:")
        for status, count in consolidated_stats.items():
            print(f"  {status}: {count}")
    print(f"\nQuality-control decision report: {QUALITY_CONTROL_DECISIONS_FILE}")
    print(f"Quality-control statistics report: {QUALITY_CONTROL_STATISTICS_FILE}")


if __name__ == "__main__":
    main()
