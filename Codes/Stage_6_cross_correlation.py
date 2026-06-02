# Charlotte Grosjean - 24.05.2026 - UNIL Master Thesis
"""
Cross-correlation analysis.

The script runs the two correlation checks used by the quality-control
pipeline:

1. MeteoSwiss station-to-station correlation.
2. MeteoSwiss vs processed Pluvimate files.

An optional Lexplore NetCDF hook is kept at the bottom of the file, but it is
not part of the default Stage 6 workflow.
"""

# %% 0 - Imports and configuration

from pathlib import Path
from datetime import datetime, time
from itertools import combinations
import os
import re
import tempfile

import numpy as np
import pandas as pd
from openpyxl import load_workbook

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "matplotlib"))
os.environ.setdefault("XDG_CACHE_HOME", tempfile.gettempdir())

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
except ImportError:
    go = None
    make_subplots = None

from B_excel_utils import format_excel_report
from A_data_processing_utils import extract_numeric, parse_date
from I_paths_config import (
    DATA_DIRECTORY,
    METEOSUISSE_DIRECTORY,
    PROCESSED_DIRECTORY,
    CROSS_CORRELATION_DIRECTORY,
)


# ============================================================================
# COMMON CONFIGURATION
# ============================================================================

# %% Directories

DATA_DIR = DATA_DIRECTORY
METEOSWISS_DIR = METEOSUISSE_DIRECTORY
PROCESSED_DIR = PROCESSED_DIRECTORY
CROSS_CORRELATION_DIR = CROSS_CORRELATION_DIRECTORY
LEXPLORE_NC_DIR = DATA_DIR / "7_LEXP_Meteostation" / "léxploremeteostation_datalakesdownload 2"

METEOSWISS_OUTPUT_DIR = CROSS_CORRELATION_DIR / "MeteoSwiss"
METEOSWISS_PROCESSED_OUTPUT_DIR = CROSS_CORRELATION_DIR / "Cross_correlation_MeteoSuisse"
LEXPLORE_OUTPUT_DIR = CROSS_CORRELATION_DIR / "Cross_correlation_LEXP_Meteostation_NC"


# %% Input and output files

METEOSWISS_STATION_FILES = {
    "LSN": METEOSWISS_DIR / "LSN_Complete.csv",
    "PUY": METEOSWISS_DIR / "PUY_Complete.csv",
    "VIT": METEOSWISS_DIR / "VIT_Complete.csv",
    "LEXP": METEOSWISS_DIR / "LEXP_WS_Complete.csv",
}

REFERENCE_METEOSWISS_STATIONS = {
    "LSN": METEOSWISS_STATION_FILES["LSN"],
    "PUY": METEOSWISS_STATION_FILES["PUY"],
    "VIT": METEOSWISS_STATION_FILES["VIT"],
}

METEOSWISS_PAIR_SUMMARY_FILE = METEOSWISS_OUTPUT_DIR / "meteosuisse_three_station_cross_correlation.csv"

LEXPLORE_NC_SERIES_FILE = "lexplore_meteostation_nc_rain_complete.csv"
LEXPLORE_METEOSWISS_EXPORT_FILE = "LEXP_WS_Complete.csv"
PROCESSED_CORRELATION_SUMMARY_BASENAME = "lausanne_cross_correlation_by_file"


# %% Analysis parameters

MIN_VALID_PAIRS = 10
METEOSWISS_STEP_MINUTES = 10
METEOSWISS_PAIR_MAX_LAG_MINUTES = 360
PROCESSED_FILE_MAX_LAG_MINUTES = 240
CREATE_DIAGNOSTIC_FIGURES = True
CREATE_INTERACTIVE_FIGURES = True
CREATE_REPORT_PNG_FIGURES = True
FIGURE_DPI = 150
MAX_SCATTER_POINTS = 5000

REPORT_TITLE_FONTSIZE = 22
REPORT_SUBTITLE_FONTSIZE = 18
REPORT_LABEL_FONTSIZE = 16
REPORT_TICK_FONTSIZE = 14
REPORT_LEGEND_FONTSIZE = 15

INTERACTIVE_TITLE_FONT_SIZE = 22
INTERACTIVE_SUBTITLE_FONT_SIZE = 40
INTERACTIVE_AXIS_TITLE_FONT_SIZE = 30
INTERACTIVE_AXIS_TICK_FONT_SIZE = 30
INTERACTIVE_LEGEND_FONT_SIZE = 38
INTERACTIVE_FIGURE_WIDTH = 2800
REPORT_DATE_TICK_FORMAT = "%Y-%m-%d\n%H:%M"
INTERACTIVE_DATE_TICK_FORMAT = "%Y-%m-%d<br>%H:%M"
INTERACTIVE_XAXIS_TITLE_STANDOFF = 55

STATION_TO_PROCESS = None  # None = process all stations.
LEXPLORE_STATION_TO_PROCESS = "LEXP"

LEXPLORE_RAIN_VARIABLE = "Rain"
LEXPLORE_RAIN_QUALITY_VARIABLE = "Rain_qual"
LEXPLORE_TIME_VARIABLE = "time"
LEXPLORE_SOURCE_NAME = "Lexplore Meteostation NC"


# ============================================================================
# METEOSWISS STATION-TO-STATION CORRELATION
# ============================================================================


# %% MeteoSwiss station-to-station functions

def load_meteoswiss_station(file_path):
    """Load a MeteoSwiss station file."""
    df = pd.read_csv(file_path, sep=None, engine="python")
    df = df[["reference_timestamp", "rre150z0"]].copy()
    df.columns = ["DateTime", "P_mm"]
    df["DateTime"] = pd.to_datetime(df["DateTime"], dayfirst=True, errors="coerce")
    df["P_mm"] = (
        df["P_mm"]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .pipe(pd.to_numeric, errors="coerce")
    )
    df = df.dropna(subset=["DateTime", "P_mm"])
    df = df[df["P_mm"] >= 0]
    df = df.sort_values("DateTime").drop_duplicates("DateTime")
    return df.set_index("DateTime")["P_mm"]


def correlation_at_lag(series_1, series_2, lag_steps):
    """Compute correlation after shifting the second series."""
    shifted_2 = series_2.shift(lag_steps)
    paired = pd.concat([series_1.rename("s1"), shifted_2.rename("s2")], axis=1).dropna()

    if len(paired) < MIN_VALID_PAIRS:
        return np.nan, len(paired)
    if paired["s1"].nunique() < 2 or paired["s2"].nunique() < 2:
        return np.nan, len(paired)

    return paired["s1"].corr(paired["s2"]), len(paired)


def find_best_pair_lag(series_1, series_2):
    """Find the best lag for two series."""
    max_lag_steps = int(METEOSWISS_PAIR_MAX_LAG_MINUTES / METEOSWISS_STEP_MINUTES)
    rows = []

    for lag_steps in range(-max_lag_steps, max_lag_steps + 1):
        corr, n_pairs = correlation_at_lag(series_1, series_2, lag_steps)
        rows.append(
            {
                "lag_steps": lag_steps,
                "lag_minutes": lag_steps * METEOSWISS_STEP_MINUTES,
                "correlation": corr,
                "n_pairs": n_pairs,
            }
        )

    lag_df = pd.DataFrame(rows).dropna(subset=["correlation"])
    if lag_df.empty:
        return np.nan, np.nan, 0

    best = lag_df.loc[lag_df["correlation"].abs().idxmax()]
    return best["lag_minutes"], best["correlation"], int(best["n_pairs"])


def run_meteoswiss_station_correlation():
    """Run MeteoSwiss station-to-station cross-correlation."""
    print("\n" + "="*80)
    print("METEOSWISS STATION-TO-STATION CORRELATION")
    print("="*80 + "\n")

    METEOSWISS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    series = {}
    for station, path in METEOSWISS_STATION_FILES.items():
        print(f"Loading {station}: {path.name}")
        series[station] = load_meteoswiss_station(path)

    results = []
    pairs = list(combinations(series.keys(), 2))

    for i, (station_1, station_2) in enumerate(pairs, start=1):
        print(f"[{i:02d}/{len(pairs):02d}] Comparing {station_1} and {station_2}")
        common = pd.concat(
            [series[station_1].rename(station_1), series[station_2].rename(station_2)],
            axis=1,
        ).dropna()

        best_lag, best_corr, n_pairs = find_best_pair_lag(
            common[station_1],
            common[station_2],
        )

        results.append(
            {
                "station_1": station_1,
                "station_2": station_2,
                "best_lag_minutes": best_lag,
                "best_correlation": best_corr,
                "n_pairs_at_best_lag": n_pairs,
                "common_start": common.index.min(),
                "common_end": common.index.max(),
                "n_pairs_no_lag": len(common),
            }
        )

    result_df = pd.DataFrame(results)
    result_df.to_csv(METEOSWISS_PAIR_SUMMARY_FILE, index=False)
    print(result_df.to_string(index=False))
    print(f"\nSaved: {METEOSWISS_PAIR_SUMMARY_FILE}")


# ============================================================================
# METEOSWISS VS PROCESSED STATION FILES
# ============================================================================

# %% MeteoSwiss vs processed-file functions

def parse_station_datetime(value):
    """Parse processed-station timestamps, including AM/PM strings with hour 0."""
    parsed = parse_date(value)
    if pd.notna(parsed):
        return parsed

    if pd.isna(value):
        return pd.NaT

    text = str(value).strip().replace(",", "")
    match = re.match(
        r"^(?P<month>[A-Za-zéû\.]+)\s+"
        r"(?P<day>\d{1,2})\s+"
        r"(?P<year>\d{4})\s+"
        r"(?P<hour>\d{1,2}):(?P<minute>\d{2}):(?P<second>\d{2})\s+"
        r"(?P<ampm>AM|PM)$",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return pd.to_datetime(text, errors="coerce")

    month_map = {
        "jan": 1, "janv": 1,
        "feb": 2, "fev": 2, "févr": 2, "fevr": 2,
        "mar": 3, "mars": 3,
        "apr": 4, "avr": 4,
        "may": 5, "mai": 5,
        "jun": 6, "juin": 6,
        "jul": 7, "juil": 7,
        "aug": 8, "août": 8, "aout": 8,
        "sep": 9, "sept": 9,
        "oct": 10,
        "nov": 11,
        "dec": 12, "déc": 12,
    }

    month_key = match.group("month").replace(".", "").lower()
    month = month_map.get(month_key)
    if month is None:
        return pd.to_datetime(text, errors="coerce")

    hour = int(match.group("hour"))
    ampm = match.group("ampm").upper()
    if ampm == "AM":
        hour = 0 if hour in (0, 12) else hour
    else:
        hour = 12 if hour in (0, 12) else hour + 12

    try:
        return pd.Timestamp(
            year=int(match.group("year")),
            month=month,
            day=int(match.group("day")),
            hour=hour,
            minute=int(match.group("minute")),
            second=int(match.group("second")),
        )
    except ValueError:
        return pd.NaT


def load_meteoswiss_source(file_path: str | Path) -> pd.DataFrame:
    """Load one MeteoSwiss 10-minute precipitation source."""
    df = pd.read_csv(file_path, sep=None, engine="python")
    required = {"reference_timestamp", "rre150z0"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    df = df[["reference_timestamp", "rre150z0"]].copy()
    df.columns = ["DateTime", "P_mm_10min"]
    df["DateTime"] = pd.to_datetime(df["DateTime"], dayfirst=True, errors="coerce")
    df["P_mm_10min"] = (
        df["P_mm_10min"]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .pipe(pd.to_numeric, errors="coerce")
    )

    df = df.dropna(subset=["DateTime", "P_mm_10min"])
    df = df[df["P_mm_10min"] >= 0]
    df = df.sort_values("DateTime").drop_duplicates(subset=["DateTime"], keep="first")
    return df.reset_index(drop=True)


def load_xlsx_precipitation(file_path: str | Path) -> pd.DataFrame:
    """Read a processed XLSX file using openpyxl streaming mode."""
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    worksheet = workbook.active

    rows = worksheet.iter_rows(min_row=7, values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        workbook.close()
        return pd.DataFrame(columns=["DateTime", "Precipitation"])

    columns = [str(value).strip() if value is not None else "" for value in header]
    lowered = [col.lower() for col in columns]

    date_idx = next(
        (idx for idx, col in enumerate(lowered) if "date" in col or "time" in col),
        0,
    )
    time_idx = next(
        (idx for idx, col in enumerate(lowered) if col in {"heure", "time"}),
        None,
    )
    precip_idx = next(
        (
            idx
            for idx, col in enumerate(lowered)
            if any(key in col for key in ["precip", "pulse", "pulsation"])
        ),
        len(columns) - 1,
    )

    datetimes = []
    precipitations = []
    for row in rows:
        date_value = row[date_idx] if date_idx < len(row) else None
        precip_value = row[precip_idx] if precip_idx < len(row) else None

        # Some exports split date and time into two columns.
        if time_idx is not None and time_idx < len(row) and date_value is not None:
            time_value = row[time_idx]
            if isinstance(date_value, datetime) and isinstance(time_value, datetime):
                if date_value.time() == time(0, 0):
                    date_value = datetime.combine(date_value.date(), time_value.time())
            elif isinstance(date_value, datetime) and isinstance(time_value, time):
                if date_value.time() == time(0, 0):
                    date_value = datetime.combine(date_value.date(), time_value)

        datetimes.append(date_value)
        precipitations.append(precip_value)

    workbook.close()
    return pd.DataFrame({"DateTime": datetimes, "Precipitation": precipitations})


def load_processed_file(file_path: str | Path) -> pd.DataFrame:
    """Load one processed station file as DateTime + P_mm."""
    file_path = Path(file_path)
    extension = file_path.suffix.lower()

    if extension == ".txt":
        df = pd.read_csv(file_path, sep='\t', skiprows=5, header=None, usecols=[0, 1, 2],
                         names=['Index', 'DateTime', 'Precipitation'])
        df['DateTime'] = df['DateTime'].apply(parse_station_datetime)
        df['Precipitation'] = df['Precipitation'].apply(extract_numeric)

    elif extension == ".xlsx":
        if file_path.name.startswith("~$"):
            return pd.DataFrame(columns=["DateTime", "P_mm"])
        df = load_xlsx_precipitation(file_path)
        df["DateTime"] = df["DateTime"].apply(parse_station_datetime)

    elif extension == ".csv":
        df = pd.read_csv(file_path)
        if 'created_at' in df.columns and 'field5' in df.columns:
            df = df[['created_at', 'field5']].copy()
            df.columns = ['DateTime', 'Precipitation']
            df['DateTime'] = pd.to_datetime(df['DateTime'], utc=True).dt.tz_localize(None)
            df['Precipitation'] = df['Precipitation'].apply(extract_numeric)
        elif {"DateTime", "P_mm"}.issubset(df.columns):
            df = df[["DateTime", "P_mm"]].copy()
            df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
            df["Precipitation"] = df["P_mm"]
        else:
            raise ValueError("CSV format not recognized: expected created_at/field5 or DateTime/P_mm")

    else:
        raise ValueError(f"Unsupported extension: {extension}")

    df = df.dropna(subset=["DateTime", "Precipitation"]).copy()
    df = df[df["Precipitation"] >= 0]

    if "P_mm" not in df.columns or extension != ".csv":
        df['P_mm'] = df['Precipitation'] * 0.01

    df = df[["DateTime", "P_mm"]]
    df = df.sort_values("DateTime")
    df = df.groupby("DateTime", as_index=False)["P_mm"].sum()
    return df.reset_index(drop=True)


def infer_timestep_minutes(series_df: pd.DataFrame) -> float | None:
    """Infer the median time step in minutes from one station file."""
    diffs = series_df["DateTime"].sort_values().diff().dropna()
    if diffs.empty:
        return None

    timestep = diffs.median().total_seconds() / 60
    if not np.isfinite(timestep) or timestep <= 0:
        return None

    rounded = round(timestep)
    if abs(timestep - rounded) < 0.05:
        return float(rounded)
    return float(timestep)


def regularize_station_series(series_df: pd.DataFrame, timestep_minutes: float) -> pd.Series:
    """
    Place the station file on a complete regular grid at its own time step.

    Some CSV timestamps drift by a few seconds over long periods. Instead of
    requiring exact equality with a perfect grid, each measurement is assigned
    to the nearest expected time step from the first timestamp.
    """
    df = series_df.sort_values("DateTime").set_index("DateTime")
    timestep = pd.Timedelta(minutes=timestep_minutes)
    start = df.index.min()

    # Divide timedeltas directly so pandas version changes in integer units
    # (ns/us) do not affect the grid assignment.
    grid_positions = np.rint((df.index - start) / timestep).astype(int)
    grid_times = start + pd.to_timedelta(grid_positions * timestep.value, unit="ns")

    gridded = (
        df.assign(GridTime=grid_times)
        .groupby("GridTime")["P_mm"]
        .sum()
        .sort_index()
    )
    target_index = pd.date_range(start=start, end=gridded.index.max(), freq=timestep)
    regularized = gridded.reindex(target_index).astype(float)
    if len(series_df) >= MIN_VALID_PAIRS and len(regularized) < len(series_df) * 0.5:
        raise ValueError(
            "Regularized station series is unexpectedly short "
            f"({len(regularized)} grid rows for {len(series_df)} source rows). "
            "This usually means an old timestamp-regularization function is being used. "
            "Restart the kernel and run the repository Stage_6_cross_correlation.py file."
        )
    return regularized


def distribute_meteo_to_target_grid(
    meteo_df: pd.DataFrame,
    target_index: pd.DatetimeIndex,
    target_step_minutes: float,
) -> pd.Series:
    """
    Distribute 10-minute MeteoSwiss totals onto a target grid.

    MeteoSwiss and target timestamps are treated as interval end times:
    - source interval = (meteo_timestamp - 10 min, meteo_timestamp]
    - target interval = (target_timestamp - target_step, target_timestamp]

    Each 10-minute total is distributed to target bins proportionally to the
    overlap duration. This conserves MeteoSwiss rainfall totals and matches the
    convention used by the regular station grids.
    """
    target_index = pd.DatetimeIndex(target_index)
    if len(target_index) == 0:
        return pd.Series(dtype=float)

    target_step = pd.Timedelta(minutes=target_step_minutes)
    source_step = pd.Timedelta(minutes=METEOSWISS_STEP_MINUTES)
    source_step_ns = source_step.value

    target_ends = target_index
    target_starts = target_ends - target_step
    target_start_min = target_starts.min()
    target_end_max = target_ends.max()

    meteo = meteo_df.copy()
    meteo["interval_start"] = meteo["DateTime"] - source_step
    meteo["interval_end"] = meteo["DateTime"]
    meteo = meteo[
        (meteo["interval_end"] > target_start_min)
        & (meteo["interval_start"] < target_end_max)
    ]

    if meteo.empty:
        return pd.Series(np.full(len(target_index), np.nan), index=target_index)

    source_starts = meteo["interval_start"].values.astype("datetime64[ns]").astype("int64")
    source_ends = meteo["interval_end"].values.astype("datetime64[ns]").astype("int64")
    source_values = meteo["P_mm_10min"].to_numpy(dtype=float)

    target_starts_ns = target_starts.values.astype("datetime64[ns]").astype("int64")
    target_ends_ns = target_ends.values.astype("datetime64[ns]").astype("int64")

    values = np.zeros(len(target_index), dtype=float)
    has_data = np.zeros(len(target_index), dtype=bool)

    for pos, (target_start_ns, target_end_ns) in enumerate(zip(target_starts_ns, target_ends_ns)):
        first_source = np.searchsorted(source_ends, target_start_ns, side="right")
        last_source = np.searchsorted(source_starts, target_end_ns, side="left")

        for source_pos in range(first_source, last_source):
            overlap_start_ns = max(target_start_ns, source_starts[source_pos])
            overlap_end_ns = min(target_end_ns, source_ends[source_pos])
            overlap_ns = max(0, overlap_end_ns - overlap_start_ns)
            if overlap_ns > 0:
                values[pos] += source_values[source_pos] * (overlap_ns / source_step_ns)
                has_data[pos] = True

    values[~has_data] = np.nan
    return pd.Series(values, index=target_index, dtype=float)


def pearson_corr_arrays(station_values: np.ndarray, meteo_values: np.ndarray) -> tuple[float, int]:
    """Compute Pearson correlation for two same-length arrays with NaN filtering."""
    mask = np.isfinite(station_values) & np.isfinite(meteo_values)
    n_pairs = int(mask.sum())
    if n_pairs < MIN_VALID_PAIRS:
        return np.nan, n_pairs

    x = station_values[mask]
    y = meteo_values[mask]
    x_centered = x - x.mean()
    y_centered = y - y.mean()
    denominator = np.sqrt(np.sum(x_centered ** 2) * np.sum(y_centered ** 2))
    if denominator == 0:
        return np.nan, n_pairs

    return float(np.sum(x_centered * y_centered) / denominator), n_pairs


def compute_best_cross_correlation(
    station_series: pd.Series,
    meteo_series: pd.Series,
    timestep_minutes: float,
) -> dict:
    """Find the lag with the largest absolute Pearson correlation."""
    max_lag_steps = int(np.floor(PROCESSED_FILE_MAX_LAG_MINUTES / timestep_minutes))
    station_values = station_series.to_numpy(dtype=float)
    meteo_values = meteo_series.to_numpy(dtype=float)
    rows = []

    for lag_steps in range(-max_lag_steps, max_lag_steps + 1):
        if lag_steps > 0:
            station_slice = station_values[lag_steps:]
            meteo_slice = meteo_values[:-lag_steps]
        elif lag_steps < 0:
            station_slice = station_values[:lag_steps]
            meteo_slice = meteo_values[-lag_steps:]
        else:
            station_slice = station_values
            meteo_slice = meteo_values

        corr, n_pairs = pearson_corr_arrays(station_slice, meteo_slice)
        rows.append(
            {
                "lag_steps": lag_steps,
                "lag_minutes": lag_steps * timestep_minutes,
                "correlation": corr,
                "n_valid_pairs": n_pairs,
            }
        )

    lag_df = pd.DataFrame(rows).dropna(subset=["correlation"])
    if lag_df.empty:
        return {
            "best_lag_steps": np.nan,
            "best_lag_minutes": np.nan,
            "best_correlation": np.nan,
            "best_abs_correlation": np.nan,
            "n_valid_pairs_at_best_lag": 0,
            "_lag_df": lag_df,
        }

    best = lag_df.loc[lag_df["correlation"].abs().idxmax()]
    return {
        "best_lag_steps": int(best["lag_steps"]),
        "best_lag_minutes": float(best["lag_minutes"]),
        "best_correlation": float(best["correlation"]),
        "best_abs_correlation": float(abs(best["correlation"])),
        "n_valid_pairs_at_best_lag": int(best["n_valid_pairs"]),
        "_lag_df": lag_df,
    }


def safe_filename(value: str) -> str:
    """Create a filesystem-safe filename fragment."""
    return "".join(char if char.isalnum() or char in "._-" else "_" for char in value)


def apply_report_time_axis(ax) -> None:
    """Format time ticks consistently with Stage 3 overlap figures."""
    locator = mdates.AutoDateLocator(minticks=5, maxticks=9)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.DateFormatter(REPORT_DATE_TICK_FORMAT))
    ax.tick_params(axis="x", rotation=0)


def build_lagged_pair(common: pd.DataFrame, lag_steps: int) -> pd.DataFrame:
    """Return station and MeteoSwiss values paired at the requested lag."""
    if lag_steps > 0:
        station_values = common["station"].iloc[lag_steps:].to_numpy()
        meteo_values = common["meteo"].iloc[:-lag_steps].to_numpy()
    elif lag_steps < 0:
        station_values = common["station"].iloc[:lag_steps].to_numpy()
        meteo_values = common["meteo"].iloc[-lag_steps:].to_numpy()
    else:
        station_values = common["station"].to_numpy()
        meteo_values = common["meteo"].to_numpy()

    pair_df = pd.DataFrame({"Station": station_values, "MeteoSwiss": meteo_values})
    return pair_df.dropna()


def save_diagnostic_figure(
    common: pd.DataFrame,
    lag_df: pd.DataFrame,
    best: dict,
    meteo_station: str,
    station: str,
    file_name: str,
    output_directory: Path,
) -> str | None:
    """Save a PNG with time series, lag-correlation curve and scatter plot."""
    if lag_df.empty or pd.isna(best.get("best_lag_steps")):
        return None

    figures_dir = output_directory / "figures" / station / meteo_station
    figures_dir.mkdir(parents=True, exist_ok=True)

    best_lag_steps = int(best["best_lag_steps"])
    best_lag_minutes = float(best["best_lag_minutes"])
    best_correlation = float(best["best_correlation"])
    pair_df = build_lagged_pair(common, best_lag_steps)
    if pair_df.empty:
        return None

    scatter_df = (
        pair_df.sample(MAX_SCATTER_POINTS, random_state=42)
        if len(pair_df) > MAX_SCATTER_POINTS
        else pair_df
    )

    figure, axes = plt.subplots(3, 1, figsize=(11, 10.5))
    figure.subplots_adjust(hspace=0.55)

    axes[0].plot(common.index, common["station"], label="Pluvimate station", color="#1f77b4", linewidth=0.8)
    axes[0].plot(common.index, common["meteo"], label=f"MeteoSwiss {meteo_station}", color="#d62728", linewidth=0.8, alpha=0.8)
    axes[0].set_ylabel("Precipitation (mm)")
    axes[0].set_title("Precipitation time series")
    axes[0].legend(loc="upper right")
    axes[0].grid(True, alpha=0.3)
    apply_report_time_axis(axes[0])

    axes[1].plot(lag_df["lag_minutes"], lag_df["correlation"], color="#333333", linewidth=1.0)
    axes[1].axvline(best_lag_minutes, color="#d62728", linestyle="--", linewidth=1.0)
    axes[1].scatter([best_lag_minutes], [best_correlation], color="#d62728", zorder=3)
    axes[1].set_xlabel("Lag (minutes)")
    axes[1].set_ylabel("Pearson correlation")
    axes[1].set_title("Correlation by lag")
    axes[1].grid(True, alpha=0.3)

    axes[2].scatter(
        scatter_df["MeteoSwiss"],
        scatter_df["Station"],
        s=8,
        alpha=0.35,
        color="#2ca02c",
        edgecolors="none",
    )
    max_axis = np.nanmax([scatter_df["MeteoSwiss"].max(), scatter_df["Station"].max()])
    if np.isfinite(max_axis) and max_axis > 0:
        axes[2].plot([0, max_axis], [0, max_axis], color="#555555", linestyle=":", linewidth=1.0)
        axes[2].set_xlim(left=0, right=max_axis * 1.05)
        axes[2].set_ylim(bottom=0, top=max_axis * 1.05)
    axes[2].set_xlabel(f"MeteoSwiss {meteo_station} precipitation (mm)")
    axes[2].set_ylabel("Pluvimate precipitation (mm)")
    axes[2].set_title("Scatter plot at best lag")
    axes[2].grid(True, alpha=0.3)

    figure_path = figures_dir / f"{safe_filename(Path(file_name).stem)}_cross_correlation.png"
    figure.savefig(figure_path, dpi=FIGURE_DPI)
    plt.close(figure)
    return str(figure_path)


def save_combined_interactive_time_series(
    plot_infos: list[dict],
    station: str,
    file_name: str,
    output_directory: Path,
) -> str | None:
    """Save one interactive HTML figure per processed file, with one panel per MeteoSwiss station."""
    valid_infos = [
        info for info in plot_infos
        if info.get("common") is not None and not pd.isna(info.get("best", {}).get("best_lag_steps"))
    ]
    if go is None or make_subplots is None or not valid_infos:
        return None

    figures_dir = output_directory / "figures" / station / "combined_meteoswiss"
    figures_dir.mkdir(parents=True, exist_ok=True)

    rows = len(valid_infos)
    subplot_titles = []
    for info in valid_infos:
        best = info["best"]
        subplot_titles.append(
            f"{info['meteo_station']}: best lag = {best['best_lag_minutes']:.1f} min, "
            f"r = {best['best_correlation']:.3f}"
        )

    figure = make_subplots(
        rows=rows,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.12,
        subplot_titles=subplot_titles,
    )

    for row, info in enumerate(valid_infos, start=1):
        common = info["common"]
        meteo_station = info["meteo_station"]
        showlegend = row == 1

        figure.add_trace(
            go.Scatter(
                x=common.index,
                y=common["station"],
                mode="lines",
                name="Pluvimate station",
                legendgroup="station",
                showlegend=showlegend,
                line=dict(color="#1f77b4", width=4),
            ),
            row=row,
            col=1,
        )
        figure.add_trace(
            go.Scatter(
                x=common.index,
                y=common["meteo"],
                mode="lines",
                name=f"MeteoSwiss {meteo_station}",
                legendgroup=f"meteo_{meteo_station}",
                showlegend=True,
                line=dict(color="#d62728", width=4),
                opacity=0.65,
            ),
            row=row,
            col=1,
        )
        figure.update_yaxes(title_text="Precipitation (mm)", row=row, col=1)

    figure.update_layout(
        title=dict(
            text=f"{station} - {file_name}",
            x=0.0,
            xanchor="left",
            y=0.99,
            yanchor="top",
            font=dict(size=INTERACTIVE_TITLE_FONT_SIZE),
        ),
        hovermode="x unified",
        dragmode="zoom",
        template="plotly_white",
        font=dict(size=INTERACTIVE_AXIS_TICK_FONT_SIZE),
        legend=dict(
            font=dict(size=INTERACTIVE_LEGEND_FONT_SIZE),
            title_font=dict(size=INTERACTIVE_LEGEND_FONT_SIZE),
            orientation="h",
            itemsizing="constant",
            yanchor="top",
            y=-0.16,
            xanchor="left",
            x=0,
            bgcolor="rgba(255,255,255,0.85)",
            borderwidth=0,
        ),
        margin=dict(l=120, r=45, t=170, b=210),
        width=INTERACTIVE_FIGURE_WIDTH,
        height=max(620 * rows, 1500),
    )
    for annotation in figure.layout.annotations:
        annotation.font.size = INTERACTIVE_SUBTITLE_FONT_SIZE
        annotation.y += 0.015

    figure.update_xaxes(
        title_font=dict(size=INTERACTIVE_AXIS_TITLE_FONT_SIZE),
        tickfont=dict(size=INTERACTIVE_AXIS_TICK_FONT_SIZE),
        rangeslider_visible=False,
    )
    figure.update_xaxes(
        title_text="Time",
        title_standoff=INTERACTIVE_XAXIS_TITLE_STANDOFF,
        tickformat=INTERACTIVE_DATE_TICK_FORMAT,
        row=rows,
        col=1,
        rangeslider_visible=False,
    )
    figure.update_yaxes(
        title_font=dict(size=INTERACTIVE_AXIS_TITLE_FONT_SIZE),
        tickfont=dict(size=INTERACTIVE_AXIS_TICK_FONT_SIZE),
        fixedrange=False,
    )

    html_path = figures_dir / f"{safe_filename(Path(file_name).stem)}_meteoswiss_comparison_interactive.html"
    figure.write_html(
        html_path,
        include_plotlyjs="cdn",
        config={
            "scrollZoom": True,
            "displaylogo": False,
            "modeBarButtonsToAdd": ["zoomIn2d", "zoomOut2d", "resetScale2d"],
        },
    )
    return str(html_path)


def save_combined_report_png(
    plot_infos: list[dict],
    station: str,
    file_name: str,
    output_directory: Path,
) -> str | None:
    """Save a report-ready PNG matching the combined interactive figure layout."""
    valid_infos = [
        info for info in plot_infos
        if info.get("common") is not None and not pd.isna(info.get("best", {}).get("best_lag_steps"))
    ]
    if not valid_infos:
        return None

    figures_dir = output_directory / "figures" / station / "combined_meteoswiss"
    figures_dir.mkdir(parents=True, exist_ok=True)

    rows = len(valid_infos)
    figure_height = max(4.4 * rows, 8.8)
    figure, axes = plt.subplots(
        rows,
        1,
        figsize=(16, figure_height),
        sharex=True,
    )
    figure.subplots_adjust(hspace=0.42)
    if rows == 1:
        axes = [axes]

    for ax, info in zip(axes, valid_infos):
        common = info["common"]
        best = info["best"]
        meteo_station = info["meteo_station"]
        best_lag_minutes = float(best["best_lag_minutes"])
        best_correlation = float(best["best_correlation"])

        ax.plot(
            common.index,
            common["station"],
            color="#1f77b4",
            linewidth=1.4,
            label="Pluvimate station",
        )
        ax.plot(
            common.index,
            common["meteo"],
            color="#d62728",
            linewidth=1.4,
            alpha=0.75,
            label=f"MeteoSwiss {meteo_station}",
        )
        ax.set_title(
            f"{meteo_station}: best lag = {best_lag_minutes:.1f} min, r = {best_correlation:.3f}",
            fontsize=REPORT_SUBTITLE_FONTSIZE,
        )
        ax.set_ylabel("Precipitation (mm)", fontsize=REPORT_LABEL_FONTSIZE)
        ax.tick_params(axis="both", labelsize=REPORT_TICK_FONTSIZE)
        ax.grid(True, alpha=0.3)
        ax.legend(loc="upper right", fontsize=REPORT_LEGEND_FONTSIZE)

    axes[-1].set_xlabel("Time", fontsize=REPORT_LABEL_FONTSIZE)
    apply_report_time_axis(axes[-1])

    png_path = figures_dir / f"{safe_filename(Path(file_name).stem)}_meteoswiss_comparison_report.png"
    figure.savefig(png_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(figure)
    return str(png_path)


def list_processed_files(base_directory: Path, station_filter: str | None = None) -> list[Path]:
    """Return all supported processed files in station folders."""
    patterns = ["*.txt", "*.csv", "*.xlsx"]
    files = []

    if station_filter:
        station_directories = [base_directory / station_filter]
    else:
        station_directories = sorted(path for path in base_directory.iterdir() if path.is_dir())

    for station_directory in station_directories:
        if not station_directory.is_dir():
            continue
        for pattern in patterns:
            files.extend(station_directory.glob(pattern))

    return sorted(path for path in files if not path.name.startswith("~$"))


def compare_processed_file_to_meteoswiss(
    file_path: Path,
    station_df: pd.DataFrame,
    station_series: pd.Series,
    timestep_minutes: float,
    meteo_station: str,
    meteo_df: pd.DataFrame,
) -> tuple[dict, dict | None]:
    """Compare one processed file to one MeteoSwiss reference station."""
    station = file_path.parent.name
    file_name = file_path.name

    base_result = {
        "MeteoSwiss_Station": meteo_station,
        "Station": station,
        "File": file_name,
        "File_path": str(file_path),
        "Status": "ok",
        "Time_step_min": timestep_minutes,
        "First_file_timestamp": station_df["DateTime"].min(),
        "Last_file_timestamp": station_df["DateTime"].max(),
        "N_station_rows": len(station_df),
        "N_station_grid_rows": len(station_series),
    }

    meteo_series = distribute_meteo_to_target_grid(meteo_df, station_series.index, timestep_minutes)
    common = pd.concat([station_series.rename("station"), meteo_series.rename("meteo")], axis=1)
    common = common.dropna(how="all")
    common_pairs = common.dropna()
    n_overlap_pairs_no_lag = len(common_pairs)

    if n_overlap_pairs_no_lag < MIN_VALID_PAIRS:
        return {
            **base_result,
            "Status": "not_enough_overlap",
            "First_common": common.index.min() if not common.empty else pd.NaT,
            "Last_common": common.index.max() if not common.empty else pd.NaT,
            "N_overlap_pairs_no_lag": n_overlap_pairs_no_lag,
        }, None

    best = compute_best_cross_correlation(
        common["station"],
        common["meteo"],
        timestep_minutes,
    )
    lag_df = best.pop("_lag_df", pd.DataFrame())

    if pd.isna(best.get("best_correlation")):
        return {
            **base_result,
            "Status": "no_valid_correlation",
            "First_common": common_pairs.index.min(),
            "Last_common": common_pairs.index.max(),
            "N_overlap_pairs_no_lag": n_overlap_pairs_no_lag,
            **best,
        }, None

    result = {
        **base_result,
        "First_common": common_pairs.index.min(),
        "Last_common": common_pairs.index.max(),
        "N_overlap_pairs_no_lag": n_overlap_pairs_no_lag,
        **best,
    }
    plot_info = {
        "meteo_station": meteo_station,
        "common": common,
        "best": best,
        "lag_df": lag_df,
    }
    return result, plot_info


def run_meteoswiss_processed_correlation():
    """Run MeteoSwiss vs processed station-file correlation."""
    print("\n" + "="*80)
    print("METEOSWISS VS PROCESSED STATION FILES")
    print("="*80 + "\n")

    selected_station = STATION_TO_PROCESS if STATION_TO_PROCESS is not None else "all stations"
    reference_stations = ", ".join(REFERENCE_METEOSWISS_STATIONS.keys())

    print(f"Processed files: {PROCESSED_DIR}")
    print(f"MeteoSwiss references: {reference_stations}")
    print(f"Output folder: {METEOSWISS_PROCESSED_OUTPUT_DIR}")
    print(f"Station selection: {selected_station}")
    print(f"Maximum lag: +/- {PROCESSED_FILE_MAX_LAG_MINUTES} minutes")
    print(f"PNG diagnostic figures: {CREATE_DIAGNOSTIC_FIGURES}")
    print(f"Interactive HTML figures: {CREATE_INTERACTIVE_FIGURES}")
    print(f"Report PNG figures: {CREATE_REPORT_PNG_FIGURES}")

    METEOSWISS_PROCESSED_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    meteo_data = {
        station: load_meteoswiss_source(file_path)
        for station, file_path in REFERENCE_METEOSWISS_STATIONS.items()
    }
    for station, meteo_df in meteo_data.items():
        print(f"Loaded MeteoSwiss {station}: {len(meteo_df)} rows")

    processed_files = list_processed_files(PROCESSED_DIR, station_filter=STATION_TO_PROCESS)
    print(f"Processed files found: {len(processed_files)}")

    if not processed_files:
        raise FileNotFoundError(
            f"No processed files found for station={STATION_TO_PROCESS!r} in {PROCESSED_DIR}"
        )

    results = []
    for idx, file_path in enumerate(processed_files, start=1):
        station = file_path.parent.name
        file_name = file_path.name
        file_results = []
        plot_infos = []

        try:
            station_df = load_processed_file(file_path)
            if len(station_df) < MIN_VALID_PAIRS:
                file_results = [
                    {
                        "MeteoSwiss_Station": meteo_station,
                        "Station": station,
                        "File": file_name,
                        "File_path": str(file_path),
                        "Status": "not_enough_station_data",
                    }
                    for meteo_station in meteo_data
                ]
            else:
                timestep_minutes = infer_timestep_minutes(station_df)
                if timestep_minutes is None:
                    file_results = [
                        {
                            "MeteoSwiss_Station": meteo_station,
                            "Station": station,
                            "File": file_name,
                            "File_path": str(file_path),
                            "Status": "could_not_infer_timestep",
                        }
                        for meteo_station in meteo_data
                    ]
                else:
                    station_series = regularize_station_series(station_df, timestep_minutes)
                    for meteo_station, meteo_df in meteo_data.items():
                        result, plot_info = compare_processed_file_to_meteoswiss(
                            file_path=file_path,
                            station_df=station_df,
                            station_series=station_series,
                            timestep_minutes=timestep_minutes,
                            meteo_station=meteo_station,
                            meteo_df=meteo_df,
                        )
                        if plot_info is not None:
                            if CREATE_DIAGNOSTIC_FIGURES:
                                result["Figure_path"] = save_diagnostic_figure(
                                    common=plot_info["common"],
                                    lag_df=plot_info["lag_df"],
                                    best=plot_info["best"],
                                    meteo_station=meteo_station,
                                    station=station,
                                    file_name=file_name,
                                    output_directory=METEOSWISS_PROCESSED_OUTPUT_DIR,
                                )
                            plot_infos.append(plot_info)
                        file_results.append(result)

                    if CREATE_INTERACTIVE_FIGURES:
                        combined_html = save_combined_interactive_time_series(
                            plot_infos=plot_infos,
                            station=station,
                            file_name=file_name,
                            output_directory=METEOSWISS_PROCESSED_OUTPUT_DIR,
                        )
                        for result in file_results:
                            result["Combined_interactive_figure_path"] = combined_html

                    if CREATE_REPORT_PNG_FIGURES:
                        save_combined_report_png(
                            plot_infos=plot_infos,
                            station=station,
                            file_name=file_name,
                            output_directory=METEOSWISS_PROCESSED_OUTPUT_DIR,
                        )

        except Exception as exc:
            file_results = [
                {
                    "MeteoSwiss_Station": meteo_station,
                    "Station": station,
                    "File": file_name,
                    "File_path": str(file_path),
                    "Status": "error",
                    "Error": f"{type(exc).__name__}: {exc}",
                }
                for meteo_station in meteo_data
            ]

        results.extend(file_results)
        status_summary = ", ".join(
            f"{result['MeteoSwiss_Station']}={result['Status']}" for result in file_results
        )
        print(f"[{idx:03d}/{len(processed_files):03d}] {station} / {file_name}: {status_summary}", flush=True)

    result_df = pd.DataFrame(results)
    preferred_columns = [
        "MeteoSwiss_Station",
        "Station",
        "File",
        "Status",
        "Time_step_min",
        "First_file_timestamp",
        "Last_file_timestamp",
        "First_common",
        "Last_common",
        "N_station_rows",
        "N_station_grid_rows",
        "N_overlap_pairs_no_lag",
        "best_lag_steps",
        "best_lag_minutes",
        "best_correlation",
        "best_abs_correlation",
        "n_valid_pairs_at_best_lag",
        "Error",
    ]
    result_df = result_df.reindex(columns=[col for col in preferred_columns if col in result_df.columns])

    suffix = f"_{STATION_TO_PROCESS}" if STATION_TO_PROCESS else ""
    csv_output = METEOSWISS_PROCESSED_OUTPUT_DIR / f"{PROCESSED_CORRELATION_SUMMARY_BASENAME}{suffix}.csv"
    xlsx_output = METEOSWISS_PROCESSED_OUTPUT_DIR / f"{PROCESSED_CORRELATION_SUMMARY_BASENAME}{suffix}.xlsx"

    result_df.to_csv(csv_output, index=False)
    with pd.ExcelWriter(xlsx_output, engine="openpyxl") as writer:
        for meteo_station in REFERENCE_METEOSWISS_STATIONS:
            sheet_df = result_df[result_df["MeteoSwiss_Station"] == meteo_station].copy()
            sheet_df.to_excel(writer, sheet_name=meteo_station, index=False)

    format_excel_report(xlsx_output)

    print(f"\nCSV exported:  {csv_output}")
    print(f"XLSX exported: {xlsx_output}")


# ============================================================================
# LEXPLORE NETCDF VS PROCESSED FILES
# ============================================================================

# %% Lexplore functions

def run_lexplore_processed_correlation():
    """Run Lexplore NetCDF vs processed station-file correlation."""
    print("\n" + "="*80)
    print("LEXPLORE NETCDF VS PROCESSED STATION FILES")
    print("="*80 + "\n")

    print(f"Lexplore NetCDF folder: {LEXPLORE_NC_DIR}")
    print(f"Processed files: {PROCESSED_DIR}")
    print(f"Output folder: {LEXPLORE_OUTPUT_DIR}")
    print(f"Station selection: {LEXPLORE_STATION_TO_PROCESS}")
    print(f"Rain variable: {LEXPLORE_RAIN_VARIABLE}")
    print(f"Quality variable: {LEXPLORE_RAIN_QUALITY_VARIABLE}")
    print(f"Time variable: {LEXPLORE_TIME_VARIABLE}")
    print(f"Reference source: {LEXPLORE_SOURCE_NAME}")
    print(f"Reference export file: {METEOSWISS_DIR / LEXPLORE_METEOSWISS_EXPORT_FILE}")
    print(f"Lexplore rain CSV: {LEXPLORE_OUTPUT_DIR / LEXPLORE_NC_SERIES_FILE}")

    # Placeholder kept until the Lexplore-specific workflow is migrated here.
    print("\nThis check is reserved for the Lexplore-specific workflow.")
    print("No files were modified.")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    run_meteoswiss_station_correlation()
    run_meteoswiss_processed_correlation()

    # Optional Lexplore NetCDF vs processed files
    # run_lexplore_processed_correlation()
