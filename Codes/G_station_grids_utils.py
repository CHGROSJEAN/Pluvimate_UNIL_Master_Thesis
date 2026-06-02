# Charlotte Grosjean - 26.05.2026 - UNIL Master Thesis
"""
STATION GRIDS UTILITIES

Consolidated utility functions for building and processing station availability and precipitation grids.

This module supports Stage_5_station_grids.py by centralizing the grid-building
logic for availability, precipitation, gap codes, summaries, and heatmaps.

Main functions:
- load_consolidated_station_data(): Load all consolidated CSV files
- load_consolidated_station_file(): Load individual station file
- infer_source_steps_minutes(): Infer temporal resolution from data
- classify_gap_code(): Mark inter/intra-file gaps
- make_global_index(): Create shared 3-minute time grid
- distribute_station_to_global_grid(): Redistribute station values to grid
- build_station_grids(): Build complete availability/precipitation/gap grids
"""

from __future__ import annotations

import os
import tempfile
from collections import OrderedDict
from glob import glob
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
import plotly.graph_objects as go
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from A_data_processing_utils import extract_numeric
from F_rain_quality_plot_utils import RAIN_COLORSCALE
from I_paths_config import (
    AVAILABILITY_DIRECTORY,
    CONSOLIDATED_DIRECTORY,
    DEFAULT_SOURCE_STEP_MINUTES,
    GRID_STEP_MINUTES,
    PRECIPITATION_DIRECTORY,
)

OUTPUT_DIRECTORY_AVAILABILITY = AVAILABILITY_DIRECTORY
OUTPUT_DIRECTORY_RAIN = PRECIPITATION_DIRECTORY


# ===== HELPER FUNCTIONS: Input loading and preprocessing =====

def load_consolidated_station_data() -> OrderedDict:
    """
    Load all consolidated station files from 2_Consolidated_Stations.

    Returns:
        OrderedDict: Maps station names to DataFrames or None if loading failed
    """
    print("[INFO] Searching for consolidated station files...")
    station_files = sorted(
        glob(os.path.join(CONSOLIDATED_DIRECTORY, "*_Consolidated_Complete.csv"))
    )
    print(f"[INFO] Found {len(station_files)} station files.")

    print("[INFO] Loading station data...")
    station_data = OrderedDict()
    for station_file in station_files:
        station_name, df = load_consolidated_station_file(station_file)
        station_data[station_name] = df
    return station_data


def load_consolidated_station_file(file_path: str) -> tuple[str, pd.DataFrame | None]:
    """
    Load one consolidated station file and keep Source_File for timestep inference.

    Args:
        file_path: Path to consolidated CSV (e.g., BETH_Consolidated_Complete.csv)

    Returns:
        tuple: (station_name, DataFrame or None)
    """
    station_name = os.path.basename(file_path).replace("_Consolidated_Complete.csv", "")

    try:
        df = pd.read_csv(file_path)
    except Exception as exc:
        print(f"Warning: Could not load consolidated CSV {file_path}: {exc}")
        return station_name, None

    if "DateTime" not in df.columns:
        print(f"Warning: 'DateTime' column missing in {file_path}.")
        return station_name, None

    df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")

    if "P_drops" in df.columns:
        df["P_drops"] = df["P_drops"].apply(extract_numeric)
    else:
        df["P_drops"] = np.nan

    if "P_mm" not in df.columns:
        df["P_mm"] = df["P_drops"].apply(
            lambda value: value * 0.01 if pd.notna(value) and value >= 0 else value
        )
    else:
        df["P_mm"] = df["P_mm"].apply(extract_numeric)

    if "Source_File" not in df.columns:
        df["Source_File"] = station_name

    df = (
        df[["Source_File", "DateTime", "P_drops", "P_mm"]]
        .dropna(subset=["DateTime"])
        .sort_values("DateTime")
        .drop_duplicates(subset=["DateTime"], keep="first")
        .reset_index(drop=True)
    )
    return station_name, df


def infer_source_steps_minutes(df: pd.DataFrame) -> pd.Series:
    """
    Infer the accumulation duration for each row, preferably by Source_File.

    A 5-minute source file gets a 5-minute interval; Pluvimate files keep 3-minute intervals.
    Gap rows are kept at the default 3-minute duration.

    Args:
        df: DataFrame with DateTime, P_mm, Source_File columns

    Returns:
        Series: Inferred step (in minutes) for each row
    """
    row_steps = pd.Series(DEFAULT_SOURCE_STEP_MINUTES, index=df.index, dtype=float)
    if df.empty:
        return row_steps

    for _, group in df.sort_values("DateTime").groupby("Source_File", dropna=False):
        real_measurements = group[group["P_mm"].ge(0)]
        diffs = real_measurements["DateTime"].sort_values().diff().dt.total_seconds() / 60
        diffs = diffs[np.isfinite(diffs) & (diffs > 0)]
        if diffs.empty:
            step_minutes = DEFAULT_SOURCE_STEP_MINUTES
        else:
            step_minutes = float(diffs.median())
            if not np.isfinite(step_minutes) or step_minutes <= 0:
                step_minutes = DEFAULT_SOURCE_STEP_MINUTES
        row_steps.loc[group.index] = step_minutes

    gap_mask = df["P_mm"].lt(0) | df["P_drops"].lt(0)
    row_steps.loc[gap_mask] = DEFAULT_SOURCE_STEP_MINUTES
    return row_steps


def classify_gap_code(p_mm: float, p_drops: float) -> float:
    """
    Return -1 or -2 for gap rows, NaN for real measurements.

    Gap codes:
    - -2: intra-file gap (missing measurement within one source file)
    - -1: inter-file gap (gap between different source files)
    """
    candidates = [p_mm, p_drops]
    if any(pd.notna(value) and value == -2 for value in candidates):
        return -2.0
    if any(pd.notna(value) and value < 0 for value in candidates):
        return -1.0
    return np.nan


# ===== CORE GRID-BUILDING FUNCTIONS =====

def make_global_index(station_data: OrderedDict) -> pd.DatetimeIndex | None:
    """
    Create the shared right-labelled 3-minute grid covering all station intervals.

    Args:
        station_data: OrderedDict of station name → DataFrame

    Returns:
        pd.DatetimeIndex: Complete time index or None if no valid data
    """
    grid_step = pd.Timedelta(minutes=GRID_STEP_MINUTES)
    interval_start_min = None
    interval_end_max = None

    for df in station_data.values():
        if df is None or df.empty:
            continue
        row_steps = infer_source_steps_minutes(df)
        interval_starts = df["DateTime"] - pd.to_timedelta(row_steps, unit="min")
        interval_ends = df["DateTime"]

        current_start = interval_starts.min()
        current_end = interval_ends.max()
        if pd.notna(current_start) and (interval_start_min is None or current_start < interval_start_min):
            interval_start_min = current_start
        if pd.notna(current_end) and (interval_end_max is None or current_end > interval_end_max):
            interval_end_max = current_end

    if interval_start_min is None or interval_end_max is None:
        return None

    first_label = (interval_start_min + grid_step).ceil(f"{GRID_STEP_MINUTES}min")
    last_label = interval_end_max.ceil(f"{GRID_STEP_MINUTES}min")
    return pd.date_range(first_label, last_label, freq=f"{GRID_STEP_MINUTES}min")


def distribute_station_to_global_grid(
    df: pd.DataFrame,
    global_index: pd.DatetimeIndex,
) -> tuple[pd.Series, pd.Series, pd.Series]:
    """
    Distribute station cumulative values to the shared 3-minute grid.

    Each station value is treated as an accumulation ending at DateTime.
    Its duration is inferred per Source_File (e.g., 5-minute data split across overlapping
    3-minute buckets, like MeteoSwiss 10-minute totals).

    Returns:
        tuple: (availability, precip, gap_codes) as Series aligned to global_index
    """
    precip_values = np.full(len(global_index), np.nan, dtype=np.float32)
    gap_values = np.full(len(global_index), np.nan, dtype=np.float32)
    availability_values = np.zeros(len(global_index), dtype=np.uint8)

    if df is None or df.empty or len(global_index) == 0:
        return (
            pd.Series(availability_values, index=global_index, dtype=np.uint8),
            pd.Series(precip_values, index=global_index, dtype=np.float32),
            pd.Series(gap_values, index=global_index, dtype=np.float32),
        )

    grid_step = pd.Timedelta(minutes=GRID_STEP_MINUTES)
    grid_start = global_index[0] - grid_step
    grid_step_ns = grid_step.value
    grid_start_ns = grid_start.value
    row_steps = infer_source_steps_minutes(df)

    precip_accumulator = np.zeros(len(global_index), dtype=float)
    has_precip_data = np.zeros(len(global_index), dtype=bool)

    for row, source_step_minutes in zip(df.itertuples(index=False), row_steps):
        interval_end = row.DateTime
        if pd.isna(interval_end):
            continue
        source_step = pd.Timedelta(minutes=float(source_step_minutes))
        source_step_ns = source_step.value
        interval_start = interval_end - source_step

        first_bin = int(np.floor((interval_start.value - grid_start_ns) / grid_step_ns))
        last_bin = int(np.floor((interval_end.value - 1 - grid_start_ns) / grid_step_ns))
        first_bin = max(first_bin, 0)
        last_bin = min(last_bin, len(global_index) - 1)
        if last_bin < first_bin:
            continue

        p_mm = row.P_mm
        p_drops = row.P_drops
        gap_code = classify_gap_code(p_mm, p_drops)
        is_measurement = pd.notna(p_mm) and p_mm >= 0

        for bin_idx in range(first_bin, last_bin + 1):
            bin_end = global_index[bin_idx]
            bin_start = bin_end - grid_step
            overlap_start = max(interval_start, bin_start)
            overlap_end = min(interval_end, bin_end)
            overlap_ns = max(0, overlap_end.value - overlap_start.value)
            if overlap_ns <= 0:
                continue

            if is_measurement:
                availability_values[bin_idx] = 1
                precip_accumulator[bin_idx] += float(p_mm) * (overlap_ns / source_step_ns)
                has_precip_data[bin_idx] = True
            elif np.isfinite(gap_code):
                if np.isnan(gap_values[bin_idx]) or gap_code < gap_values[bin_idx]:
                    gap_values[bin_idx] = gap_code

    precip_values[has_precip_data] = precip_accumulator[has_precip_data].astype(np.float32)
    return (
        pd.Series(availability_values, index=global_index, dtype=np.uint8),
        pd.Series(precip_values, index=global_index, dtype=np.float32),
        pd.Series(gap_values, index=global_index, dtype=np.float32),
    )


def build_station_grids(
    station_data: OrderedDict,
) -> tuple[pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None, pd.DatetimeIndex | None]:
    """
    Build availability, precipitation and gap-code grids on one common index.

    Returns:
        tuple: (availability_df, precip_df, gap_code_df, global_index) or (None, None, None, None)
    """
    global_index = make_global_index(station_data)
    if global_index is None:
        return None, None, None, None

    station_names = []
    availability_rows = []
    precip_rows = []
    gap_rows = []

    for station_name, df in station_data.items():
        if df is None or df.empty:
            print(f"Station {station_name}: No valid data, skipping.")
            continue
        availability, precip, gap = distribute_station_to_global_grid(df, global_index)
        if availability.sum() == 0 and precip.notna().sum() == 0 and gap.notna().sum() == 0:
            print(f"Station {station_name}: No usable measurements or gaps after gridding, skipping.")
            continue
        station_names.append(station_name)
        availability_rows.append(availability.to_numpy(dtype=np.uint8))
        precip_rows.append(precip.to_numpy(dtype=np.float32))
        gap_rows.append(gap.to_numpy(dtype=np.float32))

    if not station_names:
        return None, None, None, None

    availability_df = pd.DataFrame(
        np.vstack(availability_rows),
        index=station_names,
        columns=global_index,
    )
    precip_df = pd.DataFrame(
        np.vstack(precip_rows),
        index=station_names,
        columns=global_index,
    )
    gap_code_df = pd.DataFrame(
        np.vstack(gap_rows),
        index=station_names,
        columns=global_index,
    )
    return availability_df, precip_df, gap_code_df, global_index


# ===== OUTPUT/EXPORT FUNCTIONS (moved from Archive) =====

def save_availability_outputs(availability_df: pd.DataFrame) -> tuple[str | None, str | None]:
    """
    Export availability grid to parquet and summary Excel.

    Returns:
        tuple: (parquet_path, excel_path)
    """
    if availability_df is None or availability_df.empty:
        print('[ERROR] No availability grid to save.')
        return None, None

    parquet_path = os.path.join(OUTPUT_DIRECTORY_AVAILABILITY, 'station_availability_3min.parquet')
    try:
        availability_df.to_parquet(parquet_path)
        print(f'[INFO] Grid exported: {parquet_path}')
    except Exception as exc:
        csv_path = os.path.join(OUTPUT_DIRECTORY_AVAILABILITY, 'station_availability_3min.csv')
        availability_df.to_csv(csv_path)
        parquet_path = csv_path
        print(f'[Warning] Parquet failed; using CSV: {csv_path}')

    # Excel summary: per-station statistics
    excel_path = os.path.join(OUTPUT_DIRECTORY_AVAILABILITY, 'station_availability_summary.xlsx')
    total_steps = availability_df.shape[1]
    summary_rows = []
    for station_name, row in availability_df.iterrows():
        available_steps = int(row.sum())
        coverage_pct = 100.0 * available_steps / total_steps if total_steps > 0 else 0.0
        summary_rows.append({
            'Station': station_name,
            'Available timesteps': available_steps,
            'Total timesteps': total_steps,
            'Coverage percent': coverage_pct,
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_excel(excel_path, index=False)
    print(f'[INFO] Summary Excel exported: {excel_path}')

    return parquet_path, excel_path


def save_rain_outputs(
    precip_df: pd.DataFrame,
    gap_code_df: pd.DataFrame,
) -> tuple[str | None, str | None]:
    """
    Export precipitation grid to parquet and summary Excel.

    Returns:
        tuple: (parquet_path, excel_path)
    """
    if precip_df is None or precip_df.empty:
        print('[ERROR] No precipitation grid to save.')
        return None, None

    # Wide parquet
    parquet_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_3min.parquet')
    try:
        precip_df.to_parquet(parquet_path)
        print(f'[INFO] Precip grid exported: {parquet_path}')
    except Exception as exc:
        csv_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_3min.csv')
        precip_df.to_csv(csv_path)
        parquet_path = csv_path
        print(f'[Warning] Parquet failed; using CSV: {csv_path}')

    # Excel summary
    excel_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_summary.xlsx')
    total_steps = precip_df.shape[1]
    summary_rows = []
    for station_name, row in precip_df.iterrows():
        rain_steps = int((row > 0).sum())
        total_rain_mm = float(row.sum()) if row.notna().sum() > 0 else 0.0
        summary_rows.append({
            'Station': station_name,
            'Rainy timesteps': rain_steps,
            'Total precipitation mm': round(total_rain_mm, 2),
            'Mean precip per rainy step mm': round(total_rain_mm / rain_steps, 4) if rain_steps > 0 else 0.0,
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_excel(excel_path, index=False)
    print(f'[INFO] Summary Excel exported: {excel_path}')

    return parquet_path, excel_path


def build_daily_availability_heatmap(availability_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate 3-minute availability to daily (24-hour) summaries.

    Returns:
        pd.DataFrame: (stations × days) with daily coverage percentage
    """
    if availability_df is None or availability_df.empty:
        return pd.DataFrame()

    df_t = availability_df.transpose()
    df_t.index = pd.to_datetime(df_t.index)
    return (
        df_t
        .resample('1D', label='right', closed='right')
        .mean()
        .transpose()
        * 100
    )


def build_daily_rain_heatmap(precip_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate 3-minute precipitation to daily (24-hour) totals.

    Returns:
        pd.DataFrame: (stations × days) with daily total precipitation in mm
    """
    if precip_df is None or precip_df.empty:
        return pd.DataFrame()

    daily_rows = []
    daily_columns = None
    for station_name, row in precip_df.iterrows():
        series = pd.Series(row.values, index=precip_df.columns, dtype=np.float32)
        daily = series.resample('1D', label='right', closed='right').sum(min_count=1)
        daily_rows.append(daily.values)
        if daily_columns is None:
            daily_columns = daily.index

    daily_df = pd.DataFrame(
        np.vstack(daily_rows),
        index=precip_df.index,
        columns=daily_columns,
    )
    return daily_df


def build_30min_heatmap(precip_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate 3-minute precipitation to 30-minute totals.

    Returns:
        pd.DataFrame: (stations × 30-min periods) with 30-min total precipitation
    """
    if precip_df is None or precip_df.empty:
        return pd.DataFrame()

    precip_30min_rows = []
    period_columns = None
    for station_name, row in precip_df.iterrows():
        series = pd.Series(row.values, index=precip_df.columns, dtype=np.float32)
        precip_30min = series.resample('30min', label='right', closed='right').sum(min_count=1)
        precip_30min_rows.append(precip_30min.values)
        if period_columns is None:
            period_columns = precip_30min.index

    precip_30min_df = pd.DataFrame(
        np.vstack(precip_30min_rows),
        index=precip_df.index,
        columns=period_columns,
    )
    return precip_30min_df


def save_heatmap_html(
    df: pd.DataFrame,
    is_daily: bool = False,
    resolution_label: str = "3min",
    output_subdir: str = "availability",
) -> str | None:
    """
    Generate and save an interactive Plotly heatmap.

    Args:
        df: Data to plot (stations × time)
        is_daily: If True, normalize for display
        resolution_label: Label for output filename
        output_subdir: "availability" or "rain"

    Returns:
        Path to output HTML file or None
    """
    if df is None or df.empty:
        return None

    output_dir = OUTPUT_DIRECTORY_AVAILABILITY if output_subdir == "availability" else OUTPUT_DIRECTORY_RAIN
    if output_subdir == "availability":
        filename = "daily_availability_heatmap.html"
    else:
        filename = f"precipitation_heatmap_{resolution_label}.html"
    filepath = os.path.join(output_dir, filename)

    if output_subdir == "availability":
        fig = go.Figure(
            data=go.Heatmap(
                z=df.values,
                x=pd.to_datetime(df.columns),
                y=df.index,
                colorscale="Blues",
                zmin=0,
                zmax=100,
                hovertemplate="Station=%{y}<br>Date=%{x|%Y-%m-%d}<br>Coverage=%{z:.1f}%<extra></extra>",
                colorbar=dict(title="Coverage [%]"),
            )
        )
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="Station",
            width=1600,
            height=max(500, len(df.index) * 42 + 180),
            plot_bgcolor="white",
        )
        fig.write_html(filepath)
        print(f"[INFO] Heatmap saved: {filepath}")
        return filepath

    df_plot = df.copy()
    real_values = df_plot.to_numpy(dtype=float)
    real_values = real_values[np.isfinite(real_values) & (real_values >= 0)]
    if real_values.size == 0:
        real_max = 1.0
        z_cap = 1.0
    else:
        real_max = float(real_values.max())
        positive_values = real_values[real_values > 0]
        z_cap = 1.0 if positive_values.size == 0 else max(float(np.nanpercentile(positive_values, 99)), 1.0)

    zmax = float(np.sqrt(z_cap))
    gap_sentinel = -1.0
    gap_fraction = abs(gap_sentinel) / (zmax - gap_sentinel)
    z_values = df_plot.clip(lower=0, upper=z_cap).apply(np.sqrt).fillna(gap_sentinel)
    custom_colorscale = [[0.0, "#BDBDBD"], [gap_fraction, "#BDBDBD"]]
    custom_colorscale.extend(
        [[gap_fraction + (1 - gap_fraction) * position, color] for position, color in RAIN_COLORSCALE]
    )
    tick_original = np.array([0, z_cap * 0.25, z_cap * 0.5, z_cap * 0.75, z_cap])
    tick_values = np.sqrt(tick_original)
    tick_labels = [f"{value:.2f}" for value in tick_original]
    if real_max > z_cap:
        tick_labels[-1] = f">= {z_cap:.2f}"

    colorbar_title = "Precipitation<br>[mm/day]" if is_daily else f"Precipitation<br>[mm / {resolution_label}]"
    fig = go.Figure(
        data=go.Heatmap(
            z=z_values.values,
            x=pd.to_datetime(df_plot.columns),
            y=df_plot.index,
            colorscale=custom_colorscale,
            zmin=gap_sentinel,
            zmax=zmax,
            zsmooth=False,
            ygap=1,
            hovertemplate="Station=%{y}<br>Date=%{x}<br>Value=%{customdata}<extra></extra>",
            customdata=df_plot.round(4).astype(object).where(df_plot.notna(), "Data gap").values,
            colorbar=dict(
                title=dict(text=colorbar_title, side="right"),
                thickness=20,
                len=0.7,
                tickvals=tick_values,
                ticktext=tick_labels,
            ),
        )
    )
    if is_daily:
        width = 1800
        height = max(420, len(df_plot.index) * 35 + 160)
        x_nticks = 40
        x_title = "Date"
    else:
        n_timestamps = len(df_plot.columns)
        x_nticks = max(10, min(80, n_timestamps // 100))
        width = max(1800, min(6000, n_timestamps * 3))
        height = max(420, len(df_plot.index) * 35 + 180)
        x_title = f"Date / Time ({resolution_label} resolution)"

    fig.update_layout(
        title=dict(
            text=f"Precipitation Heatmap ({resolution_label.upper()}) - grey = data gap, white-blue = 0 mm",
            font=dict(size=15),
        ),
        xaxis=dict(title=x_title, type="date", tickangle=45, nticks=x_nticks),
        yaxis=dict(title="Station", type="category"),
        width=width,
        height=height,
        hovermode="closest",
        plot_bgcolor="white",
    )
    fig.write_html(filepath)
    print(f"[INFO] Heatmap saved: {filepath}")
    return filepath


def save_heatmap_png(
    df: pd.DataFrame,
    output_subdir: str,
    filename: str,
    is_availability: bool = False,
) -> str | None:
    """Save a static PNG companion for the report-style heatmaps."""
    if df is None or df.empty:
        return None

    df = df.sort_index()
    output_dir = OUTPUT_DIRECTORY_AVAILABILITY if output_subdir == "availability" else OUTPUT_DIRECTORY_RAIN
    filepath = os.path.join(output_dir, filename)
    plt.switch_backend("Agg")
    fig_width = 16
    fig_height = max(5, len(df.index) * 0.42 + 1.8)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))

    x_values = pd.to_datetime(df.columns)
    extent = [0, len(x_values), 0, len(df.index)]

    if is_availability:
        image = ax.imshow(
            df.values,
            aspect="auto",
            cmap="Blues",
            vmin=0,
            vmax=100,
            origin="upper",
            extent=extent,
            interpolation="nearest",
        )
        colorbar = fig.colorbar(image, ax=ax, fraction=0.025, pad=0.025)
        colorbar.set_label("Coverage [%]")
    else:
        values = df.to_numpy(dtype=float)
        positive = values[np.isfinite(values) & (values > 0)]
        rain_cap = max(float(np.nanpercentile(positive, 99)), 1.0) if positive.size else 1.0
        scaled = np.sqrt(np.clip(values, 0, rain_cap) / rain_cap)
        masked = np.ma.masked_invalid(scaled)
        cmap = LinearSegmentedColormap.from_list("rain_grid", [color for _, color in RAIN_COLORSCALE])
        cmap.set_bad("#BDBDBD")
        image = ax.imshow(
            masked,
            aspect="auto",
            cmap=cmap,
            vmin=0,
            vmax=1,
            origin="upper",
            extent=extent,
            interpolation="nearest",
        )
        colorbar = fig.colorbar(image, ax=ax, fraction=0.025, pad=0.025)
        colorbar.set_label("Precipitation [mm/day]")

    ax.set_yticks(len(df.index) - np.arange(len(df.index)) - 0.5)
    ax.set_yticklabels(df.index)
    tick_count = min(10, len(x_values))
    tick_positions = np.linspace(0, len(x_values), tick_count, endpoint=False)
    tick_indices = np.clip(tick_positions.astype(int), 0, len(x_values) - 1)
    ax.set_xticks(tick_positions + 0.5)
    ax.set_xticklabels([x_values[i].strftime("%b %Y") for i in tick_indices], rotation=45, ha="right")
    ax.set_xlabel("Date")
    ax.set_ylabel("Station")
    fig.tight_layout()
    fig.savefig(filepath, dpi=180, bbox_inches="tight")
    plt.close(fig)
    print(f"[INFO] Static heatmap saved: {filepath}")
    return filepath


def save_period_summaries(
    grid_df: pd.DataFrame,
    grid_type: str = "availability",
) -> None:
    """
    Save daily and weekly aggregation summaries.

    Args:
        grid_df: Input grid (stations × time)
        grid_type: "availability" or "precip"
    """
    if grid_df is None or grid_df.empty:
        return

    if grid_type == "availability":
        output_dir = OUTPUT_DIRECTORY_AVAILABILITY
    else:
        output_dir = OUTPUT_DIRECTORY_RAIN

    excel_path = os.path.join(output_dir, f"station_{grid_type}_period_summary.xlsx")

    # Aggregate per station over entire period
    period_summaries = []
    for station_name, row in grid_df.iterrows():
        total = row.notna().sum()
        nonzero = (row > 0).sum()
        mean_val = row.mean()
        period_summaries.append({
            'Station': station_name,
            'Total timesteps': total,
            'Nonzero timesteps': nonzero,
            'Mean value': round(mean_val, 4),
        })

    summary_df = pd.DataFrame(period_summaries)
    summary_df.to_excel(excel_path, index=False)
    print(f"[INFO] Period summary exported: {excel_path}")


def add_meteosuisse_to_grid(
    precip_df: pd.DataFrame,
    gap_code_df: pd.DataFrame,
    global_index: pd.DatetimeIndex,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Optionally add MeteoSwiss reference data to precipitation grid (stub for now).

    Returns:
        tuple: (precip_df, gap_code_df) potentially modified
    """
    # TODO: Implement MeteoSwiss integration if needed
    # For now, pass through unchanged
    return precip_df, gap_code_df


def save_qc_long_outputs(
    precip_df: pd.DataFrame,
    gap_code_df: pd.DataFrame,
) -> tuple[str | None, str | None]:
    """
    Export precipitation data in long (tidy) format for analysis.

    Returns:
        tuple: (parquet_path, csv_path)
    """
    if precip_df is None or precip_df.empty:
        return None, None

    records = []
    for station_name, row in precip_df.iterrows():
        for timestamp, value in zip(precip_df.columns, row.values):
            gap_code = gap_code_df.loc[station_name, timestamp] if station_name in gap_code_df.index else np.nan
            records.append({
                'Station': station_name,
                'DateTime': timestamp,
                'Precip_mm': value,
                'Gap_code': gap_code,
            })

    long_df = pd.DataFrame(records)

    long_parquet_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_long.parquet')
    try:
        long_df.to_parquet(long_parquet_path)
        print(f"[INFO] Long parquet exported: {long_parquet_path}")
    except Exception:
        long_parquet_path = None

    long_csv_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_long.csv')
    long_df.to_csv(long_csv_path, index=False)
    print(f"[INFO] Long CSV exported: {long_csv_path}")

    return long_parquet_path, long_csv_path


def save_monthly_qc_excel(
    precip_df: pd.DataFrame,
    gap_code_df: pd.DataFrame,
) -> str | None:
    """
    Generate monthly QC report with precipitation and gap statistics.

    Returns:
        Path to output Excel file or None
    """
    if precip_df is None or precip_df.empty:
        return None

    excel_path = os.path.join(OUTPUT_DIRECTORY_RAIN, 'station_precip_monthly_qc.xlsx')

    months = precip_df.columns.normalize().unique()
    monthly_stats = []
    for station_name, row in precip_df.iterrows():
        for month in months:
            month_mask = (precip_df.columns >= month) & (precip_df.columns < month + pd.DateOffset(months=1))
            month_values = row[month_mask]
            rain_steps = (month_values > 0).sum()
            total_mm = month_values.sum()
            monthly_stats.append({
                'Station': station_name,
                'Month': month.strftime('%Y-%m'),
                'Rainy timesteps': rain_steps,
                'Total precipitation mm': round(total_mm, 2),
            })

    stats_df = pd.DataFrame(monthly_stats)
    stats_df.to_excel(excel_path, index=False)
    print(f"[INFO] Monthly QC Excel exported: {excel_path}")

    return excel_path
