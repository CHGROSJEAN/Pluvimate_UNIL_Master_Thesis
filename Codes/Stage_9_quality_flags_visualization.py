# Charlotte Grosjean - UNIL Master Thesis
"""
Visualize final quality flags after processed-file quality control and station
consolidation.

Run this script after:
1. Stage_8_quality_flags.py
2. Stage_4_consolidate_stations_files.py

Outputs:
- daily quality-flag grid across stations;
- final quality-flag distribution by station;
- monthly valid/rejected quality timeline;
- daily PASS/FAIL/SKIP summary from Stage 8 daily reports;
- daily correlation-vs-bias diagnostic scatter.
"""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd

os.environ.setdefault("MPLCONFIGDIR", os.path.join(tempfile.gettempdir(), "matplotlib"))
os.environ.setdefault("XDG_CACHE_HOME", tempfile.gettempdir())

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import colors as mcolors
from matplotlib.colors import ListedColormap, BoundaryNorm
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

from F_rain_quality_plot_utils import (
    RAIN_COLORSCALE,
    compute_rain_cap,
    rain_color_from_value,
    sqrt_scale_rain,
)
from I_paths_config import (
    CONSOLIDATED_DIRECTORY,
    DAILY_REPORTS_DIRECTORY,
    QUALITY_FIGURES_DIRECTORY,
)


# %% 0 - Configuration

# ===== FILES =====
OUTPUT_DIRECTORY = QUALITY_FIGURES_DIRECTORY
DAILY_REPORT_DIRECTORY = DAILY_REPORTS_DIRECTORY

# Create output directory if not exist
OUTPUT_DIRECTORY.mkdir(parents=True, exist_ok=True)

# ===== PLOT CONFIGURATION =====
FIGURE_DPI = 180
A4_LANDSCAPE_FIGSIZE = (11.69, 8.27)
PLOT_TITLE_FONTSIZE = 14
PLOT_LABEL_FONTSIZE = 12
PLOT_TICK_FONTSIZE = 10.5
PLOT_LEGEND_FONTSIZE = 10.5
SPLIT_GRID_TITLE_FONTSIZE = 14
SPLIT_GRID_LABEL_FONTSIZE = 12
SPLIT_GRID_TICK_FONTSIZE = 10
SPLIT_GRID_LEGEND_FONTSIZE = 10
SPLIT_GRID_LEGEND_TITLE_FONTSIZE = 10.5
SPLIT_GRID_COLORBAR_LABEL_FONTSIZE = 10.5
SPLIT_GRID_COLORBAR_TICK_FONTSIZE = 9.5

plt.rcParams.update(
    {
        "axes.titlesize": PLOT_TITLE_FONTSIZE,
        "axes.labelsize": PLOT_LABEL_FONTSIZE,
        "xtick.labelsize": PLOT_TICK_FONTSIZE,
        "ytick.labelsize": PLOT_TICK_FONTSIZE,
        "legend.fontsize": PLOT_LEGEND_FONTSIZE,
    }
)
STATION_ORDER = [
    "BETH",
    "BOIS",
    "CHAN",
    "ELYS",
    "GEOP",
    "GRVN",
    "LEXP",
    "PONT",
    "RIPR",
    "ROUV",
    "VCLB",
]

FLAG_LABELS = {
    -2: "Intra-file gap",
    -1: "Inter-file gap",
    0: "Valid",
    1: "Uncertain",
    2: "Rejected file",
    3: "Rejected day",
    4: "Outlier",
}

FLAG_COLORS = {
    -2: "#d9d9d9",
    -1: "#f2f2f2",
    0: "#2ca25f",
    1: "#fdae61",
    2: "#b2182b",
    3: "#ef6548",
    4: "#54278f",
}

DAILY_STATUS_COLORS = {
    "PASS": "#2ca25f",
    "FAIL": "#ef6548",
    "SKIP": "#bdbdbd",
}


# %% 1 - Helper functions

def station_sort_key(station: str) -> tuple[int, str]:
    if station in STATION_ORDER:
        return STATION_ORDER.index(station), station
    return len(STATION_ORDER), station


def list_consolidated_files() -> list[Path]:
    return sorted(
        CONSOLIDATED_DIRECTORY.glob("*_Consolidated_Complete.csv"),
        key=lambda path: station_sort_key(path.name.split("_")[0]),
    )


def load_consolidated_quality_flags() -> pd.DataFrame:
    """Load final quality flags and precipitation from consolidated files."""
    records: list[pd.DataFrame] = []

    for file_path in list_consolidated_files():
        station = file_path.name.split("_")[0]
        df = pd.read_csv(
            file_path,
            usecols=lambda col: col in {"DateTime", "P_mm", "quality_flag"},
        )
        if not {"DateTime", "P_mm", "quality_flag"}.issubset(df.columns):
            print(f"Skipping {file_path.name}: missing DateTime, P_mm or quality_flag")
            continue

        df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
        df["P_mm"] = pd.to_numeric(df["P_mm"], errors="coerce")
        df["quality_flag"] = pd.to_numeric(df["quality_flag"], errors="coerce")
        df = df.dropna(subset=["DateTime", "quality_flag"])
        df["quality_flag"] = df["quality_flag"].astype(int)
        df["Station"] = station
        df["Date"] = df["DateTime"].dt.floor("D")
        df["Month"] = df["DateTime"].dt.to_period("M").dt.to_timestamp()
        records.append(df[["Station", "DateTime", "Date", "Month", "P_mm", "quality_flag"]])

    if not records:
        raise FileNotFoundError(
            f"No consolidated quality-flag files found in {CONSOLIDATED_DIRECTORY}"
        )

    output = pd.concat(records, ignore_index=True)
    output["Station"] = pd.Categorical(
        output["Station"],
        categories=sorted(output["Station"].unique(), key=station_sort_key),
        ordered=True,
    )
    return output


def daily_severity_flag(flags: pd.Series) -> int:
    """
    Return one daily flag per station using a conservative severity order.

    Outliers and rejected file/day flags dominate, then uncertain, valid, and gaps.
    """
    values = set(pd.to_numeric(flags, errors="coerce").dropna().astype(int))
    for flag in [4, 2, 3, 1, 0, -2, -1]:
        if flag in values:
            return flag
    return np.nan


def prepare_daily_flag_and_rain_grids(flags_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, float]:
    """Build aligned daily quality-flag and precipitation grids."""
    daily_flags = (
        flags_df.groupby(["Station", "Date"], observed=False)["quality_flag"]
        .apply(daily_severity_flag)
        .reset_index()
    )
    flag_grid = daily_flags.pivot(index="Station", columns="Date", values="quality_flag")
    flag_grid = flag_grid.sort_index(key=lambda index: [station_sort_key(str(value)) for value in index])
    flag_grid = flag_grid.sort_index(axis=1)

    measurements = flags_df[flags_df["P_mm"].ge(0)].copy()
    daily_rain = (
        measurements.groupby(["Station", "Date"], observed=False)["P_mm"]
        .sum(min_count=1)
        .reset_index()
    )
    rain_grid = daily_rain.pivot(index="Station", columns="Date", values="P_mm")
    rain_grid = rain_grid.reindex(index=flag_grid.index, columns=flag_grid.columns)

    rain_cap = compute_rain_cap(rain_grid.to_numpy(dtype=float))

    return flag_grid, rain_grid, rain_cap


# %% 2 - Main output generation

def save_daily_quality_grid(flags_df: pd.DataFrame) -> Path:
    """Heatmap grid: stations by day, colored by final daily quality flag."""
    daily = (
        flags_df.groupby(["Station", "Date"], observed=False)["quality_flag"]
        .apply(daily_severity_flag)
        .reset_index()
    )
    grid = daily.pivot(index="Station", columns="Date", values="quality_flag")
    grid = grid.sort_index(key=lambda index: [station_sort_key(str(value)) for value in index])
    grid = grid.sort_index(axis=1)

    flag_values = [-2, -1, 0, 1, 2, 3, 4]
    colors = [FLAG_COLORS[value] for value in flag_values]
    cmap = ListedColormap(colors)
    boundaries = [-2.5, -1.5, -0.5, 0.5, 1.5, 2.5, 3.5, 4.5]
    norm = BoundaryNorm(boundaries, cmap.N)

    figure_width = max(12, min(26, grid.shape[1] / 35))
    figure_height = max(5.5, 0.42 * len(grid.index) + 2.4)
    fig, ax = plt.subplots(figsize=(figure_width, figure_height))
    ax.imshow(grid.to_numpy(dtype=float), aspect="auto", cmap=cmap, norm=norm)

    ax.set_yticks(np.arange(len(grid.index)))
    ax.set_yticklabels(grid.index.astype(str))
    ax.set_ylabel("Station")
    ax.set_title("Daily final quality flag grid")

    date_columns = pd.to_datetime(grid.columns)
    tick_count = min(10, len(date_columns))
    if tick_count > 0:
        tick_positions = np.linspace(0, len(date_columns) - 1, tick_count, dtype=int)
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(
            [date_columns[pos].strftime("%Y-%m") for pos in tick_positions],
            rotation=45,
            ha="right",
        )
    ax.set_xlabel("Date")

    legend_handles = [
        Patch(facecolor=FLAG_COLORS[value], edgecolor="none", label=f"{value}: {FLAG_LABELS[value]}")
        for value in flag_values
    ]
    ax.legend(
        handles=legend_handles,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=3,
        frameon=False,
    )
    fig.tight_layout()

    output_path = OUTPUT_DIRECTORY / "quality_flag_grid_daily.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)
    return output_path


def save_daily_rain_quality_split_grid(flags_df: pd.DataFrame) -> Path:
    """
    Combined daily grid with two thin bands per station:
    upper band = final quality flag, lower band = daily precipitation total.
    """
    flag_grid, rain_grid, rain_cap = prepare_daily_flag_and_rain_grids(flags_df)
    n_stations = len(flag_grid.index)
    n_dates = len(flag_grid.columns)
    image = np.ones((n_stations * 2, n_dates, 4), dtype=float)

    for station_idx, station in enumerate(flag_grid.index):
        flag_row = flag_grid.loc[station].to_numpy(dtype=float)
        rain_row = rain_grid.loc[station].to_numpy(dtype=float)
        for col_idx, flag_value in enumerate(flag_row):
            if np.isfinite(flag_value):
                color = FLAG_COLORS.get(int(flag_value), "#ffffff")
                image[station_idx * 2, col_idx, :] = mcolors.to_rgba(color)
            else:
                image[station_idx * 2, col_idx, :] = mcolors.to_rgba("#ffffff")

        for col_idx, rain_value in enumerate(rain_row):
            image[station_idx * 2 + 1, col_idx, :] = rain_color_from_value(rain_value, rain_cap)

    fig, ax = plt.subplots(figsize=A4_LANDSCAPE_FIGSIZE)
    ax.imshow(image, aspect="auto", interpolation="nearest")
    fig.subplots_adjust(left=0.075, right=0.995, top=0.99, bottom=0.29)

    ax.set_yticks(np.arange(n_stations) * 2 + 0.5)
    ax.set_yticklabels(flag_grid.index.astype(str))
    ax.set_ylabel("Station", fontsize=SPLIT_GRID_LABEL_FONTSIZE)

    date_columns = pd.to_datetime(flag_grid.columns)
    tick_count = min(10, len(date_columns))
    if tick_count > 0:
        tick_positions = np.linspace(0, len(date_columns) - 1, tick_count, dtype=int)
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(
            [date_columns[pos].strftime("%Y-%m") for pos in tick_positions],
            rotation=35,
            ha="right",
        )
    ax.set_xlabel("Date", fontsize=SPLIT_GRID_LABEL_FONTSIZE, labelpad=2)
    ax.tick_params(axis="both", labelsize=SPLIT_GRID_TICK_FONTSIZE)

    for boundary in np.arange(1.5, n_stations * 2, 2):
        ax.axhline(boundary, color="white", linewidth=1.8)

    flag_handles = [
        Patch(facecolor=FLAG_COLORS[value], edgecolor="none", label=f"{value}: {FLAG_LABELS[value]}")
        for value in [-2, -1, 0, 1, 2, 3, 4]
    ]
    fig.text(
        0.5,
        0.158,
        "Upper band: quality flag",
        ha="center",
        va="center",
        fontsize=SPLIT_GRID_LEGEND_TITLE_FONTSIZE,
    )
    flag_legend = fig.legend(
        handles=flag_handles,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.143),
        ncol=7,
        columnspacing=1.05,
        handlelength=1.25,
        frameon=False,
        fontsize=SPLIT_GRID_LEGEND_FONTSIZE,
    )

    sm = matplotlib.cm.ScalarMappable(
        cmap=mcolors.LinearSegmentedColormap.from_list(
            "stage9_rain",
            [color for _, color in RAIN_COLORSCALE],
        ),
        norm=plt.Normalize(vmin=0, vmax=rain_cap),
    )
    sm.set_array([])
    cbar_ax = fig.add_axes([0.30, 0.055, 0.45, 0.026])
    cbar = fig.colorbar(sm, cax=cbar_ax, orientation="horizontal")
    cbar.ax.set_anchor("C")
    cbar.set_label(
        f"Lower band: daily precipitation [mm/day], capped at {rain_cap:.1f}",
        fontsize=SPLIT_GRID_COLORBAR_LABEL_FONTSIZE,
        labelpad=7,
    )
    cbar.ax.tick_params(labelsize=SPLIT_GRID_COLORBAR_TICK_FONTSIZE)

    output_path = OUTPUT_DIRECTORY / "daily_rain_quality_split_grid.png"
    fig.savefig(output_path, dpi=300)
    fig.savefig(output_path.with_suffix(".pdf"))
    plt.close(fig)
    return output_path


def save_daily_rain_quality_split_grid_interactive(flags_df: pd.DataFrame) -> Path | None:
    """Interactive HTML version of the split rain/quality daily grid."""
    try:
        import plotly.graph_objects as go
    except ImportError:
        print("Plotly not installed; skipping interactive split grid.")
        return None

    flag_grid, rain_grid, rain_cap = prepare_daily_flag_and_rain_grids(flags_df)
    stations = [str(station) for station in flag_grid.index]
    dates = pd.to_datetime(flag_grid.columns)
    x_labels = dates

    rain_z_rows = []
    flag_z_rows = []
    y_labels = []
    rain_hover_rows = []
    flag_hover_rows = []
    flag_values = [-2, -1, 0, 1, 2, 3, 4]
    flag_to_z = {flag: idx for idx, flag in enumerate(flag_values)}

    flag_colorscale = []
    for flag, z_value in flag_to_z.items():
        position = z_value / (len(flag_values) - 1)
        color = FLAG_COLORS[flag]
        flag_colorscale.append([max(0.0, position - 0.0001), color])
        flag_colorscale.append([min(1.0, position + 0.0001), color])

    rain_zmax = float(np.sqrt(rain_cap))
    gap_sentinel = -1.0
    gap_fraction = abs(gap_sentinel) / (rain_zmax - gap_sentinel)
    rain_colorscale = [
        [0.0, "#BDBDBD"],
        [gap_fraction, "#BDBDBD"],
    ]
    rain_colorscale.extend(
        [
            [gap_fraction + (1 - gap_fraction) * position, color]
            for position, color in RAIN_COLORSCALE
        ]
    )

    for station in stations:
        flag_values_row = flag_grid.loc[station].to_numpy(dtype=float)
        rain_values_row = rain_grid.loc[station].to_numpy(dtype=float)

        z_flag = [
            flag_to_z.get(int(value), np.nan) if np.isfinite(value) else np.nan
            for value in flag_values_row
        ]
        flag_hover = []
        for flag_value in flag_values_row:
            if np.isfinite(flag_value):
                flag_int = int(flag_value)
                flag_hover.append(f"Quality flag {flag_int}: {FLAG_LABELS.get(flag_int, '')}")
            else:
                flag_hover.append("No quality flag")

        z_rain = []
        rain_hover = []
        for rain_value in rain_values_row:
            if not np.isfinite(rain_value):
                z_rain.append(gap_sentinel)
                rain_hover.append("No precipitation data")
            elif rain_value <= 0:
                z_rain.append(0.0)
                rain_hover.append("0.00 mm/day")
            else:
                z_rain.append(float(np.sqrt(rain_cap) * sqrt_scale_rain(rain_value, rain_cap)))
                rain_hover.append(f"{rain_value:.2f} mm/day")

        flag_z_rows.extend([z_flag, [np.nan] * len(x_labels)])
        rain_z_rows.extend([[np.nan] * len(x_labels), z_rain])
        y_labels.extend([f"{station} quality", f"{station} rain"])
        flag_hover_rows.extend([flag_hover, [""] * len(x_labels)])
        rain_hover_rows.extend([[""] * len(x_labels), rain_hover])

    fig = go.Figure()
    fig.add_trace(
        go.Heatmap(
            z=rain_z_rows,
            x=x_labels,
            y=y_labels,
            colorscale=rain_colorscale,
            zmin=gap_sentinel,
            zmax=rain_zmax,
            zsmooth=False,
            hovertext=rain_hover_rows,
            hovertemplate="%{y}<br>%{x|%Y-%m-%d}<br>%{hovertext}<extra></extra>",
            hoverongaps=False,
            colorbar=dict(
                title=dict(text="Daily precipitation<br>[mm/day]"),
                tickvals=[0, np.sqrt(rain_cap * 0.25), np.sqrt(rain_cap * 0.5), np.sqrt(rain_cap * 0.75), np.sqrt(rain_cap)],
                ticktext=["0", f"{rain_cap * 0.25:.1f}", f"{rain_cap * 0.5:.1f}", f"{rain_cap * 0.75:.1f}", f">= {rain_cap:.1f}"],
            ),
        )
    )
    fig.add_trace(
        go.Heatmap(
            z=flag_z_rows,
            x=x_labels,
            y=y_labels,
            colorscale=flag_colorscale,
            zmin=0,
            zmax=len(flag_values) - 1,
            zsmooth=False,
            hovertext=flag_hover_rows,
            hovertemplate="%{y}<br>%{x|%Y-%m-%d}<br>%{hovertext}<extra></extra>",
            hoverongaps=False,
            showscale=False,
        )
    )
    fig.update_layout(
        title="Daily precipitation and final quality flags",
        xaxis=dict(title="Date", type="date", tickangle=45),
        yaxis=dict(title="Station / band", autorange="reversed"),
        height=max(650, len(y_labels) * 32 + 170),
        width=1800,
        hovermode="closest",
        plot_bgcolor="white",
    )

    output_path = OUTPUT_DIRECTORY / "daily_rain_quality_split_grid_interactive.html"
    fig.write_html(
        output_path,
        include_plotlyjs="cdn",
        config={
            "scrollZoom": True,
            "displaylogo": False,
            "modeBarButtonsToAdd": ["zoomIn2d", "zoomOut2d", "resetScale2d"],
        },
    )
    return output_path


def save_flag_distribution_by_station(flags_df: pd.DataFrame) -> Path:
    """Stacked barplot of final timestep quality flags per station."""
    counts = (
        flags_df.groupby(["Station", "quality_flag"], observed=False)
        .size()
        .unstack(fill_value=0)
        .sort_index(key=lambda index: [station_sort_key(str(value)) for value in index])
    )
    plot_flags = [-2, -1, 0, 1, 2, 3, 4]
    counts = counts.reindex(columns=plot_flags, fill_value=0)
    proportions = counts.div(counts.sum(axis=1), axis=0) * 100

    fig, ax = plt.subplots(figsize=(11, 6))
    bottom = np.zeros(len(proportions))
    x = np.arange(len(proportions.index))
    for flag in plot_flags:
        values = proportions[flag].to_numpy()
        ax.bar(
            x,
            values,
            bottom=bottom,
            color=FLAG_COLORS[flag],
            label=f"{flag}: {FLAG_LABELS[flag]}",
            width=0.78,
        )
        bottom += values

    ax.set_xticks(x)
    ax.set_xticklabels(proportions.index.astype(str), rotation=45, ha="right")
    ax.set_ylabel("Timesteps (%)")
    ax.set_ylim(0, 100)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3, frameon=False)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()

    output_path = OUTPUT_DIRECTORY / "quality_flag_distribution_by_station.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)

    proportions.round(2).to_csv(OUTPUT_DIRECTORY / "quality_flag_distribution_by_station.csv")
    return output_path


def save_measurement_flag_distribution_by_station(flags_df: pd.DataFrame) -> Path:
    """Stacked barplot of final quality flags excluding gap rows."""
    measurement_df = flags_df[flags_df["quality_flag"].isin([0, 1, 2, 3, 4])].copy()
    counts = (
        measurement_df.groupby(["Station", "quality_flag"], observed=False)
        .size()
        .unstack(fill_value=0)
        .sort_index(key=lambda index: [station_sort_key(str(value)) for value in index])
    )
    plot_flags = [0, 1, 2, 3, 4]
    counts = counts.reindex(columns=plot_flags, fill_value=0)
    proportions = counts.div(counts.sum(axis=1), axis=0) * 100

    fig, ax = plt.subplots(figsize=(10.5, 5.8))
    bottom = np.zeros(len(proportions))
    x = np.arange(len(proportions.index))
    for flag in plot_flags:
        values = proportions[flag].to_numpy()
        ax.bar(
            x,
            values,
            bottom=bottom,
            color=FLAG_COLORS[flag],
            label=f"{flag}: {FLAG_LABELS[flag]}",
            width=0.78,
        )
        bottom += values

    ax.set_xticks(x)
    ax.set_xticklabels(proportions.index.astype(str), rotation=45, ha="right")
    ax.set_ylabel("Measured timesteps (%)")
    ax.set_ylim(0, 100)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=4, frameon=False)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()

    output_path = OUTPUT_DIRECTORY / "quality_flag_distribution_by_station_no_gaps.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)

    proportions.round(2).to_csv(
        OUTPUT_DIRECTORY / "quality_flag_distribution_by_station_no_gaps.csv"
    )
    return output_path


def save_monthly_quality_timeline(flags_df: pd.DataFrame) -> Path:
    """Small-multiple monthly timeline of valid and rejected timestep shares."""
    monthly_counts = (
        flags_df.groupby(["Station", "Month", "quality_flag"], observed=False)
        .size()
        .rename("n")
        .reset_index()
    )
    total = monthly_counts.groupby(["Station", "Month"], observed=False)["n"].transform("sum")
    monthly_counts["pct"] = monthly_counts["n"] / total * 100

    valid = monthly_counts[monthly_counts["quality_flag"] == 0]
    rejected = monthly_counts[monthly_counts["quality_flag"].isin([2, 3, 4])]
    valid = valid.pivot(index="Month", columns="Station", values="pct")
    rejected = (
        rejected.groupby(["Month", "Station"], observed=False)["pct"]
        .sum()
        .reset_index()
        .pivot(index="Month", columns="Station", values="pct")
    )

    stations = sorted(flags_df["Station"].dropna().astype(str).unique(), key=station_sort_key)
    n_rows = len(stations)
    fig, axes = plt.subplots(
        n_rows,
        1,
        figsize=(12, max(6, 1.25 * n_rows)),
        sharex=True,
        constrained_layout=True,
    )
    if n_rows == 1:
        axes = [axes]

    for ax, station in zip(axes, stations):
        valid_series = valid.get(station, pd.Series(dtype=float)).sort_index()
        rejected_series = rejected.get(station, pd.Series(dtype=float)).sort_index()
        ax.plot(valid_series.index, valid_series.values, color=FLAG_COLORS[0], linewidth=1.5)
        ax.fill_between(
            rejected_series.index,
            rejected_series.values,
            color=FLAG_COLORS[3],
            alpha=0.45,
            linewidth=0,
        )
        ax.set_ylim(0, 100)
        ax.set_ylabel(station, rotation=0, ha="right", va="center")
        ax.grid(axis="y", alpha=0.2)

    axes[0].set_title("Monthly quality timeline: valid line and rejected area")
    axes[-1].xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    axes[-1].xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    axes[-1].tick_params(axis="x", rotation=45)
    fig.legend(
        handles=[
            Line2D([0], [0], color=FLAG_COLORS[0], lw=2, label="Flag 0 valid (%)"),
            Patch(facecolor=FLAG_COLORS[3], alpha=0.45, label="Flags 2+3+4 rejected (%)"),
        ],
        loc="lower center",
        ncol=2,
        frameon=False,
    )

    output_path = OUTPUT_DIRECTORY / "monthly_quality_timeline.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)
    return output_path


def parse_daily_report_station(file_path: Path) -> str:
    match = re.match(r"^(?P<station>[A-Za-z0-9]+)_", file_path.name)
    return match.group("station") if match else "Unknown"


def load_daily_reports() -> pd.DataFrame:
    """Load Stage 8 daily quality reports, if available."""
    records = []
    for file_path in sorted(DAILY_REPORT_DIRECTORY.glob("*_daily_quality.csv")):
        station = parse_daily_report_station(file_path)
        df = pd.read_csv(file_path)
        if "Daily status" not in df.columns:
            continue
        df["Station"] = station
        df["Source report"] = file_path.name
        records.append(df)

    if not records:
        return pd.DataFrame()

    output = pd.concat(records, ignore_index=True)
    output["Date"] = pd.to_datetime(output["Date"], errors="coerce")
    output["Station"] = pd.Categorical(
        output["Station"],
        categories=sorted(output["Station"].dropna().unique(), key=station_sort_key),
        ordered=True,
    )
    return output


def save_daily_pass_fail_skip_summary(daily_df: pd.DataFrame) -> Path | None:
    """Stacked barplot of daily PASS/FAIL/SKIP counts per station."""
    if daily_df.empty:
        print("No daily quality reports found; skipping PASS/FAIL/SKIP figure.")
        return None

    counts = (
        daily_df.groupby(["Station", "Daily status"], observed=False)
        .size()
        .unstack(fill_value=0)
        .sort_index(key=lambda index: [station_sort_key(str(value)) for value in index])
    )
    statuses = ["PASS", "FAIL", "SKIP"]
    counts = counts.reindex(columns=statuses, fill_value=0)

    fig, ax = plt.subplots(figsize=(10.5, 5.8))
    bottom = np.zeros(len(counts))
    x = np.arange(len(counts.index))
    for status in statuses:
        values = counts[status].to_numpy()
        ax.bar(
            x,
            values,
            bottom=bottom,
            color=DAILY_STATUS_COLORS[status],
            label=status,
            width=0.78,
        )
        bottom += values

    ax.set_xticks(x)
    ax.set_xticklabels(counts.index.astype(str), rotation=45, ha="right")
    ax.set_ylabel("Number of daily checks")
    ax.set_title("Daily quality-check outcomes by station")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3, frameon=False)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()

    output_path = OUTPUT_DIRECTORY / "daily_pass_fail_skip_by_station.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)

    counts.to_csv(OUTPUT_DIRECTORY / "daily_pass_fail_skip_by_station.csv")
    return output_path


def save_daily_correlation_bias_scatter(daily_df: pd.DataFrame) -> Path | None:
    """Scatter plot of daily correlation vs daily bias with decision thresholds."""
    if daily_df.empty:
        print("No daily quality reports found; skipping correlation/bias scatter.")
        return None

    df = daily_df.copy()
    df["Daily correlation"] = pd.to_numeric(df["Daily correlation"], errors="coerce")
    df["Daily bias"] = pd.to_numeric(df["Daily bias"], errors="coerce")
    df = df.dropna(subset=["Daily correlation", "Daily bias"])
    if df.empty:
        print("No finite daily correlation/bias values; skipping scatter.")
        return None

    fig, ax = plt.subplots(figsize=(8.5, 6.5))
    for status, group in df.groupby("Daily status"):
        ax.scatter(
            group["Daily correlation"],
            group["Daily bias"],
            s=28,
            alpha=0.68,
            color=DAILY_STATUS_COLORS.get(status, "#636363"),
            edgecolor="white",
            linewidth=0.3,
            label=status,
        )

    ax.axvline(0.40, color="#333333", linestyle="--", linewidth=1.0)
    ax.axhline(0.30, color="#333333", linestyle=":", linewidth=1.0)
    ax.axhline(-0.30, color="#333333", linestyle=":", linewidth=1.0)
    ax.set_xlabel("Daily Pearson correlation")
    ax.set_ylabel("Daily relative bias")
    ax.set_title("Daily quality-check diagnostics")
    ax.set_xlim(-1.02, 1.02)
    bias_limit = max(1.0, min(5.0, np.nanpercentile(np.abs(df["Daily bias"]), 98) * 1.15))
    ax.set_ylim(-bias_limit, bias_limit)
    ax.grid(alpha=0.25)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.14), ncol=3, frameon=False)
    fig.tight_layout()

    output_path = OUTPUT_DIRECTORY / "daily_correlation_bias_scatter.png"
    fig.savefig(output_path, dpi=FIGURE_DPI, bbox_inches="tight")
    plt.close(fig)
    return output_path


# %% 3 - Main orchestration

def main() -> None:
    print("Loading consolidated quality flags...", flush=True)
    flags_df = load_consolidated_quality_flags()
    print(f"Loaded {len(flags_df):,} consolidated timesteps.", flush=True)

    outputs: list[Path] = []
    print("Creating daily quality-flag grid...", flush=True)
    outputs.append(save_daily_quality_grid(flags_df))
    outputs.append(save_daily_rain_quality_split_grid(flags_df))
    interactive_split_grid = save_daily_rain_quality_split_grid_interactive(flags_df)
    if interactive_split_grid is not None:
        outputs.append(interactive_split_grid)

    print("Creating quality-flag distribution by station...", flush=True)
    outputs.append(save_flag_distribution_by_station(flags_df))
    outputs.append(save_measurement_flag_distribution_by_station(flags_df))

    print("Creating monthly quality timeline...", flush=True)
    outputs.append(save_monthly_quality_timeline(flags_df))

    print("Loading daily quality reports...", flush=True)
    daily_df = load_daily_reports()
    print(f"Loaded {len(daily_df):,} daily report rows.", flush=True)

    daily_summary_path = save_daily_pass_fail_skip_summary(daily_df)
    if daily_summary_path is not None:
        outputs.append(daily_summary_path)

    scatter_path = save_daily_correlation_bias_scatter(daily_df)
    if scatter_path is not None:
        outputs.append(scatter_path)

    print("\nFigures exported:")
    for output_path in outputs:
        print(f"  {output_path}")


if __name__ == "__main__":
    main()
