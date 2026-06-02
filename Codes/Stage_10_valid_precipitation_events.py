# Charlotte Grosjean - UNIL Master Thesis
"""
Identify precipitation events and visualize selected event typologies.

Events are detected on the full station network using only final quality-flag
0 measurements. A network event starts when at least one valid station measures
rain and ends when the next valid rainfall anywhere in the network is separated
by more than 90 dry minutes.

Outputs:
- one CSV with all validated network events;
- one CSV summary with station participation in the validated events;
- one CSV with selected representative event typologies;
- PNG and interactive HTML plots comparing available stations with MeteoSwiss LSN.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd

os.environ.setdefault("MPLCONFIGDIR", os.path.join(tempfile.gettempdir(), "matplotlib"))
os.environ.setdefault("XDG_CACHE_HOME", tempfile.gettempdir())

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt

from H_cross_correlation_utils import (
    distribute_meteo_to_target_grid,
    load_meteosuisse_source,
)
from I_paths_config import (
    CONSOLIDATED_DIRECTORY,
    METEOSUISSE_FILES,
    EVENTS_DIRECTORY,
)


# ===== FILES =====
OUTPUT_DIRECTORY = EVENTS_DIRECTORY


# ===== PLOT CONFIGURATION =====
PLOT_TITLE_FONTSIZE = 14
PLOT_LABEL_FONTSIZE = 12
PLOT_TICK_FONTSIZE = 10.5
PLOT_LEGEND_FONTSIZE = 10.5

plt.rcParams.update(
    {
        "axes.titlesize": PLOT_TITLE_FONTSIZE,
        "axes.labelsize": PLOT_LABEL_FONTSIZE,
        "xtick.labelsize": PLOT_TICK_FONTSIZE,
        "ytick.labelsize": PLOT_TICK_FONTSIZE,
        "legend.fontsize": PLOT_LEGEND_FONTSIZE,
    }
)


# ===== EVENT PARAMETERS =====
TIME_STEP_MINUTES = 3
DRY_SEPARATION_MINUTES = 90
MIN_EVENT_RAIN_MM = 1.0
MIN_STATION_RAIN_FOR_EVENT_MM = 0.1
PLOT_BUFFER_MINUTES = 90
MAX_SELECTED_EVENTS = 15
MIN_MULTI_STATION_SELECTED_EVENTS = 7
MIN_HIGH_AVAILABILITY_SELECTED_EVENTS = 10
MIN_RAINY_STATIONS_FOR_MULTI_STATION_SELECTION = 3
VALID_EVENT_QUALITY_FLAGS = {0}

# Only final validated measurements are used for event detection and figures.

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


def station_sort_key(station: str) -> tuple[int, str]:
    if station in STATION_ORDER:
        return STATION_ORDER.index(station), station
    return len(STATION_ORDER), station


def ensure_output_directory() -> None:
    (OUTPUT_DIRECTORY / "figures").mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIRECTORY / "interactive").mkdir(parents=True, exist_ok=True)


def list_consolidated_files() -> list[Path]:
    return sorted(
        CONSOLIDATED_DIRECTORY.glob("*_Consolidated_Complete.csv"),
        key=lambda path: station_sort_key(path.name.split("_")[0]),
    )


def load_consolidated_stations() -> dict[str, pd.DataFrame]:
    """Load consolidated station data needed for event detection."""
    station_data: dict[str, pd.DataFrame] = {}
    for file_path in list_consolidated_files():
        station = file_path.name.split("_")[0]
        df = pd.read_csv(
            file_path,
            usecols=lambda col: col in {"DateTime", "P_mm", "quality_flag"},
        )
        required = {"DateTime", "P_mm", "quality_flag"}
        if not required.issubset(df.columns):
            print(f"Skipping {file_path.name}: missing {sorted(required.difference(df.columns))}")
            continue

        df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
        df["P_mm"] = pd.to_numeric(df["P_mm"], errors="coerce")
        df["quality_flag"] = pd.to_numeric(df["quality_flag"], errors="coerce")
        df = df.dropna(subset=["DateTime", "P_mm", "quality_flag"])
        df = df.sort_values("DateTime").drop_duplicates("DateTime", keep="first")
        df = df.set_index("DateTime")
        station_data[station] = df[["P_mm", "quality_flag"]]
    return station_data


def has_valid_event_quality(flags: pd.Series) -> pd.Series:
    """Return True for quality flags accepted in event example figures."""
    return flags.isin(VALID_EVENT_QUALITY_FLAGS)


def detect_station_events(station: str, df: pd.DataFrame) -> list[dict[str, object]]:
    """Detect validated events for one station."""
    measurements = df[df["P_mm"] >= 0].copy()
    rainy_times = measurements.index[measurements["P_mm"] > 0]
    if len(rainy_times) == 0:
        return []

    dry_gap = pd.Timedelta(minutes=DRY_SEPARATION_MINUTES)
    events: list[dict[str, object]] = []
    cluster_start = rainy_times[0]
    previous_time = rainy_times[0]

    for current_time in rainy_times[1:]:
        if current_time - previous_time > dry_gap:
            event = validate_event_cluster(station, measurements, cluster_start, previous_time)
            if event is not None:
                events.append(event)
            cluster_start = current_time
        previous_time = current_time

    event = validate_event_cluster(station, measurements, cluster_start, previous_time)
    if event is not None:
        events.append(event)

    return events


def validate_event_cluster(
    station: str,
    measurements: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
) -> dict[str, object] | None:
    """Keep an event only if its interval has accepted quality flags."""
    event_window = measurements.loc[start:end]
    if event_window.empty:
        return None

    total_rain = float(event_window["P_mm"].sum())
    if total_rain < MIN_EVENT_RAIN_MM:
        return None

    if not has_valid_event_quality(event_window["quality_flag"]).all():
        return None

    duration_minutes = max(
        TIME_STEP_MINUTES,
        (end - start).total_seconds() / 60 + TIME_STEP_MINUTES,
    )
    rainy_steps = int((event_window["P_mm"] > 0).sum())
    peak_3min = float(event_window["P_mm"].max())
    mean_intensity_mm_h = total_rain / (duration_minutes / 60)

    return {
        "Station": station,
        "Event start": start,
        "Event end": end,
        "Duration min": round(duration_minutes, 1),
        "Rain total mm": round(total_rain, 4),
        "Peak 3min mm": round(peak_3min, 4),
        "Mean intensity mm/h": round(mean_intensity_mm_h, 4),
        "Rainy timesteps": rainy_steps,
    }


def station_is_valid_for_event(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> bool:
    """Return True if a station has complete accepted data over the event interval."""
    expected_index = pd.date_range(start=start, end=end, freq=f"{TIME_STEP_MINUTES}min")
    window = df.reindex(expected_index)
    if window.empty or window[["P_mm", "quality_flag"]].isna().any().any():
        return False
    return bool(
        (window["P_mm"] >= 0).all()
        and has_valid_event_quality(window["quality_flag"]).all()
    )


def valid_station_event_rain(df: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> float | None:
    """Return event rainfall if station has complete valid data, otherwise None."""
    expected_index = pd.date_range(start=start, end=end, freq=f"{TIME_STEP_MINUTES}min")
    window = df.reindex(expected_index)
    if window.empty or window[["P_mm", "quality_flag"]].isna().any().any():
        return None
    if not (
        (window["P_mm"] >= 0).all()
        and has_valid_event_quality(window["quality_flag"]).all()
    ):
        return None
    return float(window["P_mm"].sum())


def add_network_availability(
    events_df: pd.DataFrame,
    station_data: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """Add available station count/list for each event."""
    records = []
    for _, event in events_df.iterrows():
        start = pd.Timestamp(event["Event start"])
        end = pd.Timestamp(event["Event end"])
        available = []
        rainy = []
        station_totals = {}
        for station, df in station_data.items():
            station_rain = valid_station_event_rain(df, start, end)
            if station_rain is None:
                continue
            available.append(station)
            station_totals[station] = station_rain
            if station_rain >= MIN_STATION_RAIN_FOR_EVENT_MM:
                rainy.append(station)

        records.append(
            {
                "Available station count": len(available),
                "Available stations": ", ".join(sorted(available, key=station_sort_key)),
                "Rainy station count": len(rainy),
                "Rainy stations": ", ".join(sorted(rainy, key=station_sort_key)),
                "Station event totals mm": "; ".join(
                    f"{station}:{station_totals[station]:.2f}"
                    for station in sorted(station_totals, key=station_sort_key)
                ),
            }
        )
    return pd.concat([events_df.reset_index(drop=True), pd.DataFrame(records)], axis=1)


def build_valid_rain_matrix(station_data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Build a station x time matrix where values are valid precipitation only."""
    series_by_station = {}
    for station, df in station_data.items():
        valid = df["P_mm"].where(
            (df["P_mm"] >= 0) & has_valid_event_quality(df["quality_flag"])
        )
        series_by_station[station] = valid
    matrix = pd.concat(series_by_station, axis=1).sort_index()
    matrix = matrix.reindex(columns=sorted(matrix.columns, key=station_sort_key))
    return matrix


def detect_network_events(station_data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Detect events on the full station network instead of station by station.

    A network event starts when at least one valid station measures rain and ends
    when the next valid rainfall anywhere in the network is more than 90 minutes
    away. Station availability and rainfall participation are then counted over
    that common event interval.
    """
    rain_matrix = build_valid_rain_matrix(station_data)
    any_rain = rain_matrix.fillna(0).gt(0).any(axis=1)
    rainy_times = rain_matrix.index[any_rain]
    if len(rainy_times) == 0:
        return pd.DataFrame()

    dry_gap = pd.Timedelta(minutes=DRY_SEPARATION_MINUTES)
    clusters: list[tuple[pd.Timestamp, pd.Timestamp]] = []
    cluster_start = rainy_times[0]
    previous_time = rainy_times[0]
    for current_time in rainy_times[1:]:
        if current_time - previous_time > dry_gap:
            clusters.append((cluster_start, previous_time))
            cluster_start = current_time
        previous_time = current_time
    clusters.append((cluster_start, previous_time))

    records = []
    for event_id, (start, end) in enumerate(clusters, start=1):
        expected_index = pd.date_range(start=start, end=end, freq=f"{TIME_STEP_MINUTES}min")
        window = rain_matrix.reindex(expected_index)
        available_mask = window.notna().all(axis=0)
        available = list(window.columns[available_mask])
        if not available:
            continue

        totals = window[available].fillna(0).sum(axis=0)
        rainy_totals = totals[totals >= MIN_STATION_RAIN_FOR_EVENT_MM]
        if rainy_totals.empty:
            continue

        primary_station = str(rainy_totals.idxmax())
        max_station_total = float(rainy_totals.max())
        if max_station_total < MIN_EVENT_RAIN_MM:
            continue

        event_window = window[rainy_totals.index]
        peak_3min = float(event_window.max().max())
        duration_minutes = max(
            TIME_STEP_MINUTES,
            (end - start).total_seconds() / 60 + TIME_STEP_MINUTES,
        )

        records.append(
            {
                "Event ID": event_id,
                "Station": primary_station,
                "Primary station": primary_station,
                "Event start": start,
                "Event end": end,
                "Duration min": round(duration_minutes, 1),
                "Rain total mm": round(max_station_total, 4),
                "Network mean rain mm": round(float(totals.mean()), 4),
                "Network max rain mm": round(max_station_total, 4),
                "Network sum rain mm": round(float(totals.sum()), 4),
                "Peak 3min mm": round(peak_3min, 4),
                "Mean intensity mm/h": round(max_station_total / (duration_minutes / 60), 4),
                "Available station count": int(len(available)),
                "Available stations": ", ".join(sorted(available, key=station_sort_key)),
                "Rainy station count": int(len(rainy_totals)),
                "Rainy stations": ", ".join(sorted(rainy_totals.index, key=station_sort_key)),
                "Station event totals mm": "; ".join(
                    f"{station}:{totals[station]:.2f}"
                    for station in sorted(totals.index, key=station_sort_key)
                ),
            }
        )

    return pd.DataFrame(records)


def classify_events(events_df: pd.DataFrame) -> pd.DataFrame:
    """Assign simple typology labels used for automatic selection."""
    df = events_df.copy()
    duration_q75 = df["Duration min"].quantile(0.75)
    total_q75 = df["Rain total mm"].quantile(0.75)
    peak_q75 = df["Peak 3min mm"].quantile(0.75)
    intensity_q75 = df["Mean intensity mm/h"].quantile(0.75)

    labels = []
    for _, row in df.iterrows():
        if row["Duration min"] <= 120 and row["Peak 3min mm"] >= peak_q75:
            labels.append("short_intense")
        elif row["Duration min"] >= duration_q75 and row["Rain total mm"] >= total_q75:
            labels.append("long_high_total")
        elif row["Duration min"] >= duration_q75:
            labels.append("long_moderate")
        elif row["Rainy station count"] >= 3:
            labels.append("multi_station")
        elif row["Mean intensity mm/h"] >= intensity_q75:
            labels.append("high_intensity")
        else:
            labels.append("standard")

    df["Typology"] = labels
    return df


def select_representative_events(events_df: pd.DataFrame) -> pd.DataFrame:
    """Select representative events, prioritizing high station availability."""
    if events_df.empty:
        return events_df

    candidates = classify_events(events_df)
    selected_indices: list[int] = []

    high_availability_candidates = candidates.sort_values(
        [
            "Available station count",
            "Rainy station count",
            "Rain total mm",
            "Duration min",
        ],
        ascending=False,
    )
    selected_indices.extend(
        [
            int(idx)
            for idx in high_availability_candidates.index[
                : min(MIN_HIGH_AVAILABILITY_SELECTED_EVENTS, MAX_SELECTED_EVENTS)
            ]
        ]
    )

    multi_station_candidates = candidates[
        candidates["Rainy station count"] >= MIN_RAINY_STATIONS_FOR_MULTI_STATION_SELECTION
    ].sort_values(
        ["Rainy station count", "Available station count", "Rain total mm", "Duration min"],
        ascending=False,
    )
    selected_multi_station_count = int(
        candidates.loc[selected_indices, "Rainy station count"]
        .ge(MIN_RAINY_STATIONS_FOR_MULTI_STATION_SELECTION)
        .sum()
    )
    needed_multi_station = max(
        0,
        MIN_MULTI_STATION_SELECTED_EVENTS - selected_multi_station_count,
    )
    selected_indices.extend(
        [
            int(idx)
            for idx in multi_station_candidates.index
            if int(idx) not in selected_indices
        ]
        [:needed_multi_station]
    )

    priority = [
        ("short_intense", ["Peak 3min mm", "Rainy station count", "Available station count"]),
        ("long_high_total", ["Rain total mm", "Rainy station count", "Available station count"]),
        ("multi_station", ["Rainy station count", "Rain total mm", "Available station count"]),
        ("long_moderate", ["Duration min", "Rainy station count", "Available station count"]),
        ("high_intensity", ["Mean intensity mm/h", "Rainy station count", "Available station count"]),
    ]

    for typology, sort_columns in priority:
        if len(selected_indices) >= MAX_SELECTED_EVENTS:
            break
        subset = candidates[candidates["Typology"] == typology]
        if subset.empty:
            continue
        subset = subset.sort_values(sort_columns, ascending=False)
        next_indices = [idx for idx in subset.index if int(idx) not in selected_indices]
        if not next_indices:
            continue
        selected_indices.append(int(next_indices[0]))
        if len(selected_indices) >= MAX_SELECTED_EVENTS:
            break

    if len(selected_indices) < MAX_SELECTED_EVENTS:
        remaining = candidates.drop(index=selected_indices, errors="ignore")
        remaining = remaining.sort_values(
            ["Rainy station count", "Available station count", "Rain total mm"],
            ascending=False,
        )
        selected_indices.extend(
            [int(idx) for idx in remaining.index[: MAX_SELECTED_EVENTS - len(selected_indices)]]
        )

    selected = candidates.loc[selected_indices].drop_duplicates(
        subset=["Station", "Event start", "Event end"]
    )
    return selected.sort_values("Event start").reset_index(drop=True)


def load_lsn_reference(target_index: pd.DatetimeIndex) -> pd.Series:
    """Load MeteoSwiss LSN and distribute it onto the requested 3-min grid."""
    lsn_df = load_meteosuisse_source(METEOSUISSE_FILES["LSN"])
    return distribute_meteo_to_target_grid(lsn_df, target_index, TIME_STEP_MINUTES)


def build_event_plot_data(
    event: pd.Series,
    station_data: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, list[str], pd.Series]:
    """Return station precipitation and LSN reference for an event plot."""
    event_start = pd.Timestamp(event["Event start"])
    event_end = pd.Timestamp(event["Event end"])
    plot_start = event_start - pd.Timedelta(minutes=PLOT_BUFFER_MINUTES)
    plot_end = event_end + pd.Timedelta(minutes=PLOT_BUFFER_MINUTES)
    plot_index = pd.date_range(plot_start, plot_end, freq=f"{TIME_STEP_MINUTES}min")

    available_stations = [
        station.strip()
        for station in str(event["Available stations"]).split(",")
        if station.strip()
    ]
    plot_series = {}
    for station in available_stations:
        df = station_data[station].reindex(plot_index)
        series = df["P_mm"].where(
            (df["P_mm"] >= 0) & has_valid_event_quality(df["quality_flag"])
        )
        plot_series[station] = series

    station_plot_df = pd.DataFrame(plot_series, index=plot_index)
    rainy_stations = [
        station.strip()
        for station in str(event.get("Rainy stations", "")).split(",")
        if station.strip()
    ]
    lsn_series = load_lsn_reference(plot_index)
    return station_plot_df, rainy_stations, lsn_series


def safe_fragment(value: object) -> str:
    text = str(value)
    return "".join(char if char.isalnum() or char in "._-" else "_" for char in text)


def plot_selected_event(
    event: pd.Series,
    station_data: dict[str, pd.DataFrame],
    index: int,
) -> tuple[Path, Path | None]:
    """Save static PNG and interactive HTML for one selected event."""
    station_df, rainy_stations, lsn_series = build_event_plot_data(event, station_data)
    event_start = pd.Timestamp(event["Event start"])
    event_end = pd.Timestamp(event["Event end"])

    fig, ax = plt.subplots(figsize=(11, 5.8))
    for station in station_df.columns:
        is_rainy = station in rainy_stations
        ax.plot(
            station_df.index,
            station_df[station],
            linewidth=1.5 if is_rainy else 0.8,
            alpha=0.9 if is_rainy else 0.35,
            color=None if is_rainy else "#9e9e9e",
            linestyle="-" if is_rainy else "--",
            label=station,
        )
    ax.plot(
        lsn_series.index,
        lsn_series,
        color="black",
        linewidth=2.0,
        label="MeteoSwiss LSN",
    )
    ax.axvspan(event_start, event_end, color="#f0f0f0", alpha=0.55, zorder=0)
    ax.set_ylabel("Precipitation [mm / 3 min]")
    ax.set_xlabel("Time")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d %H:%M"))
    ax.tick_params(axis="x", rotation=35, labelsize=PLOT_TICK_FONTSIZE)
    ax.tick_params(axis="y", labelsize=PLOT_TICK_FONTSIZE)
    ax.grid(alpha=0.25)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=4, frameon=False)
    fig.tight_layout()

    base_name = (
        f"{index:02d}_{safe_fragment(event['Typology'])}_"
        f"{safe_fragment(event['Station'])}_{event_start:%Y%m%d_%H%M}"
    )
    png_path = OUTPUT_DIRECTORY / "figures" / f"{base_name}.png"
    fig.savefig(png_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    html_path = None
    try:
        import plotly.graph_objects as go

        figure = go.Figure()
        for station in station_df.columns:
            is_rainy = station in rainy_stations
            figure.add_trace(
                go.Scatter(
                    x=station_df.index,
                    y=station_df[station],
                    mode="lines",
                    name=station,
                    line=dict(
                        color=None if is_rainy else "#9e9e9e",
                        width=2.4 if is_rainy else 1.2,
                        dash="solid" if is_rainy else "dash",
                    ),
                    opacity=1.0 if is_rainy else 0.45,
                )
            )
        figure.add_trace(
            go.Scatter(
                x=lsn_series.index,
                y=lsn_series,
                mode="lines",
                name="MeteoSwiss LSN",
                line=dict(color="black", width=3),
            )
        )
        figure.add_vrect(
            x0=event_start,
            x1=event_end,
            fillcolor="rgba(180,180,180,0.25)",
            line_width=0,
        )
        figure.update_layout(
            xaxis_title="Time",
            yaxis_title="Precipitation [mm / 3 min]",
            hovermode="x unified",
            template="plotly_white",
            legend=dict(orientation="h", y=-0.25),
            font=dict(size=15),
            xaxis=dict(title_font=dict(size=17), tickfont=dict(size=14)),
            yaxis=dict(title_font=dict(size=17), tickfont=dict(size=14)),
            height=650,
            width=1200,
        )
        html_path = OUTPUT_DIRECTORY / "interactive" / f"{base_name}.html"
        figure.write_html(
            html_path,
            include_plotlyjs="cdn",
            config={"scrollZoom": True, "displaylogo": False},
        )
    except ImportError:
        html_path = None

    return png_path, html_path


def build_station_participation_summary(events_df: pd.DataFrame) -> pd.DataFrame:
    """Summarize event participation per station for network-detected events."""
    records = []
    for station in sorted(STATION_ORDER, key=station_sort_key):
        available_mask = events_df["Available stations"].fillna("").apply(
            lambda text: station in [item.strip() for item in str(text).split(",") if item.strip()]
        )
        rainy_mask = events_df["Rainy stations"].fillna("").apply(
            lambda text: station in [item.strip() for item in str(text).split(",") if item.strip()]
        )
        station_events = events_df[rainy_mask]
        records.append(
            {
                "Station": station,
                "N_rainy_events": int(rainy_mask.sum()),
                "N_available_events": int(available_mask.sum()),
                "Pct_network_events_rainy": round(rainy_mask.mean() * 100, 1)
                if len(events_df)
                else 0,
                "Pct_network_events_available": round(available_mask.mean() * 100, 1)
                if len(events_df)
                else 0,
                "Median_event_total_mm_when_rainy": (
                    station_events["Rain total mm"].median() if not station_events.empty else np.nan
                ),
                "Max_event_total_mm_when_rainy": (
                    station_events["Rain total mm"].max() if not station_events.empty else np.nan
                ),
            }
        )
    return pd.DataFrame(records)


def main() -> None:
    ensure_output_directory()

    print("Loading consolidated stations...", flush=True)
    station_data = load_consolidated_stations()
    print(f"Loaded {len(station_data)} stations.", flush=True)

    print("Detecting network events from all valid station data...", flush=True)
    events_df = detect_network_events(station_data)
    if events_df.empty:
        print("No validated precipitation event found.")
        return

    events_df = classify_events(events_df)
    events_df = events_df.sort_values("Event start").reset_index(drop=True)
    print(f"Network validated events: {len(events_df)}", flush=True)

    all_events_path = OUTPUT_DIRECTORY / "validated_precipitation_events.csv"
    events_df.to_csv(all_events_path, index=False)

    summary = build_station_participation_summary(events_df)
    summary_path = OUTPUT_DIRECTORY / "validated_event_count_by_station.csv"
    summary.to_csv(summary_path, index=False)

    selected = select_representative_events(events_df)
    selected_path = OUTPUT_DIRECTORY / "selected_representative_events.csv"
    selected.to_csv(selected_path, index=False)

    print("\nSelected representative events:")
    print(selected[[
        "Typology",
        "Station",
        "Event start",
        "Event end",
        "Rain total mm",
        "Duration min",
        "Available station count",
        "Rainy station count",
    ]].to_string(index=False))

    output_figures = []
    for idx, (_, event) in enumerate(selected.iterrows(), start=1):
        png_path, html_path = plot_selected_event(event, station_data, idx)
        output_figures.append(png_path)
        if html_path is not None:
            output_figures.append(html_path)

    print("\nOutputs:")
    print(f"  {all_events_path}")
    print(f"  {summary_path}")
    print(f"  {selected_path}")
    for output_path in output_figures:
        print(f"  {output_path}")


if __name__ == "__main__":
    main()
