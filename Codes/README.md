# MasterThesis Pluvimate Rainfall Processing Pipeline

This repository contains the Python workflow used to process Pluvimate station
precipitation data, compare it with MeteoSwiss references, assign quality
flags, and extract validated rainfall events for analysis.

The pipeline is organized as numbered stages (`Stage_1` to `Stage_10`) plus
shared helper modules (`A` to `I`). Run the stages in numerical order unless
you only need to regenerate a downstream product from already existing inputs.

## Environment

The project was tested with Python 3.11.0. Install the required packages from
the pinned environment file:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

To run one stage from the repository root:

```bash
python Stage_6_cross_correlation.py
```

In a Jupyter or VS Code interactive window, prefer:

```python
%run "/path/to/MasterThesis/Stage_6_cross_correlation.py"
```

Avoid running old copied cells from notebooks when regenerating official
outputs, because stale functions can overwrite correct results.

## Data Layout

All paths are centralized in `I_paths_config.py`. The default base data folder
is:

```text
3_Data/
```

Main input and output folders:

```text
1_Stations/                         raw Pluvimate station files
1_Stations_Processed/               cleaned/processed station files
1_Stations_MeteoSuisse/             MeteoSwiss reference files
2_Overlaps/                         overlap diagnostics between source files
3_Consolidated_Stations/            merged station time series with gaps
4_Grids/4.1_Availability/           station availability grids
4_Grids/4.2_Precipitation/          precipitation grids
5_Cross_correlation/                MeteoSwiss correlation outputs
6_Correction_Factor/                intensity consistency diagnostics
7_Quality_Control/                  quality-control reports/statistics
8_Quality_Control_Figures/          quality-control figures
9_Validated_Precipitation_Events/   validated event tables and figures
```

## Overall Workflow

| Stage | Script | Purpose | Main Outputs |
|---|---|---|---|
| 1 | `Stage_1_csv_complete.py` | Preprocess raw CSV downloads, convert cumulative rain to timestep rain, remove invalid values, and incrementally combine new CSV data with previously processed files. | Updated CSV files in `1_Stations_Processed/`; backups in `Z_Sauvegarde_Processed/`. |
| 2 | `Stage_2_stations_processing.py` | Inventory raw and processed files, compute file-level statistics, and merge logbook/manual comments. | `Available_Data.xlsx`. |
| 3 | `Stage_3_overlap_processing.py` | Detect overlapping periods between source files for each station, align paired sources, compute comparison metrics, and create overlap plots. | Overlap CSVs, comparison summaries, and figures in `2_Overlaps/`. |
| 4 | `Stage_4_consolidate_stations_files.py` | Merge all processed files per station, standardize precipitation, and fill missing 3-minute timesteps with gap codes. | `*_Consolidated_Complete.csv` files in `3_Consolidated_Stations/`. |
| 5 | `Stage_5_station_grids.py` | Build shared 3-minute availability, precipitation, and gap-code grids across stations. | Grid files, summaries, and heatmaps in `4_Grids/4.1_Availability/` and `4_Grids/4.2_Precipitation/`. |
| 6 | `Stage_6_cross_correlation.py` | Compare MeteoSwiss stations with each other and correlate each processed Pluvimate file against LSN, PUY, and VIT over a lag range. | `meteosuisse_three_station_cross_correlation.csv`; `lausanne_cross_correlation_by_file.csv/.xlsx`; diagnostic figures. |
| 7 | `Stage_7_station_intensity_correction.py` | Evaluate rainfall-intensity consistency using the Stage 6 best MeteoSwiss reference. It computes correction factors but does not apply corrections. | `correction_factors_analysis.xlsx`; updates/diagnostics for correction-factor interpretation. |
| 8 | `Stage_8_quality_flags.py` | Combine correlation, intensity, and daily MeteoSwiss checks to assign initial and final quality flags. Updates consolidated station CSVs with per-timestep flags. | Updated `Available_Data.xlsx`; `Quality_Control_report.xlsx`; `Quality_Control_statistics.xlsx`; daily quality reports. |
| 9 | `Stage_9_quality_flags_visualization.py` | Visualize final quality flags, daily precipitation, monthly quality timelines, station-level flag distributions, and daily diagnostics. | PNG/PDF/HTML figures and CSV summaries in `8_Quality_Control_Figures/`. |
| 10 | `Stage_10_valid_precipitation_events.py` | Detect validated network rainfall events using only final quality-flag `0` Pluvimate data, then select representative events for figures. | Event CSVs plus PNG/HTML figures in `9_Validated_Precipitation_Events/`. |

## Quality Flags

Final timestep flags used by Stages 8 to 10:

```text
-2  intra-file gap
-1  inter-file gap
 0  valid
 1  uncertain
 2  rejected file
 3  rejected day
 4  outlier
```

Stage 10 uses only `quality_flag == 0` Pluvimate measurements for event
detection and event figures. MeteoSwiss LSN is plotted as an external reference
and does not carry Pluvimate quality flags.

## Helper Modules A-I

| Module | Role |
|---|---|
| `A_data_processing_utils.py` | Date parsing, numeric extraction, file statistics, station inventory helpers, logbook/comment merging. |
| `B_excel_utils.py` | Excel formatting for reports, including headers, widths, borders, conditional highlighting, and comment styling. |
| `C_overlap_utils.py` | Overlap loading, cleaning, period detection, alignment, metrics, and overlap visualizations. |
| `D_time_convertion.py` | Manual timezone conversion and timestamp correction utilities for special raw-file cases. Run only the relevant section when needed. |
| `E_plot_config.py` | Shared plot style constants used by overlap plots. |
| `F_rain_quality_plot_utils.py` | Shared precipitation color scale and rain-color utilities for quality-control figures. |
| `G_station_grids_utils.py` | Core grid-building utilities for availability, precipitation, gap codes, summaries, and heatmaps. |
| `H_cross_correlation_utils.py` | Shared loading, regularization, MeteoSwiss redistribution, and correlation utilities used by Stages 6 to 8. |
| `I_paths_config.py` | Centralized path and shared configuration constants for all pipeline stages. |

## Important Parameters

Key thresholds are defined near the top of the relevant stage files:

```text
Stage 6: correlation lag range and minimum valid pairs
Stage 7: common rainy timestep minimum and correction-factor guardrails
Stage 8: correlation, correction-factor, daily-check, and outlier thresholds
Stage 10: event rain threshold, dry separation, and selected-event counts
```

For reporting, the correction-factor quality decision is based on the distance
from 1:

```text
D_I = |factor - 1|
```

This value describes the relative intensity mismatch between Pluvimate and the
selected MeteoSwiss reference. The factor is diagnostic only; the pipeline does
not correct the Pluvimate measurements.

## Reproducibility Notes

- Keep `requirements.txt` with the code when sharing the project.
- Do not share `.venv/`; recreate it from `requirements.txt`.
- Run stages from the repository root so relative imports resolve consistently.
- If using notebooks, restart the kernel before running a full stage with
  `%run`.
- Stage outputs can be regenerated independently only if all upstream inputs
  already exist and are up to date.
- `I_paths_config.py` contains local absolute paths. Update `DATA_DIRECTORY`
  before running the pipeline on another machine.

