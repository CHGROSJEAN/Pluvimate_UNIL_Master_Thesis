# Charlotte Grosjean - UNIL Master Thesis
"""
Centralized path configuration for all pipeline stages.

This module defines all directory and file paths used across the pipeline.
Import this module in all stages to ensure consistency and allow single-point updates.

Usage:
    from I_paths_config import DATA_DIRECTORY, PROCESSED_DIRECTORY, ...
    
All paths are defined relative to the main data directory and use Path() for cross-platform compatibility.
"""

from pathlib import Path

# ===== BASE DIRECTORIES =====
# Main data directory - base path for all pipeline data
DATA_DIRECTORY = Path(
    "/Users/charlottegrosjean/Library/Mobile Documents/com~apple~CloudDocs/2UNIL/Master/"
    "Master thesis/3_Data"
)

# ===== INPUT DIRECTORIES (Raw & Processed Station Data) =====
# Raw station data (CSV, TXT, XLSX files)
STATIONS_DIRECTORY = DATA_DIRECTORY / "1_Stations"

# Processed station files (cumulative → incremental conversion applied)
PROCESSED_DIRECTORY = DATA_DIRECTORY / "1_Stations_Processed"

# MeteoSwiss reference station data
METEOSUISSE_DIRECTORY = DATA_DIRECTORY / "1_Stations_MeteoSuisse"

# ===== OUTPUT DIRECTORIES (Pipeline Results) =====
# Consolidated station files (merged and gap-filled)
CONSOLIDATED_DIRECTORY = DATA_DIRECTORY / "3_Consolidated_Stations"

# Overlap analysis results
OVERLAPS_DIRECTORY = DATA_DIRECTORY / "2_Overlaps"

# Grid outputs (availability and precipitation grids)
AVAILABILITY_DIRECTORY = DATA_DIRECTORY / "4_Grids" / "4.1_Availability"
PRECIPITATION_DIRECTORY = DATA_DIRECTORY / "4_Grids" / "4.2_Precipitation"

# Cross-correlation analysis results
CROSS_CORRELATION_DIRECTORY = DATA_DIRECTORY / "5_Cross_correlation"

# Quality control visualizations
QUALITY_FIGURES_DIRECTORY = DATA_DIRECTORY / "8_Quality_Control_Figures"

# Validated precipitation events
EVENTS_DIRECTORY = DATA_DIRECTORY / "9_Validated_Precipitation_Events"

# Daily quality reports (from Stage 8)
DAILY_REPORTS_DIRECTORY = DATA_DIRECTORY / "7_Quality_Control" / "daily_processed_quality_reports"

# ===== REPORT/SUMMARY FILES (Excel, CSV, etc.) =====
# Master availability and station inventory
AVAILABLE_DATA_FILE = DATA_DIRECTORY / "Available_Data.xlsx"

# Logbook and comments files
PLUVIMATES_LOG_PATH = DATA_DIRECTORY / "Pluvimates_Lausanne_Log.xlsx"
FILE_COMMENTS_PATH = DATA_DIRECTORY / "Comments.xlsx"


# Cross-correlation summary (Stage 6 output used by Stage 8)
CORRELATION_SUMMARY_FILE = (
    CROSS_CORRELATION_DIRECTORY / "Cross_correlation_MeteoSuisse" / "Summary.xlsx"
)
FALLBACK_CORRELATION_SUMMARY_FILE = (
    CROSS_CORRELATION_DIRECTORY / "Cross_correlation_MeteoSuisse" / "lausanne_cross_correlation_by_file.xlsx"
)

# Correction factors analysis (Stage 7 output used by Stage 8)
CORRECTION_FACTORS_FILE = DATA_DIRECTORY / "6_Correction_Factor" / "correction_factors_analysis.xlsx"

# Quality control decision reports (Stage 8 outputs)
QUALITY_CONTROL_DECISIONS_FILE = DATA_DIRECTORY / "7_Quality_Control" / "Quality_Control_report.xlsx"
QUALITY_CONTROL_STATISTICS_FILE = DATA_DIRECTORY / "7_Quality_Control" / "Quality_Control_statistics.xlsx"

# ===== BACKUP DIRECTORIES =====
# Backup storage for processed files
PROCESSED_BACKUP_DIRECTORY = DATA_DIRECTORY / "Z_Sauvegarde_Processed"

# Time-stamped backups of consolidated stations before quality flag edits
CONSOLIDATED_BACKUP_DIRECTORY_PREFIX = "2_Consolidated_Stations_quality_flag_backup"

# Time-stamped backups of processed files before quality flag edits
PROCESSED_BACKUP_DIRECTORY_PREFIX = "1_Stations_Processed_quality_flag_backup"

# ===== METEOSWISS FILES MAPPING =====
# Dictionary mapping MeteoSwiss station codes to their file paths
# Used for cross-correlation analysis and quality comparisons
METEOSUISSE_FILES = {
    "LSN": str(METEOSUISSE_DIRECTORY / "LSN_Complete.csv"),      # Lausanne reference
    "PUY": str(METEOSUISSE_DIRECTORY / "PUY_Complete.csv"),      # Pully
    "VIT": str(METEOSUISSE_DIRECTORY / "VIT_Complete.csv"),      # Vitznau
}

# ===== HELPER FUNCTIONS =====
def ensure_output_directories() -> None:
    """Create all output directories if they don't exist."""
    output_dirs = [
        CONSOLIDATED_DIRECTORY,
        OVERLAPS_DIRECTORY,
        AVAILABILITY_DIRECTORY,
        PRECIPITATION_DIRECTORY,
        CROSS_CORRELATION_DIRECTORY,
        QUALITY_FIGURES_DIRECTORY,
        EVENTS_DIRECTORY,
        DAILY_REPORTS_DIRECTORY,
    ]
    for directory in output_dirs:
        directory.mkdir(parents=True, exist_ok=True)

def get_backup_directory(base_prefix: str) -> Path:
    """
    Generate timestamped backup directory name.
    
    Args:
        base_prefix: Prefix for the backup directory (e.g., "2_Consolidated_Stations_quality_flag_backup")
    
    Returns:
        Path to timestamped backup directory
    """
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return DATA_DIRECTORY / f"{base_prefix}_{timestamp}"

# ===== PATH CONSTANTS FOR SPECIFIC FEATURES =====
# Time resolution parameter (used across all stages)
GRID_STEP_MINUTES = 3
DEFAULT_SOURCE_STEP_MINUTES = 3
METEOSWISS_STEP_MINUTES = 10

# Resampling settings for grid construction
RESAMPLE_LABEL = 'right'
RESAMPLE_CLOSED = 'right'

# Quality flag thresholds (defined for reference, but should be in individual stage files for modification)
# These are copies - actual thresholds should be in the stage files
CORRELATION_VALID_MIN = 0.5
CORRELATION_KEEP_MIN = 0.35
CORRECTION_VALID_MAX_RELATIVE_DIFF = 0.20
CORRECTION_KEEP_MAX_RELATIVE_DIFF = 0.50
DAILY_CORRELATION_MIN = 0.40
DAILY_BIAS_MAX = 0.30
DAILY_MIN_VALID_POINTS = 40
OUTLIER_RAIN_THRESHOLD_MM_PER_TIMESTEP = 15.0
