# Charlotte Grosjean - 25.03.2026 - UNIL Master Thesis

"""
STATIONS PROCESSING PIPELINE

This script automates the complete processing workflow for pluviometer station 

1. **BEFORE Processing**: Scans raw .txt/.xlsx/.csv files → extracts stats (first/last date, count, timestep)
2. **AFTER Processing**: Re-scans processed files → updates Excel with final stats
3. **Comments Enrichment**: Merges Logbook comments (Pluvimates_Lausanne_Log.xlsx) + File comments (Comments.xlsx)

Files should be organized that way:
- data folder
    - 1_Stations (raw files). For the CSV files, the last downloaded version is kept here    
    - 1_Stations_Processed (processed files after CSV preprocessing and manual processing)
    - Pluvimates_Lausanne_Log.xlsx (logbook with comments and maintenance info)
    - Comments.xlsx (manually filled comments for specific files)
    - Sauvegarde_Processed (backup of old processed csv files before updating with new data)

Output: Available_Data.xlsx (formatted with highlights for overlaps/gaps, dynamic column widths, station grouping)
"""

# %%
import os
import pandas as pd
from glob import glob
from datetime import datetime
from A_data_processing_utils import *
from B_excel_utils import *
from B_excel_utils import colorize_comment_keywords
import shutil  # for moving files
from I_paths_config import (
    STATIONS_DIRECTORY,
    PROCESSED_DIRECTORY,
    PROCESSED_BACKUP_DIRECTORY,
    AVAILABLE_DATA_FILE,
    PLUVIMATES_LOG_PATH,
    FILE_COMMENTS_PATH,
)

# %% 1 - Before manually processing

# This section scans station folders for data files (.txt, .xlsx, .csv)
# and builds a summary Excel report with file statistics BEFORE MANUAL PROCESSING of the data. It takes 
# into account the raw data files as they are, without any cleaning or formatting, 
# to provide an overview of the initial state of the data.

base_directory = STATIONS_DIRECTORY
output_file = AVAILABLE_DATA_FILE
summary_rows_before = []

for folder_name in os.listdir(base_directory):
    folder_path = os.path.join(base_directory, folder_name)
    if os.path.isdir(folder_path):
        print(f"\nProcessing station BEFORE: {folder_name}")
        summary_row = process_station(folder_path, folder_name)  # function from data_processing_utils.py to process station files and extract stats
        summary_rows_before.extend(summary_row)

# Create and save Excel report
summary_df = pd.DataFrame(summary_rows_before)

# Sort by Station and First measurement_Before Processed
summary_df["File_date"] = summary_df["File name"].apply(
    lambda x: pd.to_datetime(os.path.splitext(x)[0].split("_")[-1], format="%Y%m%d", errors="coerce")
)
summary_df = summary_df.sort_values(by=["Station", "File_date"]).reset_index(drop=True)
summary_df = summary_df.drop(columns=["File_date"])

summary_df.to_excel(output_file, index=False)

format_excel_report(output_file)
print("Before processing completed!")


# %% 2 - After manually processing

# This section scans station folders for data files (.txt, .xlsx, .csv)
# and complete the summary Excel report with file statistics AFTER MANUAL PROCESSING of the data.

base_directory_after = PROCESSED_DIRECTORY
output_file = AVAILABLE_DATA_FILE

# Read existing BEFORE data and ensure AFTER columns exist
existing_df = pd.read_excel(output_file)
for col in ["First measurement AP", "Last measurement AP"]:
    if col not in existing_df.columns:
        existing_df[col] = pd.NaT 

for col in ["Number of measurements AP", "Duplicates AP"]:    
    if col not in existing_df.columns:        
        existing_df[col] = np.nan 

# Process AFTER files
summary_rows_after = []
for folder_name in os.listdir(base_directory_after):
    folder_path = os.path.join(base_directory_after, folder_name)
    if os.path.isdir(folder_path):
        print(f"\nProcessing station AFTER: {folder_name}")
        rows = process_station_after(folder_path, folder_name)
        summary_rows_after.extend(rows)

summary_after_df = pd.DataFrame(summary_rows_after)

# Merge and update existing Excel with AFTER data
if not summary_after_df.empty:

    updated_df = existing_df.merge(
        summary_after_df, on=["Station", "File name", "File type"],
        how="left", suffixes=("", "_new")
    )

    datetime_cols = ["First measurement AP", "Last measurement AP"]
    numeric_cols = ["Number of measurements AP", "Duplicates AP"]

    for col in datetime_cols + numeric_cols:
        new_col = col + "_new"
        
        if new_col in updated_df.columns:
            if col in datetime_cols:
                updated_df[col] = pd.to_datetime(updated_df[col], errors="coerce")
                updated_df[new_col] = pd.to_datetime(updated_df[new_col], errors="coerce")
            else:
                updated_df[col] = pd.to_numeric(updated_df[col], errors="coerce")
                updated_df[new_col] = pd.to_numeric(updated_df[new_col], errors="coerce")

            updated_df[col] = updated_df[col].combine_first(updated_df[new_col])
            updated_df.drop(columns=[new_col], inplace=True)

# Sort by Station and First measurement_Before Processed
    updated_df["File_date"] = updated_df["File name"].apply(
        lambda x: pd.to_datetime(os.path.splitext(x)[0].split("_")[-1], format="%Y%m%d", errors="coerce")
    )
    updated_df = updated_df.sort_values(by=["Station", "File_date"]).reset_index(drop=True)
    updated_df = updated_df.drop(columns=["File_date"])

    updated_df.to_excel(output_file, index=False)
    format_excel_report(output_file)
    print("After processing completed!")
else:
    print("No AFTER data found to update.")

# %% 3 - enrich with comments from Logbook + manually processing comments

available_data_path = AVAILABLE_DATA_FILE
pluvimates_log_path = PLUVIMATES_LOG_PATH
file_comments_path = FILE_COMMENTS_PATH

# Load data
available_df = pd.read_excel(available_data_path)
logbook_df = pd.read_excel(pluvimates_log_path)
comments_df = pd.read_excel(file_comments_path)

# Add logbook comments
available_df = add_logbook_comments(available_df, logbook_df, date_col="Horodateur")

# Add file-based comments (Comments.xlsx)
available_df = add_file_comments(available_df, comments_df)

# Save back to Excel (data only)
available_df.to_excel(available_data_path, index=False)

# Apply formatting (including wrap & width for comments columns)
format_excel_report(available_data_path)

# Colorize keywords in comments
colorize_comment_keywords(available_data_path)

print("Available_Data.xlsx updated with Logbook comments + Comments.xlsx info.")


# %%
