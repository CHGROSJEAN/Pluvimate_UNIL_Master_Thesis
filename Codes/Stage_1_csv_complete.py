# Charlotte Grosjean - 25.03.2026 - UNIL Master Thesis
"""
CSV FILES PREPROCESSING

This script:
1. Processes raw .csv files from pluviometer stations by: 
    - Reads raw .csv files from stations
    - Converts cumulative to time steps precipitation
    - Removes empty and negative precipitation values
    - Saves processed .csv files for further consolidation  
2. Combines old processed .csv with new .csv files incrementally, keeping only new data based on the last measurement date in the old file. 
The new data added are not manually processed but the manually processed data from the previous file is kept. 
This way, we can keep an updated version of the processed data without having to manually process the beginning of the new .csv files every time.
    - Combines old and new .csv files incrementally (adding only new data to the previous processed file)
    - Backs up old .csv files to a separate folder

At the end, we have: 
- Last raw data .csv files in 1_Stations (last downloaded version)
- Combined processed .csv files in 1_Stations_Processed
- Backed-up previous processed .csv files in Sauvegarde_Processed
"""

import os
import pandas as pd
from glob import glob
from datetime import datetime
from A_data_processing_utils import *
from B_excel_utils import *
import shutil  # for moving files 
from I_paths_config import STATIONS_DIRECTORY, PROCESSED_DIRECTORY, PROCESSED_BACKUP_DIRECTORY


# Input folder (raw files)
input_base_directory = STATIONS_DIRECTORY

# Output folder (processed files)
processed_base_directory = PROCESSED_DIRECTORY
backup_directory = PROCESSED_BACKUP_DIRECTORY
os.makedirs(processed_base_directory , exist_ok=True)

# %% 1 - Processing the .csv files. - remove empty precipitation rows and convert cumulative to incremental

for folder_name in os.listdir(input_base_directory):
    input_folder = os.path.join(input_base_directory, folder_name)
    output_folder = os.path.join(processed_base_directory , folder_name)
    if not os.path.isdir(input_folder):
        continue

    os.makedirs(output_folder, exist_ok=True)
    print(f"\n=== Processing station: {folder_name} ===")

    for csv_file in glob(os.path.join(input_folder, "*.csv")):
        df = pd.read_csv(csv_file)

        # Remove rows where field5 (precipitation) is empty
        df['field5'] = df['field5'].apply(extract_numeric)
        df = df.dropna(subset=['field5'])

        # Convert cumulative precipitation to incremental by taking the difference between consecutive rows
        df['field5'] = df['field5'].diff().fillna(0)

        # Remove negative precipitation values
        df = df[df['field5'] >= 0] 

        # --- SAVE TO PROCESSED FOLDER ---
        output_file = os.path.join(output_folder, os.path.basename(csv_file))
        df.to_csv(output_file, index=False)

        print(f"Processed {os.path.basename(csv_file)}: {len(df)} rows")

# %% 2 - Combine processed files (incremental update)

os.makedirs(backup_directory, exist_ok=True)  # create backup folder if it doesn't exist

for folder_name in os.listdir(processed_base_directory):
    folder_path = os.path.join(processed_base_directory, folder_name)

    if not os.path.isdir(folder_path):
        continue

    print(f"\n=== Combining station: {folder_name} ===")

    csv_files = glob(os.path.join(folder_path, "*.csv"))

    # Check if there are at least 2 files to combine (old + new)
    if len(csv_files) < 2:
        print("Not enough files to combine.")
        continue

    # Sort files by date in the filename
    csv_files_sorted = sorted(csv_files, key=lambda x: extract_date_from_filename(os.path.basename(x)))

    old_file = csv_files_sorted[0]  # oldest file
    new_file = csv_files_sorted[-1] # newest file

    print(f"Old: {os.path.basename(old_file)}")
    print(f"New: {os.path.basename(new_file)}")

    df_old = pd.read_csv(old_file)
    df_new = pd.read_csv(new_file)

    # Filter new data to only include rows with created_at > last created_at in old file (new data)
    last_date = df_old['created_at'].max()
    df_new_filtered = df_new[df_new['created_at'] > last_date]

    if df_new_filtered.empty:
        print("No new data to add.")
        continue

    # Combine old and new data (keep all old data + only new data from the new file)
    df_combined = pd.concat([df_old, df_new_filtered], ignore_index=True) # ingnore_index to reset index after concatenation

    # Save combined file with the same name as the newest raw CSV file in the processed folder 
    # (overwriting the previous processed file)
    new_filename = os.path.basename(new_file)
    output_file = os.path.join(folder_path, new_filename)
    df_combined.to_csv(output_file, index=False)
    print(f"Saved combined file as: {os.path.basename(output_file)} ({len(df_combined)} rows)")

    # Move old processed files to backup folder 
    for f in csv_files_sorted[:-1]:  # all files except the newest
        dest_file = os.path.join(backup_directory, os.path.basename(f))
        shutil.move(f, dest_file)
        print(f"Moved old file to backup: {os.path.basename(f)}")
