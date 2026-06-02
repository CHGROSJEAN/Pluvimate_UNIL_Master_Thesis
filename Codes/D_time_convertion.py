# Charlotte Grosjean - 10.04.2026 - UNIL Master Thesis
"""
This script performs three main functions related to time conversion for station data files:
1. It converts the date and time columns of the data files to UTC timezone, taking into account daylight saving time adjustments for the Europe/Zurich timezone. 
   The converted files are saved with "_UTC" appended to their original filenames for easy identification and checking. 
   The original files remain unchanged for reference.
2. It corrects timestamps for station files that have incorrect date assignment at midnight. Specifically,
3. A dedicated correction is provided for LEXP_20221125.txt, where days start at 01:00 and end at 00:55.
   All timestamps from 00:00 to 00:55 are shifted to the next day. The corrected files are saved with "_corrected" appended to their original filenames.

Depending on what need to be done, you can run the import libraries cell and just the relevant section (UTC conversion or correction of timestamps).
"""

# %% import libraries

import pandas as pd
import os
import numpy as np

from A_data_processing_utils import parse_date, combine_excel_cell_datetime
from openpyxl import load_workbook
from datetime import datetime, time, timedelta


# File to convert (change this to the file you want to convert)
File = '/Users/charlottegrosjean/Library/Mobile Documents/com~apple~CloudDocs/2UNIL/Master/Master thesis/3_Data/1_Stations/BETH/BETH_20260401.txt'


#%% 1 - UTC conversion 
""" 
This section converts the date and time columns of the data files to UTC timezone, 
taking into account daylight saving time adjustments for the Europe/Zurich timezone. 
The converted files are saved with "_UTC" appended to their original filenames for easy identification and checking. 
The original files remain unchanged for reference.
"""

# Detect file type and process accordingly
if File.endswith('.xlsx'):
    print("Processing Excel file (.xlsx)...")
    
    # read the header (first 6 rows) to keep it for the output file
    header = pd.read_excel(File, header=None, nrows=6)

    # read the data starting from row 7 (skip the header)
    df = pd.read_excel(File, skiprows=7)

    df.columns = ["Date", "Heure", "Pulsations (Pulses)"]

    # Excel serial → datetime
    df["DateTime"] = pd.to_datetime(df["Date"])

    # Localize to Europe/Zurich and convert to UTC
    df["DateTime"] = df["DateTime"].dt.tz_localize(
        "Europe/Zurich",
        ambiguous="infer",
        nonexistent="shift_forward"
    )

    df["DateTime"] = df["DateTime"].dt.tz_convert("UTC")

    # convert UTC datetime to Excel serial format 
    excel_serial = (df["DateTime"] - pd.Timestamp("1899-12-30", tz="UTC")) / pd.Timedelta(days=1)

    df["Date"] = excel_serial
    df["Heure"] = excel_serial

    df = df.drop(columns=["DateTime"])

    # Save the updated dataframe back to an Excel file (with the same header)
    output_file = File.replace(".xlsx", "_UTC.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        header.to_excel(writer, index=False, header=False)
        df.to_excel(writer, index=False, startrow=6)
        
        # Set display formats for Date and Heure columns to have the same format as before (dd.mm.yyyy for Date and hh:mm:ss for Heure)
        worksheet = writer.book.active
        for row in range(7, 7 + len(df)):
            worksheet.cell(row=row, column=1).number_format = 'dd.mm.yyyy'  # Date column
            worksheet.cell(row=row, column=2).number_format = 'hh:mm:ss'   # Heure column

    print("✔ UTC conversion completed (Excel serial format)")

elif File.endswith('.txt'):
    print("Processing text file (.txt)...")
    
    # Read the file to extract metadata header (first 5 lines)
    with open(File, 'r') as f:
        header_lines = []
        data_start_line = 0
        for i, line in enumerate(f):
            if i < 5:
                header_lines.append(line)
                data_start_line = i + 1
            else:
                break
    
    # Read the data starting after the header (skip the first 5 lines)
    # Use header=None to prevent pandas from treating the first data row as headers
    df = pd.read_csv(File, sep='\t', skiprows=5, header=None, names=['Index', 'Time', 'Value'])
    
    # Convert the time column from string to datetime
    df['Time'] = pd.to_datetime(df['Time'])
    
    # Localize to Europe/Zurich (UTC+1) and convert to UTC
    # Note: This handles daylight saving time adjustments
    df['Time'] = df['Time'].dt.tz_localize(
        "Europe/Zurich",
        ambiguous="infer",
        nonexistent="shift_forward"
    )
    df['Time'] = df['Time'].dt.tz_convert("UTC")
    
    # Convert back to string format (without timezone info for text file)
    df['Time'] = df['Time'].dt.strftime('%d %b %Y %H:%M:%S')
    
    # Save the updated dataframe back to a text file with the same header
    output_file = File.replace(".txt", "_UTC.txt")
    
    with open(output_file, 'w') as f:
        # Write the header
        for header_line in header_lines:
            f.write(header_line)
        # Write the data (without index and without header)
        df.to_csv(f, sep='\t', index=False, header=False)
    
    print(f"✔ UTC conversion completed (Europe/Zurich timezone with DST)")
    print(f"Output file: {output_file}")

else:
    print("Error: File must be .xlsx or .txt format")


#%% 2 - UTC conversion with fixed offset (no DST)

# UTC offset in hours
#UTC+1 → UTC: UTC_OFFSET_HOURS = 1
#UTC+2 → UTC: UTC_OFFSET_HOURS = 2
#UTC-1 → UTC: UTC_OFFSET_HOURS = -1
#UTC-5 → UTC: UTC_OFFSET_HOURS = -5

UTC_OFFSET_HOURS = 1

# ============================================================================

# Detect file type and process accordingly
file_root, file_extension = os.path.splitext(File)
file_extension = file_extension.lower()

if file_extension in ['.xlsx', '.xls']:
    print(f"Processing Excel file ({file_extension}) with fixed UTC{UTC_OFFSET_HOURS:+d} offset (no DST)...")

    if file_extension == '.xls':
        raise ValueError("Exact Excel formatting preservation is only supported for .xlsx files")

    output_file = f"{file_root}_UTC.xlsx"
    workbook = load_workbook(File)
    worksheet = workbook.active

    for row in range(8, worksheet.max_row + 1):
        date_cell = worksheet.cell(row=row, column=1)
        time_cell = worksheet.cell(row=row, column=2)

        if date_cell.value is None:
            continue

        # This function from data_processing_utils.py combines the date and time from the two cells into a single datetime object. 
        datetime_value = combine_excel_cell_datetime(date_cell.value, time_cell.value)                                                                                                  
        if datetime_value is None:
            continue

        utc_datetime = datetime_value - timedelta(hours=UTC_OFFSET_HOURS)
        date_cell.value = utc_datetime
        time_cell.value = utc_datetime

    workbook.save(output_file)
    workbook.close()

    print(f"✔ UTC conversion completed (Fixed UTC{UTC_OFFSET_HOURS:+d} offset → UTC, no DST)")
    print(f"Output file: {output_file}")

elif file_extension == '.txt':
    print(f"Processing text file (.txt) with fixed UTC{UTC_OFFSET_HOURS:+d} offset (no DST)...")
    
    # Read the file to extract metadata header (first 5 lines)
    with open(File, 'r') as f:
        header_lines = []
        for i, line in enumerate(f):
            if i < 5:
                header_lines.append(line)
            else:
                break
    
    # Read the data starting after the header (skip the first 5 lines)
    # Use header=None to prevent pandas from treating the first data row as headers
    df = pd.read_csv(File, sep='\t', skiprows=5, header=None, names=['Index', 'Time', 'Value'])
    
    # Convert the time column from string to datetime using the shared parser
    # from data_processing_utils.py (handles French and English month names).
    raw_time = df['Time'].copy()
    df['Time'] = raw_time.apply(parse_date)
    if df['Time'].isna().any():
        examples = raw_time.loc[df['Time'].isna()].head(3).tolist()
        raise ValueError(f"Could not parse some datetime values. Examples: {examples}")
    
    # Convert fixed UTC offset to UTC
    # To convert UTC+X to UTC: subtract X hours
    # To convert UTC-X to UTC: add X hours (subtract negative value)
    df['Time'] = df['Time'] - pd.Timedelta(hours=UTC_OFFSET_HOURS)
    
    # Convert back to string format (without timezone info for text file)
    df['Time'] = df['Time'].dt.strftime('%d %b %Y %H:%M:%S')
    
    # Save the updated dataframe back to a text file with the same header
    output_file = f"{file_root}_UTC.txt"
    
    with open(output_file, 'w') as f:
        # Write the header
        for header_line in header_lines:
            f.write(header_line)
        # Write the data (without index and without header)
        df.to_csv(f, sep='\t', index=False, header=False)
    
    print(f"✔ UTC conversion completed (Fixed UTC{UTC_OFFSET_HOURS:+d} offset → UTC, no DST)")
    print(f"Output file: {output_file}")

else:
    print("Error: File must be .xlsx, .xls, or .txt format")


#%% 3. Correction of timestamps for station files with incorrect date assignment at midnight

# --- Download file ---
output_path = File.replace('.txt', '_corrected.txt')

# Ignore the first 5 lines of metadata header and read the data
df = pd.read_csv(File, sep='\t', skiprows=5, header=None, names=['Index', 'DateTime', 'Precipitation'])

# Parse the DateTime column to datetime objects (if not already parsed)
df['DateTime'] = df['DateTime'].apply(parse_date)

print("Before correction :")
print(df.head(15).to_string())

# --- Correction : all timestamps from 00:00 to 00:55 → next day ---
mask = (df['DateTime'].dt.hour == 0) & (df['DateTime'].dt.minute <= 55)
df.loc[mask, 'DateTime'] = df.loc[mask, 'DateTime'] + pd.Timedelta(days=1)

print("\nAfter correction :")
print(df.head(15).to_string())

# --- Save ---
df.to_csv(output_path, sep='\t', index=False)
print(f"\nCorrected file saved : {output_path}")
