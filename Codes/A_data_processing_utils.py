# Charlotte Grosjean - 23.03.2026 - UNIL Master Thesis

"""
DATA PROCESSING UTILS
Utility functions for processing station data files:
- parse_date: robust date parsing for multiple formats
- extract_numeric: extract numeric precipitation values from strings
- compute_stats: compute first/last measurement dates, count, and median time step
- process_file_txt/xlsx/csv: process individual files and compute stats
- process_station: process all files in a station directory and compile summary stats
- add_logbook_comments: merge logbook comments into summary based on station and date
- add_file_comments: merge manual comments from Comments.xlsx into summary based on station and file name
"""


import os
import pandas as pd
import numpy as np
from glob import glob
from datetime import datetime
from copy import copy
from openpyxl import load_workbook
import re

# -------------------------
# UTILITY FUNCTIONS
# -------------------------

def parse_date(date_string):
    """Parse multiple date formats into a standardized pandas datetime object. 
    
    Supported formats: 
    1. "DD Month YYYY HH:MM:SS" (French and English months) 
    2. "Month DD YYYY HH:MM:SS AM/PM" (English, 12-hour format) 
    3. Fallback: automatic pandas parsing 
    
    Parameters 
    ---------- 
    date_string : any 
        Input date value (string, NaN, etc.) 
    
    Returns 
    ------- 
    pandas.Timestamp or np.nan 
        Parsed datetime or NaN if parsing fails
    """

    # --- Handle missing or empty values ---
    date_str = str(date_string).strip()

    if pd.isna(date_string) or date_str == "":
        return np.nan

    # --- Month mapping (French + English) ---
    month_map = {
        'janv': '01', 'fevr': '02', 'févr': '02', 'mars': '03', 'avr': '04',
        'mai': '05', 'juin': '06', 'juil': '07', 'août': '08', 'sept': '09',
        'oct': '10', 'nov': '11', 'déc': '12', 'jan': '01', 'feb': '02',
        'mar': '03', 'apr': '04', 'may': '05', 'jun': '06', 'jul': '07',
        'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }

    try:
        parts = date_str.replace(",", "").split()
        
        # --- Format 1: "DD Month YYYY HH:MM:SS" --- .txt files
        if len(parts) == 4:
            day, month, year, time = parts
            month_key = month.replace('.', '').lower()

            if month_key in month_map:
                month_num = month_map[month_key]
                new_date = f"{day}.{month_num}.{year} {time}"
                
                return pd.to_datetime(new_date, format="%d.%m.%Y %H:%M:%S", errors='coerce')

        # --- Format 2: "Month DD YYYY HH:MM:SS AM/PM" --- .txt files
        if len(parts) == 5 and parts[-1] in ['AM', 'PM']:
            month, day, year, time, ampm = parts
            month_key = month.lower()

            if month_key in month_map:
                month_num = month_map[month_key]

                # Convert time to 24h format
                time_obj = datetime.strptime(time, "%I:%M:%S")
                if ampm == 'PM' and time_obj.hour != 12:
                    time_obj = time_obj.replace(hour=time_obj.hour + 12)
                elif ampm == 'AM' and time_obj.hour == 12:
                    time_obj = time_obj.replace(hour=0)

                new_date = f"{day}.{month_num}.{year} {time_obj.strftime('%H:%M:%S')}"

                return pd.to_datetime(new_date, format="%d.%m.%Y %H:%M:%S", errors='coerce')

        # --- Format 3: Fallback to pandas automatic parsing --- .xlsx and .csv files
        return pd.to_datetime(date_str, errors='coerce')

    except Exception:
        return np.nan
    

def extract_numeric(value):
    """Extract numeric value from string (in some .txt files, e.g. '0 per Logging Interval' → 0)."""
    
    if pd.isna(value) or value == '':
        return np.nan
    
    # Convert to string
    value = str(value)
    
    # Extract first number (integer or float)
    match = re.search(r"[-+]?\d*\.?\d+", value)
    
    if match:
        x = float(match.group())
        return int(x) if x.is_integer() else x
    
    return np.nan


def compute_stats(df):
    """Compute first date, last date, row count, and median time step (in minutes)."""
    df = df.dropna(subset=['DateTime'])

    if len(df) == 0:
        return None, None, 0, None

    df = df.sort_values('DateTime')

    first = df['DateTime'].iloc[0]
    last = df['DateTime'].iloc[-1]
    n = len(df)

    timestep = None
    if len(df) > 1:
        diffs = df['DateTime'].diff().dropna()
        if len(diffs) > 0:
            timestep = diffs.median().total_seconds() / 60

    return first, last, n, timestep

# -------------------------
# File processing 
# -------------------------

def process_file_txt(file_path):
    """
    Process a raw TXT pluviometer file and compute basic statistics.

    Reads the TXT file, skips metadata rows, parses the date column, 
    extracts numeric precipitation values, and returns the first and last 
    measurement dates, total number of measurements, and median time step 
    between measurements (in minutes).
    """
    try:
        df = pd.read_csv(file_path, sep='\t', skiprows=5, header=None,      # 5 first row in .txt are metadata
                         usecols=[0, 1, 2],
                         names=['Index', 'DateTime', 'Precipitation'])
        df['DateTime'] = df['DateTime'].apply(parse_date)                   # parse multiple date formats
        df['Precipitation'] = df['Precipitation'].apply(extract_numeric)    # parse numeric values from strings
        
        first, last, n, timestep = compute_stats(df)        
        duplicates = count_duplicates(df)       
        return first, last, n, timestep, duplicates                         # first/last date, row count, median time step, duplicates
    except Exception as e:
        print(f"Error TXT {file_path}: {e}")
        return None, None, 0, None, 0


def process_file_xlsx(file_path):
    """Process a single .xlsx file and return stats."""
    try:
        df = pd.read_excel(file_path, skiprows=6)                           # 6 first row in .xlsx are metadata
        df.columns = df.columns.str.strip()

        if 'DateTime' not in df.columns:
            df.rename(columns={df.columns[0]: 'DateTime'}, inplace=True)
        if 'Precipitation' not in df.columns and len(df.columns) > 1:
            df.rename(columns={df.columns[2]: 'Precipitation'}, inplace=True)

        # Filter out raw numeric Excel serial values (from UTC_convert artifacts)
        # Keep only actual datetime objects or strings, discard pure numbers
        if 'DateTime' in df.columns:
            df = df[df['DateTime'].apply(lambda x: not isinstance(x, (int, float)) or pd.isna(x))].copy()

        if pd.api.types.is_numeric_dtype(df['DateTime']):
            df['DateTime'] = pd.to_datetime(df['DateTime'], unit='d', origin='1899-12-30')
        else:
            df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
        if 'Precipitation' in df.columns:
            df['Precipitation'] = df['Precipitation'].apply(extract_numeric)

        first, last, n, timestep = compute_stats(df)        
        duplicates = count_duplicates(df)       
        return first, last, n, timestep, duplicates 
    except Exception as e:
        print(f"Error XLSX {file_path}: {e}")
        return None, None, 0, None, 0


def process_file_csv(file_path):
    """Process a single .csv file and return stats."""
    try:
        df = pd.read_csv(file_path)

        # make sure we have the expected columns (created_at for date and field5 for precipitation)
        if 'created_at' not in df.columns or 'field5' not in df.columns:                
            return None, None, 0, None

        # Keep only relevant columns and rename them (LoRa files have more information as humidity, temperature, etc. 
        # but we only keep date and precipitation for the stats)
        df = df[['created_at', 'field5']].copy()
        df.columns = ['DateTime', 'Precipitation']


        df['DateTime'] = pd.to_datetime(df['DateTime'], utc=True).dt.tz_localize(None)  
        df['Precipitation'] = df['Precipitation'].apply(extract_numeric)                # parse numeric values from strings

        first, last, n, timestep = compute_stats(df)        
        duplicates = count_duplicates(df)       
        return first, last, n, timestep, duplicates       # first/last date, row count, median time step, duplicates
    except Exception as e:
        print(f"Error CSV {file_path}: {e}")
        return None, None, 0, None, 0


def combine_excel_cell_datetime(date_value, time_value):
    """Combine Excel Date and Heure cell values into one datetime."""
    if isinstance(date_value, datetime):
        base_date = date_value.date()
        base_time = date_value.time()
    elif isinstance(date_value, (int, float)):
        full_datetime = datetime(1899, 12, 30) + timedelta(days=float(date_value))
        base_date = full_datetime.date()
        base_time = full_datetime.time()
    else:
        parsed_date = pd.to_datetime(date_value, errors="coerce", dayfirst=True)
        if pd.isna(parsed_date):
            return None
        base_date = parsed_date.date()
        base_time = parsed_date.time()

    if isinstance(time_value, datetime):
        base_time = time_value.time()
    elif isinstance(time_value, time):
        base_time = time_value
    elif isinstance(time_value, (int, float)):
        seconds = round((float(time_value) % 1) * 24 * 60 * 60)
        base_time = (datetime.min + timedelta(seconds=seconds)).time()
    elif time_value is not None:
        parsed_time = pd.to_datetime(time_value, errors="coerce")
        if not pd.isna(parsed_time):
            base_time = parsed_time.time()

    return datetime.combine(base_date, base_time)

# -------------------------
# Station processing (before and after)
# -------------------------

def process_station(directory, station_name):
    """Process all data files (.txt, .xlsx, .csv) in a station directory."""
    # List to store all summary data rows
    summary_rows = []
    all_txt_files = glob(os.path.join(directory, "*.txt"))
    all_excel_files = glob(os.path.join(directory, "*.xlsx"))
    all_csv_files = glob(os.path.join(directory, "*.csv"))

    # Process TXT files
    for file in all_txt_files:
        first, last, n, timestep, duplicates = process_file_txt(file)
        if first is not None:  # Only add if processing succeeded
            summary_rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".txt",
                "First measurement": first,
                "Last measurement": last,
                "Time step (min)": timestep,
                "Number of measurements": n,
                "Duplicates": duplicates,
                "First measurement AP": pd.NaT,
                "Last measurement AP": pd.NaT,
                "Number of measurements AP": np.nan,
                "Duplicates AP": np.nan,
                "Logbook comments": "",
                "Comments": ""
            })

    # Process XLSX files
    for file in all_excel_files:
        first, last, n, timestep, duplicates = process_file_xlsx(file)
        if first is not None:
            summary_rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".xlsx",
                "First measurement": first,
                "Last measurement": last,
                "Time step (min)": timestep,
                "Number of measurements": n,
                "Duplicates": duplicates,
                "First measurement AP": pd.NaT,
                "Last measurement AP": pd.NaT,
                "Number of measurements AP": np.nan,
                "Duplicates AP": np.nan,
                "Logbook comments": "",
                "Comments": ""
            })

    # Process CSV files
    for file in all_csv_files:
        first, last, n, timestep, duplicates = process_file_csv(file)
        if first is not None:
            summary_rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".csv",
                "First measurement": first,
                "Last measurement": last,
                "Time step (min)": timestep,
                "Number of measurements": n,
                "Duplicates": duplicates,
                "First measurement AP": pd.NaT,
                "Last measurement AP": pd.NaT,
                "Number of measurements AP": np.nan,
                "Duplicates AP": np.nan,
                "Logbook comments": "",
                "Comments": ""
            })
    return summary_rows

def process_station_after(directory, station_name):
    """Process all data files AFTER processing."""
    rows = []
    all_txt_files = glob(os.path.join(directory, "*.txt"))
    all_excel_files = glob(os.path.join(directory, "*.xlsx"))
    all_csv_files = glob(os.path.join(directory, "*.csv"))

    # Process TXT files
    for file in all_txt_files:
        first, last, n, _, duplicates = process_file_txt(file)
        if first is not None:
            rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".txt",
                "First measurement AP": first,
                "Last measurement AP": last,
                "Number of measurements AP": n,
                "Duplicates AP": duplicates
            })

    # Process XLSX files
    for file in all_excel_files:
        first, last, n, _, duplicates = process_file_xlsx(file)
        if first is not None:
            rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".xlsx",
                "First measurement AP": first,
                "Last measurement AP": last,
                "Number of measurements AP": n,
                "Duplicates AP": duplicates
            })

    # Process CSV files
    for file in all_csv_files:
        first, last, n, _, duplicates = process_file_csv(file)
        if first is not None:
            rows.append({
                "Station": station_name,
                "File name": os.path.basename(file),
                "File type": ".csv",
                "First measurement AP": first,
                "Last measurement AP": last,
                "Number of measurements AP": n,
                "Duplicates AP": duplicates
            })

    return rows

def count_duplicates(df):
    """Count duplicate DateTime values in a dataframe."""
    if 'DateTime' not in df.columns or len(df) == 0:
        return 0
    
    df = df.dropna(subset=['DateTime'])
    # Count duplicates (keep=False counts all occurrences of duplicates, not just subsequent ones)
    duplicates = df['DateTime'].duplicated(keep=False).sum()
    return duplicates


# -------------------------
# Add Logbook Comments and manual comments to summary report
# -------------------------

# Utility function to extract date from filename to join with logbook comments

def extract_date_from_filename(filename):
    """
    Extract date from filename in format AAAA_YYYYMMDD.ext or AAAA_BBBB_YYYYMMDD.ext.
    Returns a datetime.date object or None if parsing fails.
    """
    base = str(filename).strip()
    
    if not base:
        return None

    if "." in base:
        base = base.rsplit(".", 1)[0]   # remove extension --> rsplit : only 1 separation to remove the part after the last (first from right) "."  
                                        # to handle cases where filename contains multiple dots. [0] to get the part before the extension

    parts = base.split("_")

    for part in reversed(parts):
        if len(part) == 8 and part.isdigit():
            try:
                return datetime.strptime(part, "%Y%m%d").date()
            except ValueError:
                continue

    if len(parts) >= 2:
        date_str = parts[1]
        try:
            return datetime.strptime(date_str, "%Y%m%d").date()
        except ValueError:
            return None

    return None


def add_logbook_comments(available_df: pd.DataFrame, logbook_df: pd.DataFrame,                      
                         date_col: str = "Horodateur") -> pd.DataFrame:
    """
    Fill 'Logbook comments' column in available_df using pluvimates logbook. 
    Match on Station + date extracted from 'File name' vs logbook date column.

    available_df = Available_Data created by stations_processing.py with columns: 'Station', 'File name', etc.
    logbook_df = file from pluvimates logbook with columns: 'Station ID', 'Horodateur' (date), 'Comments', etc.
    """

    # Ensure column exists
    if "Logbook comments" not in available_df.columns:
        available_df["Logbook comments"] = ""
    else:
        if not pd.api.types.is_object_dtype(available_df["Logbook comments"]):
            available_df["Logbook comments"] = available_df["Logbook comments"].astype(object)

    if date_col not in logbook_df.columns:
        raise KeyError(f"The column '{date_col}' was not found in logbook dataframe")

    # Precompute logbook dates
    log_dates = pd.to_datetime(logbook_df[date_col], errors="coerce").dt.date

    # Iterate rows
    for index, row in available_df.iterrows():
        station = row.get("Station")
        filename = row.get("File name")
        date = extract_date_from_filename(filename)

        if pd.isna(station) or date is None:
            continue

        matching_log = logbook_df[
            (logbook_df["Station ID"] == station) & (log_dates == date)
        ]

        if len(matching_log) > 0:
            available_df.at[index, "Logbook comments"] = str(matching_log["Comments"].iloc[0])

    return available_df


def add_file_comments(available_df: pd.DataFrame, comments_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge 'Comments' from Comments.xlsx into available_df based on Station + File name.
    comments_df must contain columns: 'Station', 'File name', 'Comments'.
    """
    # Ensure Comments column exists
    if "Comments" not in available_df.columns:
        available_df["Comments"] = ""

    # Merge with suffixes to avoid overwriting existing Comments if they exist
    merged = available_df.merge(
        comments_df[["Station", "File name", "Comments"]],
        on=["Station", "File name"],
        how="left",
        suffixes=("", "_from_comments")
    )

    # combine Comments columns if both exist (keep original Comments if not empty, otherwise take from comments_df)
    if "Comments_from_comments" in merged.columns:
        merged["Comments"] = merged["Comments"].replace("", pd.NA)
        merged["Comments_from_comments"] = merged["Comments_from_comments"].replace("", pd.NA)
        merged["Comments"] = merged["Comments"].combine_first(merged["Comments_from_comments"])
        merged = merged.drop(columns=["Comments_from_comments"])

    return merged
