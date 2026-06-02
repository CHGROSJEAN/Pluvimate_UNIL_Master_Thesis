# Charlotte Grosjean - 22.04.2026 - UNIL Master Thesis
"""
OVERLAP ANALYSIS UTILS
This module contains utility functions for identifying, processing, and comparing overlapping periods between different data sources for pluviometer stations. It includes functions to:
- Load and clean station data files
- Identify overlapping time periods between sources
- Save overlapping data for further analysis
- Align overlapping data series using various methods (exact, tolerance-based, interpolation)
- Compute comparison metrics (mean difference, RMSE, correlation)
- Generate visualizations of aligned series
"""

import os
import pandas as pd
import numpy as np
from glob import glob
from collections import defaultdict
import re
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import plotly.graph_objects as go
import plotly.io as pio
from scipy import signal
from A_data_processing_utils import *
import E_plot_config as cfg # centralized plot styling configuration

# -------------------------
# Overlap identification functions
# -------------------------

# helper functions for identify_overlaps.py
def clean_and_prepare(df, file_name):
    """
    Clean a raw dataframe:
    - Extract numeric precipitation
    - Remove invalid rows (NaNs)
    """
    df['Precipitation'] = df['Precipitation'].apply(extract_numeric) # extract_numeric function defined in data_processing_utils.py to extract numeric values from precipitation column
    df['Precipitation'] = df['Precipitation']*0.01  # Convert drops to mm (1 drop = 0.01 mm)
    valid = df.dropna(subset=['DateTime', 'Precipitation']).copy()
    valid['Source_File'] = file_name
    return valid[['Source_File', 'DateTime', 'Precipitation']]

def load_station_files(directory):
    """
    Load all files for a station and return list of (source, df) tuples.
    """
    station_name = os.path.basename(directory)
    txt_files = glob(os.path.join(directory, "*.txt"))
    xlsx_files = glob(os.path.join(directory, "*.xlsx"))
    csv_files = glob(os.path.join(directory, "*.csv"))

    sources = []

    # --- TXT ---
    for file_path in txt_files:
        try:
            df = pd.read_csv(file_path, sep='\t', skiprows=5, header=None,
                             usecols=[0, 1, 2],
                             names=['Index', 'DateTime', 'Precipitation'])
            print(f"  TXT {os.path.basename(file_path)}: loaded {len(df)} raw rows")
            df['DateTime'] = df['DateTime'].apply(parse_date) # parse_date function defined in data_processing_utils.py to parse date strings into datetime objects
            valid = clean_and_prepare(df, os.path.basename(file_path)) # function defined above to clean and prepare data
            print(f"    -> After cleaning: {len(valid)} valid rows")
            if not valid.empty:
                sources.append((os.path.basename(file_path), valid))
            else:
                print(f"    -> WARNING: No valid data after cleaning!")
        except Exception as e:
            print(f"Error TXT {file_path}: {e}")

    # --- XLSX ---
    for file_path in xlsx_files:
        if os.path.basename(file_path).startswith('~$'):
            continue
        try:
            df = pd.read_excel(file_path, skiprows=6)
            print(f"  XLSX {os.path.basename(file_path)}: loaded {len(df)} raw rows")
            df.columns = df.columns.str.strip()
            date_col = next((c for c in df.columns if 'date' in c.lower() or 'time' in c.lower()), df.columns[0])
            precip_col = next((c for c in df.columns if 'precip' in c.lower() or 'pulse' in c.lower()), None)
            print(f"    -> Detected columns: date='{date_col}', precip='{precip_col}'")
            if precip_col:
                df['DateTime'] = df[date_col].apply(parse_date) # parse_date function defined in data_processing_utils.py to parse date strings into datetime objects
                df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
                df['Precipitation'] = df[precip_col]
                valid = clean_and_prepare(df, os.path.basename(file_path)) # function defined above to clean and prepare data
                print(f"    -> After cleaning: {len(valid)} valid rows")
                if not valid.empty:
                    sources.append((os.path.basename(file_path), valid))
                else:
                    print(f"    -> WARNING: No valid data after cleaning!")
            else:
                print(f"    -> WARNING: No precipitation column found!")
        except Exception as e:
            print(f"Error XLSX {file_path}: {e}")

    # --- CSV ---
    for file_path in csv_files:
        try:
            df = pd.read_csv(file_path)
            print(f"  CSV {os.path.basename(file_path)}: loaded {len(df)} raw rows")
            if 'created_at' in df.columns and 'field5' in df.columns:
                df = df[['created_at', 'field5']].copy()
                df.columns = ['DateTime', 'Precipitation']
                df['DateTime'] = pd.to_datetime(df['DateTime'], utc=True).dt.tz_localize(None)
                valid = clean_and_prepare(df, os.path.basename(file_path)) # function defined above to clean and prepare data
                print(f"    -> After cleaning: {len(valid)} valid rows")
                if not valid.empty:
                    sources.append((os.path.basename(file_path), valid))
                else:
                    print(f"    -> WARNING: No valid data after cleaning!")
            else:
                print(f"    -> WARNING: Missing required columns (created_at, field5)")
        except Exception as e:
            print(f"Error CSV {file_path}: {e}")

    return station_name, sources

# find overlaps and save data for each station
def find_overlaps(sources):
    """
    Find overlapping periods between sources.
    Returns list of (source1, source2, start_overlap, end_overlap) tuples.
    Only returns overlaps where BOTH sources have actual data in the overlap period.
    """
    overlaps = []
    for i in range(len(sources)):
        for j in range(i+1, len(sources)):
            source1, df1 = sources[i]
            source2, df2 = sources[j]

            start1 = df1['DateTime'].min()
            end1 = df1['DateTime'].max()
            start2 = df2['DateTime'].min()
            end2 = df2['DateTime'].max()

            print(f"  Checking overlap between {source1} and {source2}:")
            print(f"    {source1}: {len(df1)} points from {start1} to {end1}")
            print(f"    {source2}: {len(df2)} points from {start2} to {end2}")

            start_overlap = max(start1, start2)
            end_overlap = min(end1, end2)

            if start_overlap < end_overlap:
                # Check if both sources actually have data in the overlap period
                mask1 = (df1['DateTime'] >= start_overlap) & (df1['DateTime'] <= end_overlap)
                mask2 = (df2['DateTime'] >= start_overlap) & (df2['DateTime'] <= end_overlap)

                actual_data1 = df1[mask1]
                actual_data2 = df2[mask2]

                print(f"    Potential overlap period: {start_overlap} to {end_overlap}")
                print(f"    {source1}: {len(actual_data1)} actual points in overlap")
                print(f"    {source2}: {len(actual_data2)} actual points in overlap")

                # Only consider it a valid overlap if both sources have meaningful data
                min_points_threshold = 1  # Require at least 1 point from each source

                if len(actual_data1) >= min_points_threshold and len(actual_data2) >= min_points_threshold:
                    overlaps.append((source1, source2, start_overlap, end_overlap))
                    print(f"    -> VALID OVERLAP FOUND: {start_overlap} to {end_overlap}")
                else:
                    print(f"    -> INSUFFICIENT DATA: {source1} has {len(actual_data1)} points, {source2} has {len(actual_data2)} points (need ≥{min_points_threshold} each)")
            else:
                print(f"    -> No overlap")

    return overlaps

def save_overlap_data(station_name, source1, df1, source2, df2, start_overlap, end_overlap, output_dir):
    """
    Save overlapping data for each source.
    """
    print(f"  Saving overlap data for period: {start_overlap} to {end_overlap}")

    # Filter data for overlap period
    mask1 = (df1['DateTime'] >= start_overlap) & (df1['DateTime'] <= end_overlap)
    overlap_df1 = df1[mask1].copy()

    mask2 = (df2['DateTime'] >= start_overlap) & (df2['DateTime'] <= end_overlap)
    overlap_df2 = df2[mask2].copy()

    print(f"    {source1}: {len(overlap_df1)} points in overlap period")
    print(f"    {source2}: {len(overlap_df2)} points in overlap period")

    if overlap_df1.empty:
        print(f"    WARNING: {source1} has no data in overlap period!")
    if overlap_df2.empty:
        print(f"    WARNING: {source2} has no data in overlap period!")

    # Format dates for filename
    start_str = start_overlap.strftime('%Y%m%d')
    end_str = end_overlap.strftime('%Y%m%d')

    # Save files
    file1 = f"{source1.replace('.','_')}_overlap_{start_str}_{end_str}.csv"
    file2 = f"{source2.replace('.','_')}_overlap_{start_str}_{end_str}.csv"

    overlap_df1.to_csv(os.path.join(output_dir, file1), index=False)
    overlap_df2.to_csv(os.path.join(output_dir, file2), index=False)

    print(f"Saved overlap files: {file1} and {file2}")


# -------------------------
# Overlap comparison functions
# -------------------------

# Pattern used to group overlap files by overlap period
FILE_PATTERN = re.compile(r"_overlap_([0-9_]+)\.csv$")

# helper functions for compare_overlaps.py

def load_and_prepare_df(path):
    """Load CSV and prepare for comparison."""
    df = pd.read_csv(path)
    # Convert string dates to pandas Timestamp
    df['DateTime'] = pd.to_datetime(df['DateTime'])
    # Sort in time order and drop rows with missing core values
    df = df.sort_values('DateTime').dropna(subset=['DateTime', 'Precipitation'])
    # Ensure precipitation values are numeric
    df['Precipitation'] = pd.to_numeric(df['Precipitation'], errors='coerce')
    df = df.dropna(subset=['Precipitation']) # Drop rows where precipitation could not be converted to numeric
    return df

def get_overlap_key(filename):
    """Extract overlap period key from filename."""
    match = FILE_PATTERN.search(filename)
    return match.group(1) if match else None

# Align two dataframes on DateTime 


def align_data(df1, df2, method='exact', tolerance_minutes=3):
    """
    Align two dataframes temporally.

    Methods:
    - 'exact': only exact timestamp matches
    - 'tolerance': timestamps within tolerance_minutes of each other
    - 'nearest': nearest neighbor interpolation
    - 'linear': linear interpolation
    """
    if method == 'exact':
        # Merge on exact DateTime only.
        # This is the strictest alignment, used when timestamps are synchronized.
        merged = pd.merge(df1[['DateTime', 'Precipitation']],
                         df2[['DateTime', 'Precipitation']],
                         on='DateTime', suffixes=('_1', '_2'))
    elif method == 'tolerance':
        # Align points that are close in time, within tolerance_minutes.
        # This helps compare series when measurements are not exactly synchronized.
        merged_rows = []
        df2_sorted = df2.sort_values('DateTime')

        for _, row1 in df1.iterrows():
            dt1 = row1['DateTime']
            # Find closest timestamp in df2 within tolerance
            mask = (df2_sorted['DateTime'] >= dt1 - pd.Timedelta(minutes=tolerance_minutes)) & \
                   (df2_sorted['DateTime'] <= dt1 + pd.Timedelta(minutes=tolerance_minutes))

            closest_matches = df2_sorted[mask]
            if not closest_matches.empty:
                # Take the closest timestamp within tolerance
                closest_idx = (closest_matches['DateTime'] - dt1).abs().idxmin()
                row2 = closest_matches.loc[closest_idx]
                merged_rows.append({
                    'DateTime': dt1,  # Use df1 timestamp as reference
                    'Precipitation_1': row1['Precipitation'],
                    'Precipitation_2': row2['Precipitation']
                })

        merged = pd.DataFrame(merged_rows)
    else:
        # Interpolate both series on a common 3-minute grid.
        # This is useful when neither exact nor tolerance-based matching yields enough points.
        df1_idx = df1[['DateTime', 'Precipitation']].set_index('DateTime')
        df2_idx = df2[['DateTime', 'Precipitation']].set_index('DateTime')

        # Create common time range for overlap period
        start = max(df1_idx.index.min(), df2_idx.index.min())
        end = min(df1_idx.index.max(), df2_idx.index.max())
        common_times = pd.date_range(start=start, end=end, freq='3min')  # Assuming 3min resolution

        # Reindex and interpolate missing values in time
        df1_interp = df1_idx.reindex(common_times).interpolate(method='nearest')
        df2_interp = df2_idx.reindex(common_times).interpolate(method='nearest')

        merged = pd.DataFrame({
            'DateTime': common_times,
            'Precipitation_1': df1_interp['Precipitation'],
            'Precipitation_2': df2_interp['Precipitation']
        }).dropna()

    return merged

def save_aligned_data(aligned_df, station_name, overlap_key, file_paths, output_dir, alignment_method, metrics, cross_corr_info=None):
    """
    Save aligned data series to a CSV file with metadata header.
    This creates a single file with both series temporally aligned.
    """
    from datetime import datetime
    
    # Create 'aligned_data' subfolder within the station folder
    aligned_dir = os.path.join(output_dir, 'aligned_data')
    os.makedirs(aligned_dir, exist_ok=True)
    
    # Add source file names as columns for reference
    source1_name = os.path.basename(file_paths[0]).replace('.csv', '')
    source2_name = os.path.basename(file_paths[1]).replace('.csv', '')
    
    # Rename columns to indicate source
    output_df = aligned_df.copy()
    output_df = output_df.rename(columns={
        'Precipitation_1': 'Precip_SourceFile1',
        'Precipitation_2': 'Precip_SourceFile2'
    })
    
    # Create filename
    output_filename = f"{station_name}_overlap_{overlap_key}_aligned.csv"
    output_path = os.path.join(aligned_dir, output_filename)
    
    # Write metadata header and data to CSV using text mode
    with open(output_path, 'w') as f:
        # Write metadata header
        f.write("# ALIGNED OVERLAP DATA - METADATA\n")
        f.write(f"# Station: {station_name}\n")
        f.write(f"# Overlap Period: {overlap_key}\n")
        f.write(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"# Source File 1: {source1_name}.csv\n")
        f.write(f"# Source File 2: {source2_name}.csv\n")
        f.write(f"# Alignment Method: {alignment_method}\n")
        f.write(f"# Number of Aligned Points: {len(output_df)}\n")
        f.write(f"# Mean Difference (Precip1 - Precip2): {metrics.get('mean_diff', 'N/A')}\n")
        f.write(f"# RMSE (Root Mean Square Error): {metrics.get('rmse', 'N/A')}\n")
        f.write(f"# MAE (Mean Absolute Error): {metrics.get('mae', 'N/A')}\n")
        f.write(f"# Correlation: {metrics.get('correlation', 'N/A')}\n")
        f.write(f"# Max Difference: {metrics.get('max_diff', 'N/A')}\n")
        f.write(f"# Min Difference: {metrics.get('min_diff', 'N/A')}\n")
        if cross_corr_info:
            best_lag, best_corr = cross_corr_info
            f.write(f"# Cross-Correlation Best Lag: {best_lag:.2f} hours\n")
            f.write(f"# Cross-Correlation Max Value: {best_corr:.3f}\n")
        f.write("#\n")
    
    # Append data to CSV (with header)
    output_df.to_csv(output_path, mode='a', index=False)
    return output_path


# Visualization of aligned series
def plot_aligned_series_interactive(aligned_df, station_name, overlap_key, output_dir, source_1_name='Source 1', source_2_name='Source 2'):
    """Create an interactive aligned time-series plot and save it as HTML."""
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=aligned_df['DateTime'],
        y=aligned_df['Precipitation_1'],
        mode='lines',
        name=source_1_name,
        marker=dict(size=4),
        line=dict(width=1)
    ))
    fig.add_trace(go.Scatter(
        x=aligned_df['DateTime'],
        y=aligned_df['Precipitation_2'],
        mode='lines',
        name=source_2_name,
        marker=dict(size=4),
        line=dict(width=1, dash='dash')
    ))

    fig.update_layout(
        title=dict(
        text=f'{station_name} - Overlap {overlap_key}',
        x=0.5
        ),
        xaxis_title='DateTime',
        yaxis_title='Precipitation (mm)/ 3 min',
        hovermode='x unified',
        template='plotly_white'
    )

    plot_filename = f"{station_name}_overlap_{overlap_key}_aligned_interactive.html"
    plot_path = os.path.join(output_dir, plot_filename)
    pio.write_html(fig, plot_path, include_plotlyjs='cdn', full_html=True)

    return plot_path

# compute comparison metrics

def compute_metrics(df):
    """Compute comparison metrics between two aligned series."""
    if df.empty:
        return {
            'n_points': 0,
            'mean_diff': np.nan,
            'std_diff': np.nan,
            'rmse': np.nan,
            'mae': np.nan,
            'correlation': np.nan,
            'max_diff': np.nan,
            'min_diff': np.nan
        }

    diff = df['Precipitation_1'] - df['Precipitation_2']

    return {
        'n_points': len(df),
        'mean_diff': diff.mean(),
        'std_diff': diff.std(),
        'rmse': np.sqrt((diff ** 2).mean()),
        'mae': diff.abs().mean(),
        'correlation': df['Precipitation_1'].corr(df['Precipitation_2']),
        'max_diff': diff.max(),
        'min_diff': diff.min()
    }

# Compute cross-correlation 

def compute_cross_correlation(df, max_lag_hours=24):
    """
    Compute cross-correlation between two precipitation series.
    Returns correlation coefficients for different time lags.
    """
    if df.empty or len(df) < 10:
        return None, None, None

    # Convert max_lag_hours to number of points (assuming 3-minute intervals)
    max_lag_points = int(max_lag_hours * 60 / 3)  # 3-minute intervals

    # Extract precipitation series
    series1 = df['Precipitation_1'].values
    series2 = df['Precipitation_2'].values

    # Compute cross-correlation
    corr = signal.correlate(series1, series2, mode='full', method='auto')
    lags = signal.correlation_lags(len(series1), len(series2), mode='full')

    # Normalize by the geometric mean of autocorrelations at lag 0
    norm_factor = np.sqrt(np.sum(series1**2) * np.sum(series2**2))
    if norm_factor > 0:
        corr = corr / norm_factor

    # Keep only lags within the specified range
    valid_mask = np.abs(lags) <= max_lag_points
    lags = lags[valid_mask]
    corr = corr[valid_mask]

    # Convert lags to hours
    lags_hours = lags * 3 / 60  # 3 minutes to hours

    # Find the lag with maximum correlation
    if len(corr) > 0:
        max_corr_idx = np.argmax(np.abs(corr))
        best_lag_hours = lags_hours[max_corr_idx]
        best_corr = corr[max_corr_idx]
    else:
        best_lag_hours = 0
        best_corr = 0

    return lags_hours, corr, (best_lag_hours, best_corr)

# Visualization of cross-correlation results
def plot_cross_correlation(lags_hours, corr, best_lag_info, station_name, overlap_key, source_1_name, source_2_name, output_dir):
    """
    Create a plot of cross-correlation vs time lag.
    """
    best_lag_hours, best_corr = best_lag_info

    plt.switch_backend('Agg')
    plt.figure(figsize=(7, 7))

    # Plot cross-correlation
    plt.plot(lags_hours, corr, 'b-', linewidth=2, alpha=0.8, label='Cross-correlation')
    plt.axhline(y=0, color='k', linestyle='--', alpha=0.5)

    # Highlight the maximum correlation point
    plt.plot(best_lag_hours, best_corr, 'ro', markersize=8, label=f'Best lag: {best_lag_hours:.2f}h')

    # Add vertical line at lag 0
    plt.axvline(x=0, color='gray', linestyle=':', alpha=0.7)

    plt.xlabel('Time Lag (hours)', fontsize=cfg.LABEL_FONTSIZE)
    plt.ylabel('Cross-Correlation', fontsize=cfg.LABEL_FONTSIZE)
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    # Save plot as PNG and PDF. PDF is preferable for report figures.
    plot_base = f"{station_name}_overlap_{overlap_key}_cross_correlation"
    png_path = os.path.join(output_dir, f"{plot_base}.png")
    pdf_path = os.path.join(output_dir, f"{plot_base}.pdf")
    plt.savefig(png_path, dpi=cfg.DPI, bbox_inches='tight')
    plt.savefig(pdf_path, bbox_inches='tight')
    plt.close()

    return png_path

def plot_cross_correlation_interactive(lags_hours, corr, best_lag_info, station_name, overlap_key, source_1_name, source_2_name, output_dir):
    """Create an interactive cross-correlation plot and save it as HTML."""
    best_lag_hours, best_corr = best_lag_info

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=lags_hours,
        y=corr,
        mode='lines',
        name='Cross-correlation',
        line=dict(color='blue')
    ))
    fig.add_trace(go.Scatter(
        x=[best_lag_hours],
        y=[best_corr],
        mode='markers',
        name=f'Best lag: {best_lag_hours:.2f}h',
        marker=dict(color='red', size=10)
    ))
    fig.add_shape(type='line', x0=0, x1=0, y0=min(corr), y1=max(corr), line=dict(color='gray', dash='dash'))

    fig.update_layout(
        title=dict(
        text=f'Cross-correlation - (Overlap: {overlap_key})',
        x=0.5
        ),     
        xaxis_title='Time Lag (hours)',
        yaxis_title='Cross-Correlation',
        hovermode='x unified',
        template='plotly_white'
    )

    plot_filename = f"{station_name}_overlap_{overlap_key}_cross_correlation.html"
    plot_path = os.path.join(output_dir, plot_filename)
    pio.write_html(fig, plot_path, include_plotlyjs='cdn', full_html=True)

    return plot_path


# -------------------------
# Overlap visualization functions
# -------------------------

# helper functions for visualize_overlaps.py
def load_overlap_csv(path):
    df = pd.read_csv(path)
    if 'DateTime' not in df.columns:
        raise ValueError(f"The file {path} does not contain a 'DateTime' column.")
    df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
    if df['DateTime'].isna().all():
        raise ValueError(f"The file {path} contains only invalid DateTime values.")
    return df.sort_values('DateTime')

def get_label(df, file_path):
    if 'Source_File' in df.columns and not df['Source_File'].isna().all():
        return str(df['Source_File'].iloc[0])
    return os.path.basename(file_path)

# Single file visualizatin 
PLOT_FILENAME_SUFFIX = "_overlap_plot.png"
OVERLAP_LINE_COLORS = ["#1f77b4", "#d62728", "#2ca02c", "#9467bd"]
OVERLAP_LINE_ALPHAS = [1.0, 0.75, 0.75, 0.75]
OVERLAP_LINE_WIDTH = 1.4

def visualize_single_overlap(file1, file2, save_plot=True, show_plot=True):
    df1 = load_overlap_csv(file1)
    df2 = load_overlap_csv(file2)

    source1 = get_label(df1, file1)
    source2 = get_label(df2, file2)

    plt.figure(figsize=(14, 6))
    plt.plot(
        df1['DateTime'],
        df1['Precipitation'],
        label=f'Source: {source1}',
        color=OVERLAP_LINE_COLORS[0],
        linewidth=OVERLAP_LINE_WIDTH,
        alpha=OVERLAP_LINE_ALPHAS[0],
    )
    plt.plot(
        df2['DateTime'],
        df2['Precipitation'],
        label=f'Source: {source2}',
        color=OVERLAP_LINE_COLORS[1],
        linewidth=OVERLAP_LINE_WIDTH,
        alpha=OVERLAP_LINE_ALPHAS[1],
    )

    plt.xlabel('Time', fontsize=cfg.LABEL_FONTSIZE)
    plt.ylabel('Precipitation [mm]', fontsize=cfg.LABEL_FONTSIZE)
    plt.legend(fontsize=cfg.LEGEND_FONTSIZE)
    plt.grid(True, alpha=cfg.GRID_ALPHA)
    plt.tick_params(axis='both', labelsize=cfg.TICK_FONTSIZE)
    plt.xticks(rotation=45)
    plt.tight_layout()

    if save_plot:
        output_dir = os.path.join(os.path.dirname(file1), 'plots')
        os.makedirs(output_dir, exist_ok=True)
        output_filename = f"single_overlap_{os.path.basename(file1).replace('.csv', '')}_{os.path.basename(file2).replace('.csv', '')}{PLOT_FILENAME_SUFFIX}"
        output_path = os.path.join(output_dir, output_filename)
        plt.savefig(output_path, dpi=cfg.DPI, bbox_inches='tight')
        print(f"Plot saved to: {output_path}")

    if show_plot:
        plt.show()

    plt.close()

def visualize_single_overlap_scatter_comparison(file1, file2, save_plot=True, show_plot=True, aligned_file_path=None):
    """Create a comparative scatter plot: Series 1 vs Series 2 (each on an axis)."""
    try:
        # If an aligned file is provided and exists, use it
        if aligned_file_path and os.path.exists(aligned_file_path):
            df = pd.read_csv(aligned_file_path, comment='#')
            
            # Look for aligned precipitation columns
            col_1 = None
            col_2 = None
            for col in df.columns:
                if 'Precip_SourceFile1' in col or 'Precipitation_1' in col:
                    col_1 = col
                elif 'Precip_SourceFile2' in col or 'Precipitation_2' in col:
                    col_2 = col
            
            if col_1 is None or col_2 is None:
                print(f"Warning: Could not find precipitation columns in aligned file, using original data")
                # Fall back to direct comparison
                df1 = load_overlap_csv(file1)
                df2 = load_overlap_csv(file2)
                # Align on exact DateTime
                merged = pd.merge(df1[['DateTime', 'Precipitation']], df2[['DateTime', 'Precipitation']], 
                                 on='DateTime', suffixes=('_1', '_2'))
                col_1 = 'Precipitation_1'
                col_2 = 'Precipitation_2'
                df = merged
        else:
            # Align the two files on DateTime
            df1 = load_overlap_csv(file1)
            df2 = load_overlap_csv(file2)
            merged = pd.merge(df1[['DateTime', 'Precipitation']], df2[['DateTime', 'Precipitation']], 
                             on='DateTime', suffixes=('_1', '_2'))
            col_1 = 'Precipitation_1'
            col_2 = 'Precipitation_2'
            df = merged
        
        if df.empty:
            print("No common data points found for comparison")
            return
        
        source1 = get_label(df1, file1) if 'df1' in locals() else 'Source File 1'
        source2 = get_label(df2, file2) if 'df2' in locals() else 'Source File 2'
        
        plt.figure(figsize=(10, 10))
        
        # Create scatter plot: Source File 1 vs Source File 2
        plt.scatter(df[col_1], df[col_2], s=30, alpha=0.6, color='steelblue', edgecolors='black', linewidth=0.5)
        
        # Add reference line (perfect agreement: y=x)
        min_val = min(df[col_1].min(), df[col_2].min())
        max_val = max(df[col_1].max(), df[col_2].max())
        plt.plot([min_val, max_val], [min_val, max_val], 'r--', linewidth=2, label='Perfect agreement', alpha=0.7)
        
        plt.xlabel(f'{source1} - Precipitation (mm)', fontsize=cfg.LABEL_FONTSIZE)
        plt.ylabel(f'{source2} - Precipitation (mm)', fontsize=cfg.LABEL_FONTSIZE)
        plt.title(f'Comparative Scatter Plot\n{source1} vs {source2}', fontsize=cfg.TITLE_FONTSIZE)
        plt.legend(fontsize=cfg.LEGEND_FONTSIZE)
        plt.grid(True, alpha=cfg.GRID_ALPHA)
        plt.tick_params(axis='both', labelsize=cfg.TICK_FONTSIZE)
        plt.axis('equal')
        plt.tight_layout()
        
        if save_plot:
            output_dir = os.path.join(os.path.dirname(file1), 'plots')
            os.makedirs(output_dir, exist_ok=True)
            output_filename = f"single_overlap_comparison_{os.path.basename(file1).replace('.csv', '')}_{os.path.basename(file2).replace('.csv', '')}{PLOT_FILENAME_SUFFIX}"
            output_path = os.path.join(output_dir, output_filename)
            plt.savefig(output_path, dpi=cfg.DPI, bbox_inches='tight')
            print(f"Comparative scatter plot saved to: {output_path}")
        
        if show_plot:
            plt.show()
        
        plt.close()
        
    except Exception as e:
        print(f"Error creating comparative scatter plot: {e}")

# All files visualization

def build_groups(station_dir):
    groups = defaultdict(list)
    for file_path in glob(os.path.join(station_dir, "*.csv")):
        key = get_overlap_key(os.path.basename(file_path))
        if key:
            groups[key].append(file_path)
    return groups

def plot_group(file_paths, output_path, station_name, overlap_key):
    plt.switch_backend('Agg')
    plt.figure(figsize=(14, 6))

    sorted_files = sorted(file_paths)
    for i, file_path in enumerate(sorted_files):
        df = load_overlap_csv(file_path)
        label = get_label(df, file_path)
        plt.plot(
            df['DateTime'],
            df['Precipitation'],
            color=OVERLAP_LINE_COLORS[i % len(OVERLAP_LINE_COLORS)],
            linewidth=OVERLAP_LINE_WIDTH,
            alpha=OVERLAP_LINE_ALPHAS[i % len(OVERLAP_LINE_ALPHAS)],
            label=label,
        )

    plt.xlabel('Time', fontsize=cfg.LABEL_FONTSIZE)
    plt.ylabel('Precipitation [mm]', fontsize=cfg.LABEL_FONTSIZE)
    plt.legend(fontsize=cfg.LEGEND_FONTSIZE)
    plt.grid(True, alpha=cfg.GRID_ALPHA)
    plt.tick_params(axis='both', labelsize=cfg.TICK_FONTSIZE)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(output_path, dpi=cfg.DPI, bbox_inches='tight')
    pdf_path = os.path.splitext(output_path)[0] + ".pdf"
    plt.savefig(pdf_path, bbox_inches='tight')
    plt.close()

def plot_group_report_zoom(file_paths, output_dir, station_name, overlap_key, start_time, end_time, formats=('pdf', 'png'), dpi=600):
    """Create a publication-quality static zoom plot for a selected time window."""
    plt.switch_backend('Agg')
    start_time = pd.to_datetime(start_time)
    end_time = pd.to_datetime(end_time)

    fig, ax = plt.subplots(figsize=(14, 7))
    plotted_any_data = False
    sorted_files = sorted(file_paths)
    for i, file_path in enumerate(sorted_files):
        df = load_overlap_csv(file_path)
        mask = (df['DateTime'] >= start_time) & (df['DateTime'] <= end_time)
        df_zoom = df.loc[mask]
        if df_zoom.empty:
            continue

        plotted_any_data = True
        label = get_label(df_zoom, file_path)
        ax.plot(
            df_zoom['DateTime'],
            df_zoom['Precipitation'],
            color=OVERLAP_LINE_COLORS[i % len(OVERLAP_LINE_COLORS)],
            linewidth=OVERLAP_LINE_WIDTH,
            alpha=OVERLAP_LINE_ALPHAS[i % len(OVERLAP_LINE_ALPHAS)],
            label=label
        )

    if not plotted_any_data:
        plt.close(fig)
        raise ValueError(f"No data in selected window: {start_time} to {end_time}")

    ax.set_xlabel('Time', fontsize=cfg.LABEL_FONTSIZE)
    ax.set_ylabel('Precipitation [mm]', fontsize=cfg.LABEL_FONTSIZE)
    ax.legend(fontsize=cfg.LEGEND_FONTSIZE, frameon=True, framealpha=0.95)
    ax.grid(True, alpha=0.3)
    ax.tick_params(axis='both', labelsize=cfg.TICK_FONTSIZE)
    ax.set_xlim(start_time, end_time)
    locator = mdates.AutoDateLocator(minticks=7, maxticks=12)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M'))
    fig.autofmt_xdate(rotation=35, ha='right')
    fig.tight_layout()

    os.makedirs(output_dir, exist_ok=True)
    start_label = start_time.strftime('%Y%m%d_%H%M')
    end_label = end_time.strftime('%Y%m%d_%H%M')
    output_paths = []

    for output_format in formats:
        output_format = output_format.lower().lstrip('.')
        output_filename = f"{station_name}_overlap_{overlap_key}_zoom_{start_label}_{end_label}.{output_format}"
        output_path = os.path.join(output_dir, output_filename)
        if output_format in ('png', 'jpg', 'jpeg', 'tif', 'tiff'):
            fig.savefig(output_path, dpi=dpi, bbox_inches='tight')
        else:
            fig.savefig(output_path, bbox_inches='tight')
        output_paths.append(output_path)

    plt.close(fig)
    return output_paths

def plot_group_scatter_comparison(aligned_file_path, output_path, station_name, overlap_key):
    """Create a comparative scatter plot: Series 1 vs Series 2."""
    try:
        df = pd.read_csv(aligned_file_path, comment='#')
        
        # Look for aligned precipitation columns
        col_1 = None
        col_2 = None
        for col in df.columns:
            if 'Precip_SourceFile1' in col or 'Precipitation_1' in col:
                col_1 = col
            elif 'Precip_SourceFile2' in col or 'Precipitation_2' in col:
                col_2 = col
        
        if col_1 is None or col_2 is None:
            print(f"    -> Warning: Could not find precipitation columns in {aligned_file_path}")
            return
        
        plt.figure(figsize=(10, 10))
        
        # Create scatter plot: Source File 1 vs Source File 2
        plt.scatter(df[col_1], df[col_2], s=30, alpha=0.6, color='steelblue', edgecolors='black', linewidth=0.5)
        
        # Add reference line (perfect agreement: y=x)
        min_val = min(df[col_1].min(), df[col_2].min())
        max_val = max(df[col_1].max(), df[col_2].max())
        plt.plot([min_val, max_val], [min_val, max_val], 'r--', linewidth=2, label='Perfect agreement', alpha=0.7)
        
        plt.xlabel('Source File 1 - Precipitation (mm)', fontsize=cfg.LABEL_FONTSIZE)
        plt.ylabel('Source File 2 - Precipitation (mm)', fontsize=cfg.LABEL_FONTSIZE)
        plt.title(f"{station_name} - Overlap {overlap_key}", fontsize=cfg.TITLE_FONTSIZE)
        plt.legend(fontsize=cfg.LEGEND_FONTSIZE)
        plt.grid(True, alpha=cfg.GRID_ALPHA)
        plt.tick_params(axis='both', labelsize=cfg.TICK_FONTSIZE)
        plt.axis('equal')
        plt.tight_layout()
        plt.savefig(output_path, dpi=cfg.DPI, bbox_inches='tight')
        plt.close()
        
    except Exception as e:
        print(f"    -> Error creating comparative scatter plot: {e}")
