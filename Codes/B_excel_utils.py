# Charlotte Grosjean - 25.03.2026 - UNIL Master Thesis

# This script contains utility functions for formatting Excel reports, such as applying styles, 
# adjusting column widths, and highlighting specific conditions in the data. 
# It is used in the Stations_Processing.py script to enhance 
# the readability of the summary report generated after processing station data files.


from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd

# Try to import rich text formatting (optional, may not be available in all versions)
try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    HAS_RICH_TEXT = True
except ImportError:
    HAS_RICH_TEXT = False

# -------------------------
# Main function to format the Available_Data.xlsx report with all styling and conditional formatting
# -------------------------


# General formatting: highlight overlapping data and time steps != 3

def format_excel_report(file_path):
    """Apply all Excel formatting.
    - Header styling
    - Freeze panes
    - Auto-adjust column widths
    - Add station separator borders
    - Highlight time steps != 3 in red
    - Highlight overlaps (red) and gaps (orange) in After columns
    - Format comments columns (dynamic width + wrap text)"""
    
    try:
        wb = load_workbook(file_path, rich_text=True)
    except TypeError:
        # Older openpyxl versions do not support the rich_text argument.
        wb = load_workbook(file_path)
    ws = wb.active

    # Header formatting
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    ws.freeze_panes = "A2"  # Freeze the header row 

    # Auto-adjust column widths
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        col_letter = get_column_letter(column_cells[0].column)
        ws.column_dimensions[col_letter].width = length + 2

    # Station separator borders
    thick_border = Border(bottom=Side(style='thick'))
    for i in range(2, ws.max_row):
        station_current = ws.cell(row=i, column=1).value
        station_next = ws.cell(row=i+1, column=1).value
        if station_current != station_next:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=i, column=col).border = thick_border

    ws.auto_filter.ref = ws.dimensions # Apply auto-filter to the entire data range

    # Highlight time step != 3
    red_font = Font(color="FF0000")
    col_timestep = next((idx for idx, cell in enumerate(ws[1], 1) if cell.value == "Time step (min)"), None)
    if col_timestep:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_timestep)
            if cell.value != 3 and cell.value is not None:
                cell.font = red_font

    # Highlight overlaps and gaps BETWEEN FILES OF THE SAME STATION
    red_fill = PatternFill(start_color="FFD7D7", end_color="FFD7D7", fill_type="solid")
    #orange_fill = PatternFill(start_color="FFDDA0", end_color="FFDDA0", fill_type="solid")

    col_station = 1  # Column A = Station
    col_last = next((idx for idx, cell in enumerate(ws[1], 1) if cell.value == "Last measurement"), None)
    col_first = next((idx for idx, cell in enumerate(ws[1], 1) if cell.value == "First measurement"), None)

    if col_last and col_first:
        i = 2  # Start at row 2 (data rows)
        while i < ws.max_row:
            current_station = ws.cell(row=i, column=col_station).value
            
            # Find end of current station group
            j = i
            while j <= ws.max_row and ws.cell(row=j, column=col_station).value == current_station:
                j += 1
            station_end = j - 1  # Last row of this station
            
            # Check overlaps WITHIN this station only
            k = i
            while k < station_end:  # Don't compare last file of station
                # Last of current file > First of next file (same station)
                last_value = ws.cell(row=k, column=col_last).value
                first_next = ws.cell(row=k+1, column=col_first).value
                
                if last_value and first_next:
                    try:
                        if pd.to_datetime(last_value) > pd.to_datetime(first_next):
                            ws.cell(row=k, column=col_last).fill = red_fill
                    except:
                        pass
                
                k += 1
            
            i = station_end + 1  # Move to next station

    col_last_processed = next((idx for idx, cell in enumerate(ws[1], 1) if cell.value == "Last measurement AP"), None)
    col_first_processed = next((idx for idx, cell in enumerate(ws[1], 1) if cell.value == "First measurement AP"), None)

    if col_last_processed and col_first_processed:
        i = 2  # Start at row 2 (data rows)
        while i < ws.max_row:
            current_station = ws.cell(row=i, column=col_station).value
            
            # Find end of current station group
            j = i
            while j <= ws.max_row and ws.cell(row=j, column=col_station).value == current_station:
                j += 1
            station_end = j - 1  # Last row of this station
            
            # Check overlaps WITHIN this station only
            k = i
            while k < station_end:  # Don't compare last file of station
                # Last of current file > First of next file (same station)
                last_value = ws.cell(row=k, column=col_last_processed).value
                first_next = ws.cell(row=k+1, column=col_first_processed).value
                
                if last_value and first_next:
                    try:
                        if pd.to_datetime(last_value) > pd.to_datetime(first_next):
                            ws.cell(row=k, column=col_last_processed).fill = red_fill
                    except:
                        pass
                
                k += 1
            
            i = station_end + 1  # Move to next station
    
    # Apply alternating row colors (zebra striping)
    light_gray_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    for row_idx in range(2, ws.max_row + 1):
        # Alternate colors: even rows = light gray, odd rows = white
        fill_to_apply = light_gray_fill if row_idx % 2 == 0 else white_fill
        
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            # Only apply striping if cell doesn't already have a special fill (overlaps/gaps highlighting)
            if cell.fill is None or cell.fill.start_color.index == "00000000":  # No fill or default
                cell.fill = fill_to_apply
    
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):        
        for cell in row:            
            cell.alignment = Alignment(                
                vertical='center', 
                wrapText=cell.alignment.wrapText or False
            )
    
    format_comments_columns(ws)

    wb.save(file_path)


# Format comments columns 

def format_comments_columns(ws):
    """
    Set dynamic width (min 30, max 100) + wrap text for comments columns.
    """
    col_logbook = None
    col_comments = None
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == "Logbook comments":
            col_logbook = idx
        if cell.value == "Comments":
            col_comments = idx

    for col_idx in [col_logbook, col_comments]:
        if col_idx:
            col_letter = get_column_letter(col_idx)
            
            # Calculate column width based on content 
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in ws[col_letter])
            calculated_width = max_length + 2
            ws.column_dimensions[col_letter].width = min(max(calculated_width, 30), 100)
            
            # Enable text wrapping
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.alignment = cell.alignment.copy(wrapText=True)


def colorize_comment_keywords(file_path):
    """
    Colorize specific keywords in comments columns:
    - "Logbook: " in green
    - "Manually processing:" in orange
    - "Overlap: " in blue
    - "Cross-correlation MeteoSwiss:" in red
    """
    try:
        wb = load_workbook(file_path, rich_text=True)
    except TypeError:
        # Older openpyxl versions do not support the rich_text argument.
        wb = load_workbook(file_path)
    ws = wb.active
    
    # Find comments columns (case-insensitive)
    comments_cols = []
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        if cell.value and "comment" in str(cell.value).lower():
            comments_cols.append(col_idx)
    
    if not comments_cols:
        return
    
    # Define keywords and their colors
    keyword_colors = {
        "Logbook: ": "FF00B050",                      # Green
        "Manually processing:": "FFFFA500",          # Orange
        "Overlap: ": "FF0070C0",                      # Blue
        "Cross-correlation MeteoSwiss:": "FFFF0000"  # Red
    }
    
    # Process each comments column
    for col_idx in comments_cols:
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            text_value = cell.value
            
            if not text_value:
                continue
            
            text_value = str(text_value)
            
            # Find all keywords with their positions
            keyword_positions = []
            for keyword in keyword_colors.keys():
                start = 0
                while True:
                    pos = text_value.find(keyword, start)
                    if pos == -1:
                        break
                    keyword_positions.append((pos, len(keyword), keyword))
                    start = pos + len(keyword)
            
            # Skip if no keywords found
            if not keyword_positions:
                continue
            
            # Sort by position to avoid overlaps
            keyword_positions.sort()
            
            # Build the rich text parts
            parts = []
            last_end = 0
            
            for pos, kw_len, keyword in keyword_positions:
                # Add plain text before keyword
                if pos > last_end:
                    parts.append(text_value[last_end:pos])
                
                # Add colored keyword with InlineFont (if rich text available)
                if HAS_RICH_TEXT:
                    color = keyword_colors[keyword]
                    inline_font = InlineFont(b=True, color=color)
                    parts.append(TextBlock(inline_font, keyword))
                else:
                    parts.append(keyword)
                last_end = pos + kw_len
            
            # Add remaining text
            if last_end < len(text_value):
                parts.append(text_value[last_end:])
            
            # Set CellRichText if we have parts (if rich text available)
            if parts and HAS_RICH_TEXT:
                cell.value = CellRichText(*parts)
            elif parts:
                cell.value = ''.join(str(p) for p in parts)
    
    wb.save(file_path)
