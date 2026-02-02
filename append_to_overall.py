"""Append pulled violation reports to OVERALL VIOLATIONS REPORT excel file."""
import subprocess

subprocess.run(
    ["taskkill", "/f", "/im", "EXCEL.EXE"],
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL
)

import os
import shutil
import glob
from flask import app
import pandas as pd
from datetime import datetime, timedelta
import xlwings as xw

def refresh_pivot_tables_and_filter(wb, sheet_name, latest_date, latest_date_alt=None):
    from datetime import datetime
    
    try:
        sheet = wb.sheets[sheet_name]
        print(f"\n📊 Refreshing pivot tables in '{sheet_name}'...")
        
        excel_sheet = sheet.api
        pivot_tables = excel_sheet.PivotTables()
        
        print(f"  Found {pivot_tables.Count} pivot table(s)")
        
        for i in range(1, pivot_tables.Count + 1):
            pivot = pivot_tables.Item(i)
            print(f"  → Refreshing '{pivot.Name}'...")
            
            pivot.RefreshTable()
            
            # Determine which date format to use based on pivot table name
            if pivot.Name == "PivotTable2":  # Over Speeding uses different format
                filter_date = latest_date_alt if latest_date_alt else latest_date
            else:
                filter_date = latest_date
            
            # ✅ Try to filter by RPT_DT
            try:
                rpt_dt_field = pivot.PivotFields("RPT_DT")
                
                # First, try as Page Field (the standard way)
                try:
                    rpt_dt_field.ClearAllFilters()
                    rpt_dt_field.CurrentPage = filter_date
                    print(f"    ✓ RPT_DT filtered to {filter_date} (Page Field)")
                except Exception as page_error:
                    print(f"    ⚠ Page Field method failed: {page_error}")
                    
                    # List all available dates to see the format
                    print(f"    📋 Available dates in RPT_DT field:")
                    pivot_items = rpt_dt_field.PivotItems()
                    available_dates = []
                    for item_idx in range(1, min(pivot_items.Count + 1, 20)):
                        try:
                            item = pivot_items.Item(item_idx)
                            item_name = str(item.Name)
                            available_dates.append(item_name)
                            print(f"       {item_idx}. '{item_name}'")
                        except:
                            pass
                    
                    # Try different date format variations
                    date_formats_to_try = [
                        filter_date,
                        datetime.strptime(filter_date, "%Y-%m-%d").strftime("%d.%m.%Y") if "-" in filter_date else filter_date,
                        datetime.strptime(filter_date, "%d.%m.%Y").strftime("%Y-%m-%d") if "." in filter_date else filter_date,
                    ]
                    
                    matched_date = None
                    for date_format in date_formats_to_try:
                        if date_format in available_dates:
                            matched_date = date_format
                            print(f"    ✓ Found matching format: '{matched_date}'")
                            break
                    
                    if matched_date:
                        try:
                            rpt_dt_field.ClearAllFilters()
                            rpt_dt_field.CurrentPage = matched_date
                            print(f"    ✓ RPT_DT filtered to {matched_date} (Converted format)")
                        except Exception as convert_error:
                            print(f"    ⚠ Still failed with converted format: {convert_error}")
                    else:
                        print(f"    ✗ Could not find date '{filter_date}' in any format")
                        print(f"    Available formats tried: {date_formats_to_try[:3]}...")
                    
            except Exception as e:
                print(f"    ⚠ RPT_DT filter failed completely: {e}")
                # Try to list all available fields for debugging
                try:
                    print(f"    Available pivot fields in {pivot.Name}:")
                    for field_idx in range(1, pivot.PivotFields().Count + 1):
                        try:
                            field = pivot.PivotFields(field_idx)
                            print(f"      Field {field_idx}: {field.Name}")
                        except:
                            pass
                except Exception as list_err:
                    print(f"    ⚠ Could not list pivot fields: {list_err}")
        
        print("  ✓ All pivot tables refreshed")
        return True
    
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def find_overall_excel(base_folder):
    overall_files = glob.glob(os.path.join(base_folder, "OVERALL VIOLATIONS REPORT *.xlsx"))
    if not overall_files:
        return None, None
    latest_file = max(overall_files, key=os.path.getmtime)
    try:
        date_str = os.path.basename(latest_file).replace("OVERALL VIOLATIONS REPORT ", "").replace(".xlsx","").strip()
    except Exception:
        date_str = None
    return latest_file, date_str


def get_yesterday_date_string():
    """Get yesterday's date in DD.MM.YYYY format."""
    yesterday = datetime.now() - timedelta(days=1)
    return yesterday.strftime("%d.%m.%Y")

def get_latest_file(folder, keyword):
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if keyword in f and f.endswith('.xlsx')
    ]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def extract_date_from_event_time(event_time_str):
    """Extract date from event time string."""
    try:
        if pd.isna(event_time_str) or str(event_time_str).strip() == '':
            return ''
        
        dt = pd.to_datetime(event_time_str, errors='coerce')
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    
    return ''

def extract_date_ddmmyyyy(event_time_str):
    """Extract date in DD.MM.YYYY format from event time string.
    
    CRITICAL: Uses dayfirst=True to correctly interpret DD.MM.YYYY format dates.
    Without this, '01.02.2026' would be interpreted as January 2nd instead of February 1st.
    """
    try:
        if pd.isna(event_time_str) or str(event_time_str).strip() == '':
            return ''

        # Use dayfirst=True to correctly interpret DD.MM.YYYY format
        dt = pd.to_datetime(event_time_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime("%d.%m.%Y")
    except Exception:
        pass

    return ''


def determine_offense(beginning_time):
    """Determine offense type from Beginning time."""
    try:
        dt = pd.to_datetime(beginning_time, errors='coerce')
        if pd.notna(dt):
            hour = dt.hour
            if hour in [4, 5]:
                return "Early start "
            elif hour in [20, 21, 22, 23]:
                return "Night driving "
    except Exception:
        pass
    
    return ""

def prepare_idling_data(raw_df, existing_df):
    """Prepare idling data for appending."""
    if '№' in raw_df.columns:
        raw_df = raw_df.drop(columns=['№'])
    
    column_mapping = {
        'Grouping': 'TRUCK NO',
        'Event time': 'Event time',
        'Time received': 'Time received',
        'Event text': 'Event text',
        'Location': 'Location',
        'Count': 'NO OF EVENTS'
    }
    
    raw_df = raw_df.rename(columns=column_mapping)
    raw_df['RPT_DT'] = raw_df['Event time'].apply(extract_date_from_event_time)
    
    target_columns = ['TRUCK NO', 'Event time', 'RPT_DT', 'Time received', 'Event text', 'Location', 'NO OF EVENTS']
    
    for col in target_columns:
        if col not in raw_df.columns:
            raw_df[col] = ''
    
    raw_df = raw_df[target_columns]
    
    # Remove duplicates
    if not existing_df.empty and 'TRUCK NO' in existing_df.columns and 'Event time' in existing_df.columns:
        existing_keys = set(
            existing_df['TRUCK NO'].astype(str) + '|' + existing_df['Event time'].astype(str)
        )
        
        raw_df['_check_key'] = raw_df['TRUCK NO'].astype(str) + '|' + raw_df['Event time'].astype(str)
        original_count = len(raw_df)
        raw_df = raw_df[~raw_df['_check_key'].isin(existing_keys)]
        raw_df = raw_df.drop(columns=['_check_key'])
        
        removed = original_count - len(raw_df)
        if removed > 0:
            print(f"    Removed {removed} duplicate rows")
    
    return raw_df


def prepare_harsh_brake_data(raw_df, existing_df):
    """Prepare harsh brake data for appending."""
    if '№' in raw_df.columns:
        raw_df = raw_df.drop(columns=['№'])
    
    column_mapping = {
        'Grouping': 'Row Labels',
        'Event time': 'Event time',
        'Event text': 'Event text',
        'Location': 'Location',
        'Count': 'Count of Time received'
    }
    
    raw_df = raw_df.rename(columns=column_mapping)
    raw_df['RPT_DT'] = raw_df['Event time'].apply(extract_date_from_event_time)
    raw_df['DRIVER NAME'] = ''
    
    target_columns = ['Row Labels', 'Event time', 'RPT_DT', 'DRIVER NAME', 'Event text', 'Location', 'Count of Time received']
    
    for col in target_columns:
        if col not in raw_df.columns:
            raw_df[col] = ''
    
    raw_df = raw_df[target_columns]
    
    # Remove duplicates
    if not existing_df.empty and 'Row Labels' in existing_df.columns and 'Event time' in existing_df.columns:
        existing_keys = set(
            existing_df['Row Labels'].astype(str) + '|' + existing_df['Event time'].astype(str)
        )
        
        raw_df['_check_key'] = raw_df['Row Labels'].astype(str) + '|' + raw_df['Event time'].astype(str)
        original_count = len(raw_df)
        raw_df = raw_df[~raw_df['_check_key'].isin(existing_keys)]
        raw_df = raw_df.drop(columns=['_check_key'])
        
        removed = original_count - len(raw_df)
        if removed > 0:
            print(f"    Removed {removed} duplicate rows")
    
    return raw_df


def prepare_speed_data(raw_df, existing_df):
    """Prepare speed violation data for appending."""
    if '№' in raw_df.columns:
        raw_df = raw_df.drop(columns=['№'])
    
    column_mapping = {
        'Grouping': 'TRUCK NO',
        'Time': 'Time',
        'Max speed': 'MAX SPEED',
        'Location': 'Location',
        'Speed limit': 'Speed limit',
        'Count': 'Count'
    }
    
    raw_df = raw_df.rename(columns=column_mapping)
    
    # CRITICAL: Use extract_date_ddmmyyyy which now has dayfirst=True
    # This ensures '01.02.2026 11:05:42 am' is correctly interpreted as February 1st
    raw_df['RPT_DT'] = raw_df['Time'].apply(extract_date_ddmmyyyy)
    raw_df['DRIVER NAME'] = ''
    
    target_columns = ['TRUCK NO', 'Time', 'RPT_DT', 'DRIVER NAME', 'MAX SPEED', 'Location', 'Speed limit', 'Count']
    
    for col in target_columns:
        if col not in raw_df.columns:
            raw_df[col] = ''
    
    raw_df = raw_df[target_columns]
    
    # Remove duplicates
    if not existing_df.empty and 'TRUCK NO' in existing_df.columns and 'Time' in existing_df.columns:
        existing_keys = set(
            existing_df['TRUCK NO'].astype(str) + '|' + existing_df['Time'].astype(str)
        )
        
        raw_df['_check_key'] = raw_df['TRUCK NO'].astype(str) + '|' + raw_df['Time'].astype(str)
        original_count = len(raw_df)
        raw_df = raw_df[~raw_df['_check_key'].isin(existing_keys)]
        raw_df = raw_df.drop(columns=['_check_key'])
        
        removed = original_count - len(raw_df)
        if removed > 0:
            print(f"    Removed {removed} duplicate rows")
    
    return raw_df


def prepare_night_driving_data(raw_df, existing_df):
    """Prepare night driving data for appending."""
    if '№' in raw_df.columns:
        raw_df = raw_df.drop(columns=['№'])
    
    column_mapping = {
        'Grouping': 'Vehicle no',
        'Beginning': 'Beginning',
        'Initial location': 'Initial location',
        'End': 'End',
        'Final location': 'Final location',
        'Duration': 'DURATION',
        'Mileage': 'Mileage'
    }
    
    raw_df = raw_df.rename(columns=column_mapping)
    raw_df['Driver name'] = ''
    raw_df['TM NAME'] = ''
    raw_df['TC NAME'] = ''
    raw_df['RPT_DT'] = raw_df['Beginning'].apply(extract_date_from_event_time)
    raw_df['Offense'] = raw_df['Beginning'].apply(determine_offense)
    
    target_columns = ['Vehicle no', 'Driver name', 'TM NAME', 'TC NAME', 'Beginning', 'RPT_DT', 
                      'Initial location', 'End', 'Final location', 'DURATION', 'Mileage', 'Offense']
    
    for col in target_columns:
        if col not in raw_df.columns:
            raw_df[col] = ''
    
    raw_df = raw_df[target_columns]
    
    # Remove duplicates
    if not existing_df.empty and 'Vehicle no' in existing_df.columns and 'Beginning' in existing_df.columns:
        existing_keys = set(
            existing_df['Vehicle no'].astype(str) + '|' + existing_df['Beginning'].astype(str)
        )
        
        raw_df['_check_key'] = raw_df['Vehicle no'].astype(str) + '|' + raw_df['Beginning'].astype(str)
        original_count = len(raw_df)
        raw_df = raw_df[~raw_df['_check_key'].isin(existing_keys)]
        raw_df = raw_df.drop(columns=['_check_key'])
        
        removed = original_count - len(raw_df)
        if removed > 0:
            print(f"    Removed {removed} duplicate rows")
    
    return raw_df


def append_to_sheet_xlwings(sheet, new_data_df, has_sn=True):
    """Append data to a sheet using xlwings with formatting preservation.
    
    Args:
        sheet: xlwings Sheet object
        new_data_df: DataFrame to append
        has_sn: Whether sheet has S/N column (Column A)
        
    Returns:
        Number of rows appended
    """
    if new_data_df.empty:
        return 0

    last_row = sheet.used_range.last_cell.row
    style_row = max(2, last_row) if last_row > 1 else 2
    rows_added = 0

    # ---- S/N logic (ONLY if has_sn=True) ----
    if has_sn:
        try:
            last_sn_value = sheet.range(f'A{last_row}').value
            if last_sn_value and str(last_sn_value).replace('.0', '').isdigit():
                start_sn = int(float(last_sn_value)) + 1
            else:
                start_sn = last_row
        except Exception:
            start_sn = last_row
    else:
        start_sn = None

    for _, row_data in new_data_df.iterrows():
        current_row = last_row + 1 + rows_added

        # ---- Write S/N ONLY if enabled ----
        start_col = 1
        if has_sn:
            sn_cell = sheet.range(f'A{current_row}')
            sn_cell.value = start_sn + rows_added
            try:
                ref_cell = sheet.range(f'A{style_row}')
                sn_cell.api.Font.Name = ref_cell.api.Font.Name
                sn_cell.api.Font.Size = ref_cell.api.Font.Size
                sn_cell.api.Font.Bold = ref_cell.api.Font.Bold
                sn_cell.number_format = ref_cell.number_format
            except Exception:
                pass
            start_col = 2  # Data starts from column B

        # ---- Write data columns ----
        for col_idx, col_name in enumerate(new_data_df.columns, start=start_col):
            value = row_data[col_name]
            if pd.isna(value):
                value = ''

            col_letter = xw.utils.col_name(col_idx)
            cell = sheet.range(f'{col_letter}{current_row}')
            cell.value = value

            try:
                ref_cell = sheet.range(f'{col_letter}{style_row}')
                cell.api.Font.Name = ref_cell.api.Font.Name
                cell.api.Font.Size = ref_cell.api.Font.Size
                cell.api.Font.Bold = ref_cell.api.Font.Bold
                cell.number_format = ref_cell.number_format
            except Exception:
                pass

        rows_added += 1

    return rows_added


def read_forsheq_grand_totals(wb):
    """
    Reads values from FOR SHEQ pivot tables.
    Returns a dict mapping violation name -> count
    Also returns Night driving and Early start values separately
    """
    sheet = wb.sheets["FOR SHEQ"]
    pivots = sheet.api.PivotTables()

    results = {}
    night_driving_value = None
    early_start_value = None

    for i in range(1, pivots.Count + 1):
        pivot = pivots.Item(i)
        name = pivot.Name
        
        print(f"  Reading pivot table: {name}")

        try:
            # For PivotTable9 which has both Night driving and Early start in COLUMNS
            # For PivotTable9 which has both Night driving and Early start in COLUMNS
            if name == "PivotTable9":
                table_range = pivot.TableRange1
                print(f"    PivotTable9 table has {table_range.Rows.Count} rows, {table_range.Columns.Count} columns")
                
                # Find the Grand Total row (last row of the table)
                grand_total_row_idx = table_range.Rows.Count
                
                # Since headers are empty, we need to check the ROW LABELS instead
                # Look at row 2 (first data row after header) to see what offense types are listed
                print(f"    Checking row labels to identify columns:")
                
                for col_idx in range(1, table_range.Columns.Count + 1):
                    header_cell = table_range.Cells(1, col_idx)
                    header = str(header_cell.Value).strip().upper() if header_cell.Value else ""
                    
                    # Get the value from the Grand Total row for this column
                    value_cell = table_range.Cells(grand_total_row_idx, col_idx)
                    value = value_cell.Value
                    
                    # Also check what's in row 2 (first data row) to see the label
                    label_cell = table_range.Cells(2, col_idx) if table_range.Rows.Count > 1 else None
                    label = str(label_cell.Value).strip().upper() if label_cell and label_cell.Value else ""
                    
                    print(f"      Column {col_idx}: Header='{header}', Row2Label='{label}', Grand Total={value}")
                    
                    if value and str(value).replace('.0', '').replace('.', '').isdigit():
                        try:
                            count = int(float(value))
                        except:
                            count = 0
                        
                        # Skip column 1 (it's the row labels column)
                        if col_idx == 1:
                            continue
                        
                        # ✅ Match by row label if header is empty - DON'T use fallback!
                        if label and ("EARLY" in label or "START" in label):
                            early_start_value = count
                            results["Early start"] = count
                            print(f"        → Column {col_idx} is Early start: {count}")
                        elif label and ("NIGHT" in label or "DRIVING" in label):
                            night_driving_value = count
                            results["Night driving"] = count
                            print(f"        → Column {col_idx} is Night driving: {count}")
                        # ❌ REMOVE THIS FALLBACK - it causes the bug!
                        # The fallback was assigning values when it shouldn't
                
                # ✅ After the loop, set any unassigned values to 0
                if early_start_value is None:
                    early_start_value = 0
                    results["Early start"] = 0
                    print(f"    → Early start not found, setting to 0")
                
                if night_driving_value is None:
                    night_driving_value = 0
                    results["Night driving"] = 0
                    print(f"    → Night driving not found, setting to 0")
                
            else:
                # For other pivot tables, just get the grand total
                grand_total = pivot.DataBodyRange.Cells(
                    pivot.DataBodyRange.Rows.Count,
                    pivot.DataBodyRange.Columns.Count
                ).Value
                print(f"    Grand total value: {grand_total}")
                
                if name == "PivotTable8":
                    results["Exceeded idle"] = int(grand_total or 0)
                    print(f"    → Mapped to 'Exceeded idle'")
                elif name == "PivotTable6":
                    results["HARSH BRAKE"] = int(grand_total or 0)
                    print(f"    → Mapped to 'HARSH BRAKE'")
                elif name in ["PivotTable2", "PivotTable7"]:
                    results["Over Speeding"] = int(grand_total or 0)
                    print(f"    → Mapped to 'Over Speeding'")
                
        except Exception as e:
            print(f"    ⚠ Could not read pivot table: {e}")
            import traceback
            traceback.print_exc()

    print(f"  Grand totals extracted: {results}")
    print(f"  Night driving (for D4): {night_driving_value}")
    print(f"  Early start (for C4): {early_start_value}")
    
    return results, night_driving_value, early_start_value


def update_overall_summary_daily_row(wb, totals_dict, report_date, night_driving_val, early_start_val):
    """
    Updates OVERALL SUMMARY daily row safely even if empty rows exist above table
    Now also writes Night driving to C4 and Early start to D4 directly
    """
    # Find the OVERALL SUMMARY sheet (handle trailing spaces)
    summary_sheet = None
    for sheet in wb.sheets:
        if 'OVERALL SUMMARY' in sheet.name.upper():
            summary_sheet = sheet
            break
    
    if not summary_sheet:
        raise Exception("OVERALL SUMMARY sheet not found")
    
    sheet = summary_sheet

    used = sheet.used_range
    values = used.value

    header_row = None

    # 1️⃣ Find header row dynamically - look for common header keywords
    header_keywords = ["DAYS", "DAY", "DATE", "NIGHT DRIVING", "EARLY START", "EXCEEDED IDLE", 
                       "HARSH BRAKE", "OVER SPEEDING", "OVERSPEED"]
    
    for r_idx, row in enumerate(values, start=1):
        if row:
            # Convert row to uppercase strings and check for multiple keywords
            row_upper = [str(c).upper().strip() if c else "" for c in row]
            matches = sum(1 for keyword in header_keywords if any(keyword in cell for cell in row_upper))
            
            # If we find 3+ keywords in a row, it's likely the header
            if matches >= 3:
                header_row = r_idx
                print(f"  Found header row at row {header_row}")
                print(f"  Headers: {[str(c).strip() if c else '' for c in row[:15]]}")
                break

    if not header_row:
        print("  ⚠ Could not find header row automatically")
        print("  First 10 rows of OVERALL SUMMARY:")
        for i, row in enumerate(values[:10], start=1):
            print(f"    Row {i}: {[str(c)[:30] if c else '' for c in row[:10]]}")
        raise Exception("Header row not found in OVERALL SUMMARY")

    # 2️⃣ Map headers to columns (1-based indexing for Excel)
    headers = values[header_row - 1]
    col_map = {}
    
    print(f"  Raw headers (first 15): {headers[:15]}")
    
    for idx, h in enumerate(headers):
        if h:
            # Excel columns are 1-based, enumerate is 0-based, so we add 1
            header_str = str(h).strip()
            excel_col = idx + 1
            col_map[header_str] = excel_col
            col_map[header_str.upper()] = excel_col
            
            # Print first few mappings for debugging
            if idx < 10:
                col_letter = xw.utils.col_name(excel_col)
                print(f"    Header '{header_str}' -> Column {excel_col} ({col_letter})")
    
    print(f"  Available columns: {list(col_map.keys())[:20]}")

    # 3️⃣ Find the daily data row (skip monthly summary row)
    daily_row = header_row + 2  # Skip the monthly summary row
    
    print(f"  Daily data row (to update): {daily_row}")

    # 4️⃣ Determine if we need to shift columns (column A is empty/for row numbers)
    cell_a_value = sheet.range((daily_row, 1)).value
    col_shift = 0
    if cell_a_value is None or str(cell_a_value).strip() in ['', 'None']:
        print(f"  Column A is empty (row numbers), shifting all data by +1 column")
        col_shift = 1
    
    # 5️⃣ Update DATE
    date_column_names = ["DAYS", "DAY", "DATE", "Days", "Day", "Date"]
    date_col_idx = None
    
    for col_name in date_column_names:
        if col_name in col_map:
            date_col_idx = col_map[col_name] + col_shift
            break
    
    if date_col_idx:
        date_col_letter = xw.utils.col_name(date_col_idx)
        print(f"  Writing date '{report_date}' to cell {date_col_letter}{daily_row}")
        sheet.range((daily_row, date_col_idx)).value = report_date
        sheet.range((daily_row, date_col_idx)).number_format = "dd-mmm-yy"
        print(f"  ✓ Updated date to {report_date}")
    else:
        print(f"  ⚠ Date column not found (looked for: {date_column_names})")

    # 6️⃣ DIRECTLY write Night driving to C4 and Early start to D4
    # 6️⃣ Write Night driving and Early start using header mapping, not direct cell references
    print(f"\n  📝 Writing PivotTable9 values:")

    # Find the actual columns for Night driving and Early start
    night_col_idx = None
    early_col_idx = None

    for col_name in ["Night driving", "NIGHT DRIVING"]:
        if col_name in col_map:
            night_col_idx = col_map[col_name] + col_shift
            break

    for col_name in ["Early start", "EARLY START"]:
        if col_name in col_map:
            early_col_idx = col_map[col_name] + col_shift
            break

    if early_start_val is not None and early_col_idx:
        early_col_letter = xw.utils.col_name(early_col_idx)
        sheet.range((daily_row, early_col_idx)).value = early_start_val
        print(f"    ✓ Cell {early_col_letter}{daily_row} (Early start) = {early_start_val}")
    else:
        if early_col_idx:
            # If no value, write 0
            early_col_letter = xw.utils.col_name(early_col_idx)
            sheet.range((daily_row, early_col_idx)).value = 0
            print(f"    ✓ Cell {early_col_letter}{daily_row} (Early start) = 0 (no data)")

    if night_driving_val is not None and night_col_idx:
        night_col_letter = xw.utils.col_name(night_col_idx)
        sheet.range((daily_row, night_col_idx)).value = night_driving_val
        print(f"    ✓ Cell {night_col_letter}{daily_row} (Night driving) = {night_driving_val}")
    else:
        if night_col_idx:
            # If no value, write 0
            night_col_letter = xw.utils.col_name(night_col_idx)
            sheet.range((daily_row, night_col_idx)).value = 0
            print(f"    ✓ Cell {night_col_letter}{daily_row} (Night driving) = 0 (no data)")
        
  
    # 7️⃣ Update other violation columns (APPLY shift like the others!)
    violations_to_update = ["Exceeded idle", "HARSH BRAKE", "Over Speeding"]
    violations_updated = 0

    for violation, value in totals_dict.items():
        # Skip Night driving and Early start as we handled them directly
        if violation in ["Night driving", "Early start"]:
            continue
            
        # Only update if it's in our list of violations to update
        if violation not in violations_to_update:
            print(f"  ⊘ Skipping '{violation}' (not in update list)")
            continue
            
        # APPLY col_shift here too!
        col_idx = None
        if violation in col_map:
            col_idx = col_map[violation] + col_shift  # ✅ ADD col_shift
        elif violation.upper() in col_map:
            col_idx = col_map[violation.upper()] + col_shift  # ✅ ADD col_shift
        
        if col_idx:
            col_letter = xw.utils.col_name(col_idx)
            print(f"  Writing '{violation}' = {value} to cell {col_letter}{daily_row}")
            sheet.range((daily_row, col_idx)).value = value
            violations_updated += 1
            print(f"  ✓ Updated {violation}: {value}")
        else:
            print(f"  ⚠ Column '{violation}' not found in headers")

    print(f"  Total violations updated: {violations_updated}/{len(violations_to_update)}")

def append_violations_to_overall(raw_reports_folder, overall_excel_folder):
    """Append pulled violation data to OVERALL excel file using xlwings.
    
    Args:
        raw_reports_folder: Folder containing raw pulled reports
        overall_excel_folder: Folder containing OVERALL VIOLATIONS REPORT.xlsx
        
    Returns:
        bool: True if successful, False otherwise
    """
    print(f"\n{'='*60}")
    print("APPENDING TO OVERALL VIOLATIONS REPORT")
    print(f"{'='*60}\n")
    
    # Find OVERALL excel file
    overall_path, current_date = find_overall_excel(overall_excel_folder)
    
    if not overall_path:
        print("✗ OVERALL VIOLATIONS REPORT excel file not found")
        print(f"  Looked in: {overall_excel_folder}")
        return False
    
    print(f"✓ Found OVERALL excel: {os.path.basename(overall_path)}")
    print(f"  Current date in filename: {current_date}\n")
    
    # Create backup
    backup_folder = os.path.join(overall_excel_folder, "backup")
    os.makedirs(backup_folder, exist_ok=True)
    backup_path = os.path.join(backup_folder, os.path.basename(overall_path))
    shutil.copyfile(overall_path, backup_path)
    print(f"✓ Created backup in {backup_folder}")  

    
    app = None
    wb = None
    
    try:
        # Open Excel with xlwings
        print("📖 Opening OVERALL excel with xlwings...")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.enable_events = False
        # Run in background
        wb = app.books.open(
            overall_path,
            update_links=False,
            read_only=False
        )

        
        print(f"Available sheets: {[sheet.name for sheet in wb.sheets]}\n")
        
        # Process IDLING VIOLATION
        print("📊 Processing IDLING VIOLATION...")
        
        if "IDLING VIOLATION" not in [sheet.name for sheet in wb.sheets]:
            print("  ✗ Sheet 'IDLING VIOLATION' not found")
        else:
            sheet = wb.sheets["IDLING VIOLATION"]
            existing_idling = sheet.used_range.options(pd.DataFrame, header=1, index=False).value
            # Force RPT_DT column to text (column D example)
            sheet.range("D:D").number_format = "@"
            # Format Event time column as DateTime
            sheet.range("C:C").number_format = "yyyy-mm-dd hh:mm:ss"
            sheet.range("E:E").number_format = "yyyy-mm-dd hh:mm:ss"


            print(f"  Current rows: {len(existing_idling) if existing_idling is not None else 0}")
            
            raw_idling_path = get_latest_file(raw_reports_folder, 'IDLING')

            if not raw_idling_path:
                print("  ⚠ No raw idling report found")
            else:
                print(f"  Reading: {os.path.basename(raw_idling_path)}")
                
                raw_idling = pd.read_excel(raw_idling_path, sheet_name='Live Data')
                print(f"  Raw data rows: {len(raw_idling)}")
                
                prepared_data = prepare_idling_data(raw_idling, existing_idling if existing_idling is not None else pd.DataFrame())
                
                if prepared_data.empty:
                    print("  ℹ No new data to append (all duplicates)")
                else:
                    rows_added = append_to_sheet_xlwings(sheet, prepared_data)
                    print(f"    ✓ Appended {rows_added} rows to IDLING VIOLATION")
        
        # Process HARSH BRAKE VIOLATION
        print("\n📊 Processing HARSH BRAKE VIOLATION...")
        
        if "HARSH BRAKE VIOLATION" not in [sheet.name for sheet in wb.sheets]:
            print("  ✗ Sheet 'HARSH BRAKE VIOLATION' not found")
        else:
            sheet = wb.sheets["HARSH BRAKE VIOLATION"]
            existing_harsh = sheet.used_range.options(pd.DataFrame, header=1, index=False).value
            # Force RPT_DT column to text (column D example)
            sheet.range("D:D").number_format = "@"
            sheet.range("C:C").number_format = "yyyy-mm-dd hh:mm:ss"

            print(f"  Current rows: {len(existing_harsh) if existing_harsh is not None else 0}")
            
            raw_harsh_path = get_latest_file(raw_reports_folder, 'HARSH_BRAKE_SUMMARY')

            if not raw_harsh_path:
                print("  ⚠ No raw harsh brake report found")
            else:
                print(f"  Reading: {os.path.basename(raw_harsh_path)}")
                
                raw_harsh = pd.read_excel(raw_harsh_path, sheet_name='Sheet1')
                print(f"  Raw data rows: {len(raw_harsh)}")
                
                prepared_data = prepare_harsh_brake_data(raw_harsh, existing_harsh if existing_harsh is not None else pd.DataFrame())
                
                if prepared_data.empty:
                    print("  ℹ No new data to append (all duplicates)")
                else:
                    rows_added = append_to_sheet_xlwings(sheet, prepared_data)
                    print(f"    ✓ Appended {rows_added} rows to HARSH BRAKE VIOLATION")
        
        # Process OVER SPEEDING VIOLATION
        print("\n📊 Processing OVER SPEEDING VIOLATION...")
        
        speed_sheet = None
        for sheet in wb.sheets:
            if 'OVER SPEEDING' in sheet.name.upper() or 'OVERSPEED' in sheet.name.upper():
                speed_sheet = sheet
                break
        
        if not speed_sheet:
            print("  ✗ Sheet 'OVER SPEEDING VIOLATION' not found")
        else:
            print(f"  Using sheet: '{speed_sheet.name}'")
            existing_speed = speed_sheet.used_range.options(pd.DataFrame, header=1, index=False).value
            # Force RPT_DT column to text (column D example)
            speed_sheet.range("D:D").number_format = "@"

            print(f"  Current rows: {len(existing_speed) if existing_speed is not None else 0}")
            
            raw_speed_path = get_latest_file(raw_reports_folder, 'SPEED_VIOLATION')

            if not raw_speed_path:
                print("  ⚠ No raw speed violation report found")
            else:
                print(f"  Reading: {os.path.basename(raw_speed_path)}")
                
                raw_speed = pd.read_excel(raw_speed_path, sheet_name='Live Data')
                print(f"  Raw data rows: {len(raw_speed)}")
                
                prepared_data = prepare_speed_data(raw_speed, existing_speed if existing_speed is not None else pd.DataFrame())
                
                if prepared_data.empty:
                    print("  ℹ No new data to append (all duplicates)")
                else:
                    rows_added = append_to_sheet_xlwings(speed_sheet, prepared_data)
                    print(f"    ✓ Appended {rows_added} rows to {speed_sheet.name}")
        
        # Process NIGHT DRIVING REPORT
        print("\n📊 Processing NIGHT DRIVING REPORT...")
        
        night_sheet = None
        for sheet in wb.sheets:
            if 'NIGHT DRIVING' in sheet.name.upper():
                night_sheet = sheet
                break
        
        if not night_sheet:
            print("  ✗ Sheet 'NIGHT DRIVING REPORT' not found")
        else:
            print(f"  Using sheet: '{night_sheet.name}'")
            existing_night = night_sheet.used_range.options(pd.DataFrame, header=1, index=False).value
            # Force RPT_DT column to text (column D example)
            night_sheet.range("F:F").number_format = "@"
            night_sheet.range("J:J").number_format = "@"
            night_sheet.range("E:E").number_format = "yyyy-mm-dd hh:mm:ss"
            night_sheet.range("H:H").number_format = "yyyy-mm-dd hh:mm:ss"


            print(f"  Current rows: {len(existing_night) if existing_night is not None else 0}")
            
            raw_night_path = get_latest_file(raw_reports_folder, 'NIGHT_DRIVING')

            if not raw_night_path:
                print("  ⚠ No raw night driving report found")
            else:
                print(f"  Reading: {os.path.basename(raw_night_path)}")
                
                raw_night = pd.read_excel(raw_night_path, sheet_name='Live Data')
                print(f"  Raw data rows: {len(raw_night)}")
                
                prepared_data = prepare_night_driving_data(raw_night, existing_night if existing_night is not None else pd.DataFrame())
                
                if prepared_data.empty:
                    print("  ℹ No new data to append (all duplicates)")
                else:
                    rows_added = append_to_sheet_xlwings(night_sheet, prepared_data, has_sn=False)
                    print(f"    ✓ Appended {rows_added} rows to {night_sheet.name}")

    
        
        # Update filename date
       # Update filename date
        yesterday_date = get_yesterday_date_string()
        yesterday_date_formatted = yesterday_date  # DD.MM.YYYY

        today_date = datetime.today().strftime("%d-%m-%Y")

        # Convert to YYYY-MM-DD for filtering
        try:
            dt = datetime.strptime(yesterday_date_formatted, "%d.%m.%Y")
            latest_date_filter = dt.strftime("%Y-%m-%d")  # 2026-01-23
            latest_date_filter_ddmmyyyy = dt.strftime("%d.%m.%Y")  # 23.01.2026 for Over Speeding
        except Exception:
            latest_date_filter = yesterday_date_formatted
            latest_date_filter_ddmmyyyy = yesterday_date_formatted

            # 2. After processing ALL sheets (before pivot table refresh), add:
       # After all data appended, before pivot refresh
        print(f"\n🔄 Forcing formulas to recalculate...")
        try:
            # Toggle calculation mode to force full recalc
            wb.app.api.Calculation = -4135  # xlCalculationManual
            wb.app.api.Calculation = -4105  # xlCalculationAutomatic
            wb.app.api.CalculateFull()
            print(f"  ✓ All formulas recalculated")
        except Exception as e:
            print(f"  ⚠ Recalculation warning: {e}")


                
        print(f"\n🔄 Refreshing Pivot Tables...")
        refresh_pivot_tables_and_filter(wb, "FOR SHEQ", latest_date_filter, latest_date_filter_ddmmyyyy)
        grand_totals, night_driving_val, early_start_val = read_forsheq_grand_totals(wb)
        update_overall_summary_daily_row(wb, grand_totals, dt.strftime("%d-%b-%y"), night_driving_val, early_start_val)
        
        # Save workbook
        new_filename = f"OVERALL VIOLATIONS REPORT {today_date}.xlsx"

        new_path = os.path.join(overall_excel_folder, new_filename)
        
        print(f"\n💾 Saving updated OVERALL excel...")
        wb.save(new_path)
        wb.close()
        app.quit()
        
        print(f"✓ Saved: {new_filename}")
        
        # If filename changed, remove old file
        if new_path != overall_path:
            try:
                os.remove(overall_path)
                print(f"✓ Removed old file: {os.path.basename(overall_path)}")
            except Exception as e:
                print(f"⚠ Could not remove old file: {e}")
        
        print(f"\n{'='*60}")
        print("✓ OVERALL EXCEL UPDATED SUCCESSFULLY")
        print(f"{'='*60}\n")
        
        return True
        
    except Exception as e:
        print(f"\n✗ Error appending to OVERALL excel: {e}")
        import traceback
        traceback.print_exc()
        
        # Clean up xlwings
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
            except:
                pass
        
        print(f"\n⚠ Restoring from backup...")
        try:
            shutil.copyfile(backup_path, overall_path)
            print(f"✓ Original file restored")
        except Exception as restore_error:
            print(f"✗ Failed to restore: {restore_error}")
        
        return False


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("Usage: python append_to_overall.py <raw_reports_folder> <overall_excel_folder>")
        sys.exit(1)
    
    raw_folder = sys.argv[1]
    overall_folder = sys.argv[2]
    
    append_violations_to_overall(raw_folder, overall_folder)