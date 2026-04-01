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


def get_today_date_string():
    """Get today's date in DD.MM.YYYY format."""
    return datetime.now().strftime("%d.%m.%Y")


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


def find_latest_raw_file(raw_reports_folder, token):
    """Find latest report file in raw folder by token."""
    files = [
        os.path.join(raw_reports_folder, f)
        for f in os.listdir(raw_reports_folder)
        if token in f and f.endswith('.xlsx')
    ]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def parse_report_date(value):
    """Parse a report date in either YYYY-MM-DD or DD.MM.YYYY style."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            continue
    try:
        dt = pd.to_datetime(text, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return None


def to_date_series(series):
    """Convert a pandas series of mixed date strings to date objects."""
    return series.apply(parse_report_date)


def update_overall_summary_row(wb, rpt_date):
    """Update OVERALL SUMMARY totals for a specific report date."""
    if rpt_date is None or "OVERALL SUMMARY " not in [s.name for s in wb.sheets]:
        return

    def read_sheet_df(sheet_name):
        if sheet_name not in [s.name for s in wb.sheets]:
            return pd.DataFrame()
        df = wb.sheets[sheet_name].used_range.options(pd.DataFrame, header=1, index=False).value
        if df is None:
            return pd.DataFrame()
        return df

    idling_df = read_sheet_df("IDLING VIOLATION")
    idling_total = 0
    if not idling_df.empty and "RPT_DT" in idling_df.columns:
        mask = to_date_series(idling_df["RPT_DT"]) == rpt_date
        if "NO OF EVENTS" in idling_df.columns:
            idling_total = pd.to_numeric(idling_df.loc[mask, "NO OF EVENTS"], errors="coerce").fillna(0).sum()

    harsh_df = read_sheet_df("HARSH BRAKE VIOLATION")
    harsh_total = 0
    if not harsh_df.empty and "RPT_DT" in harsh_df.columns:
        mask = to_date_series(harsh_df["RPT_DT"]) == rpt_date
        if "Count of Time received" in harsh_df.columns:
            harsh_total = pd.to_numeric(harsh_df.loc[mask, "Count of Time received"], errors="coerce").fillna(0).sum()

    speed_df = read_sheet_df(" OVER SPEEDING VIOLATION ")
    speed_total = 0
    if not speed_df.empty and "RPT_DT" in speed_df.columns:
        mask = to_date_series(speed_df["RPT_DT"]) == rpt_date
        if "Count" in speed_df.columns:
            speed_total = pd.to_numeric(speed_df.loc[mask, "Count"], errors="coerce").fillna(0).sum()

    night_df = read_sheet_df("NIGHT DRIVING REPORT ")
    night_driving_total = 0
    early_start_total = 0
    if not night_df.empty and "RPT_DT" in night_df.columns:
        mask = to_date_series(night_df["RPT_DT"]) == rpt_date
        offense_col = "Offense" if "Offense" in night_df.columns else None
        if offense_col:
            offenses = night_df.loc[mask, offense_col].astype(str).str.strip().str.lower()
            night_driving_total = int((offenses == "night driving").sum())
            early_start_total = int((offenses == "early start").sum())

    summary_sheet = wb.sheets["OVERALL SUMMARY "]
    summary_sheet.range("B4").value = rpt_date.strftime("%d-%b-%Y")
    summary_sheet.range("C4").value = int(night_driving_total)
    summary_sheet.range("D4").value = int(early_start_total)
    summary_sheet.range("G4").value = float(idling_total)
    summary_sheet.range("H4").value = float(harsh_total)
    summary_sheet.range("I4").value = float(speed_total)


def update_for_sheq_rpt_dt_filters(wb):
    """Refresh FOR SHEQ pivots and set RPT_DT page filters to latest available date."""
    if "FOR SHEQ" not in [s.name for s in wb.sheets]:
        return None

    sheq = wb.sheets["FOR SHEQ"]
    latest_global = None
    rpt_fields = []

    try:
        pivot_count = sheq.api.PivotTables().Count
    except Exception:
        return None

    for i in range(1, pivot_count + 1):
        pt = sheq.api.PivotTables(i)
        try:
            pt.RefreshTable()
        except Exception:
            pass

        rpt_field = None
        try:
            fields = pt.PivotFields()
            for j in range(1, fields.Count + 1):
                field = fields.Item(j)
                if str(field.Name).strip().upper() == "RPT_DT":
                    rpt_field = field
                    break
        except Exception:
            continue

        if rpt_field is None:
            continue

        candidates = []
        try:
            items = rpt_field.PivotItems()
            for k in range(1, items.Count + 1):
                try:
                    name = str(items.Item(k).Name).strip()
                except Exception:
                    continue
                if not name or name.lower() == "(blank)":
                    continue
                parsed = parse_report_date(name)
                if parsed:
                    candidates.append((parsed, name))
        except Exception:
            continue

        if not candidates:
            continue

        candidates.sort(key=lambda x: x[0])
        latest_date, _ = candidates[-1]
        if latest_global is None or latest_date > latest_global:
            latest_global = latest_date
        rpt_fields.append((rpt_field, candidates))

    if latest_global is not None:
        for rpt_field, candidates in rpt_fields:
            try:
                rpt_field.EnableMultiplePageItems = False
            except Exception:
                pass
            target_name = None
            for date_value, name in candidates:
                if date_value == latest_global:
                    target_name = name
                    break
            if target_name is None:
                target_name = candidates[-1][1]
            try:
                rpt_field.CurrentPage = target_name
            except Exception:
                for alt in (latest_global.strftime("%Y-%m-%d"), latest_global.strftime("%d.%m.%Y")):
                    try:
                        rpt_field.CurrentPage = alt
                        break
                    except Exception:
                        continue

    return latest_global


def determine_offense(beginning_time):
    """Determine offense type from Beginning time."""
    try:
        dt = pd.to_datetime(beginning_time, errors='coerce')
        if pd.notna(dt):
            hour = dt.hour
            if hour in [4, 5]:
                return "Early start"
            elif hour in [20, 21, 22, 23]:
                return "Night driving"
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
    raw_df['RPT_DT'] = raw_df['Time'].apply(extract_date_from_event_time)
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
            
            raw_idling_path = find_latest_raw_file(raw_reports_folder, 'IDLING')

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
            
            raw_harsh_path = find_latest_raw_file(raw_reports_folder, 'HARSH_BRAKE_SUMMARY')

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
            
            raw_speed_path = find_latest_raw_file(raw_reports_folder, 'SPEED_VIOLATION')

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
            
            raw_night_path = find_latest_raw_file(raw_reports_folder, 'NIGHT_DRIVING')

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
        
        # Refresh FOR SHEQ pivots and enforce latest RPT_DT page filters
        latest_rpt_date = update_for_sheq_rpt_dt_filters(wb)
        if latest_rpt_date:
            update_overall_summary_row(wb, latest_rpt_date)
            print(f"✓ Updated OVERALL SUMMARY totals for {latest_rpt_date.strftime('%d.%m.%Y')}")

        # Update filename date
        today_date = get_today_date_string()
        new_filename = f"OVERALL VIOLATIONS REPORT {today_date}.xlsx"
        new_path = os.path.join(overall_excel_folder, new_filename)
        
        # Save workbook
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
