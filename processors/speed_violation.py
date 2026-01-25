"""Speed violation report processor."""

import re
import pandas as pd
import json
import os
import glob
from datetime import datetime, timedelta

TEMPLATE_ID = 3
TEMPLATE_NAME = "01_80 KPH_RPT_SPEED VIOLATION REPORT"


def find_latest_speed_json(folder):
    """Return the latest SPEED_VIOLATION JSON file in the folder or its 'raw' subfolder."""
    # First try the raw subfolder
    raw_folder = os.path.join(folder, "raw")
    if os.path.exists(raw_folder):
        pattern = os.path.join(raw_folder, "*SPEED_VIOLATION*_debug.json")
        files = glob.glob(pattern)
        if files:
            files.sort(key=os.path.getmtime, reverse=True)
            print(f"✓ Found SPEED_VIOLATION JSON in raw subfolder: {os.path.basename(files[0])}")
            return files[0]
    
    # Fall back to main folder
    pattern = os.path.join(folder, "*SPEED_VIOLATION*_debug.json")
    files = glob.glob(pattern)
    if not files:
        print(f"⚠ No SPEED_VIOLATION JSON found in {folder} or {raw_folder}")
        return None
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def process_speed_violation(df, template_id, api, json_folder=None):
    """Process speed violation report DataFrame.

    - Filters rows to keep only speeds >= 85 km/h
    - Fills missing Time values from latest SPEED_VIOLATION JSON by Grouping column
      and adjusts to Tanzania time (+3 hours)
    - Formats Time/Date columns to DD.MM.YYYY HH:MM:SS am/pm
    - Removes Speed, Avg speed, and Driver columns
    """

    # -------------------------------
    # Fill missing Time from JSON backup
    # -------------------------------
    if json_folder:
        json_backup_path = find_latest_speed_json(json_folder)
        if json_backup_path:
            try:
                with open(json_backup_path, "r", encoding="utf-8") as f:
                    backup_data = json.load(f)

                # Extract truck/unit and Time from nested JSON
                json_records = []
                for row in backup_data:
                    try:
                        truck_id = row['c'][1]
                        time_val_str = row['c'][3]['t']  # e.g., "07.01.2026 07:44:06 am"
                        # Convert string → datetime and add 3 hours
                        time_val = pd.to_datetime(time_val_str, dayfirst=True) + timedelta(hours=3)
                        json_records.append({
                            'Grouping': truck_id,
                            'Time': time_val
                        })
                    except Exception:
                        continue

                if json_records:
                    df_backup = pd.DataFrame(json_records)

                    if 'Grouping' in df.columns and 'Time' in df_backup.columns:
                        df = df.merge(
                            df_backup,
                            on='Grouping',
                            how='left',
                            suffixes=('', '_backup')
                        )
                        # Fill missing Time values
                        df['Time'] = df['Time'].combine_first(df['Time_backup'])
                        df = df.drop(columns=['Time_backup'])
                        print(f"✓ Filled missing 'Time' values from JSON backup using 'Grouping' (+3h Tanzania)")
                    else:
                        print("⚠ 'Grouping' column not found or JSON missing Time — skipping JSON fill")
                else:
                    print("⚠ JSON contained no valid records for filling")

            except Exception as e:
                print(f"⚠ Failed to fill missing times from JSON backup: {e}")

    print("Time column preview BEFORE processing:")
    time_cols = [c for c in df.columns if 'time' in str(c).lower()]
    for col in time_cols:
        print(df[col].head(10))

    if int(template_id) != TEMPLATE_ID:
        return df

    try:
        # ------------------------------------------------------------------
        # Find speed column
        # ------------------------------------------------------------------
        speed_col = None
        for c in df.columns:
            lc = str(c).lower()
            if 'speed' in lc or 'km/h' in lc or 'kph' in lc or 'kmh' in lc:
                speed_col = c
                break

        if speed_col is None:
            print("  Warning: No speed column found")
            return df

        # ------------------------------------------------------------------
        # Extract numeric speed
        # ------------------------------------------------------------------
        def extract_speed_numeric(s):
            try:
                if pd.isna(s):
                    return None
                sstr = str(s).replace('km/h','').replace('kph','').replace('kmh','').replace('km','').replace(',', '.').strip()
                m = re.search(r"[-+]?[0-9]*\.?[0-9]+", sstr)
                if m:
                    return float(m.group(0))
            except Exception:
                return None
            return None

        df['_speed_num'] = df[speed_col].apply(extract_speed_numeric)
        original_count = len(df)

        df = df[df['_speed_num'].notna() & (df['_speed_num'] >= 85)]
        df = df.drop(columns=['_speed_num'])

        print(f"  Filtered speed: {original_count} -> {len(df)} rows (speed >= 85 km/h)")

        # ------------------------------------------------------------------
        # Format Time / Date columns
        # ------------------------------------------------------------------
        for col in df.columns:
            lc = str(col).lower()
            if 'time' in lc or 'date' in lc:
                print(f"\n=== Processing datetime column: {col} ===")
                print("Sample raw values:")
                print(df[col].head(10).to_list())

                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime('%d.%m.%Y %I:%M:%S %p').str.lower()
                    print("  Column already datetime → formatted directly")
                    continue

                raw = df[col].astype(str).str.strip()
                parsed = pd.to_datetime(raw, dayfirst=True, errors='coerce')
                mask = parsed.notna()

                failed_rows = df.loc[~mask, [col]].copy()
                if not failed_rows.empty:
                    print(f"  ❌ Unparsed values in '{col}':")
                    print(failed_rows.assign(raw_value=raw[~mask]))

                df.loc[mask, col] = parsed.loc[mask].dt.strftime('%d.%m.%Y %I:%M:%S %p').str.lower()
                print(f"  Parsed: {mask.sum()} | Failed: {(~mask).sum()}")

        # ------------------------------------------------------------------
        # Remove ONLY specific columns
        # ------------------------------------------------------------------
        cols_to_remove = []
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if col_lower in ['speed', 'avg speed', 'average speed', 'driver']:
                cols_to_remove.append(col)

        if cols_to_remove:
            df = df.drop(columns=cols_to_remove)
            print(f"  Removed columns: {cols_to_remove}")

    except Exception as e:
        print(f"  Warning: Speed violation processing failed: {e}")
        import traceback
        traceback.print_exc()

    return df