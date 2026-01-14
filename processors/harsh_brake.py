"""Harsh brake violation report processor - Using Template 41 per unit."""

import re
import pandas as pd
import json
import time
import os

SUMMARY_TEMPLATE_ID = 89
DETAIL_TEMPLATE_ID = 41


def choose_unit_column(df):
    """Choose the best unit/vehicle identifier column in a DataFrame."""
    preferred = ['grouping', 'group', 'vehicle', 'truck', 'name', 'unit']
    for p in preferred:
        for c in df.columns:
            if p in str(c).lower():
                return c
    return df.columns[0] if len(df.columns) > 0 else None


def find_column(df, keywords):
    """Find a column containing any of the keywords."""
    if df is None or df.empty:
        return None
    for c in df.columns:
        lc = str(c).lower()
        for k in keywords:
            if k in lc:
                return c
    return None


def extract_unit_ids_from_json(summary_path):
    """Extract unit name → unit ID mapping from summary JSON backup."""
    unit_name_to_id = {}

    json_path = summary_path.replace('.xlsx', '_rows_debug.json')
    if not os.path.exists(json_path):
        return {}

    with open(json_path, 'r', encoding='utf-8') as f:
        rows = json.load(f)

    for row in rows:
        if not isinstance(row, dict) or 'c' not in row:
            continue

        cells = row['c']
        unit_name = None
        unit_id = None

        for idx, cell in enumerate(cells):
            if idx == 1:
                if isinstance(cell, str):
                    unit_name = cell.strip()
                elif isinstance(cell, dict) and 't' in cell:
                    unit_name = str(cell['t']).strip()
            if isinstance(cell, dict) and 'u' in cell:
                unit_id = int(cell['u'])

        if unit_name and unit_id:
            unit_name_to_id[unit_name] = unit_id

    print(f"  ✓ Extracted {len(unit_name_to_id)} unit mappings from JSON")
    return unit_name_to_id


def pull_details_for_all_units(api, unit_ids, temp_folder):
    """Pull Template 41 report for each unit and combine results."""
    os.makedirs(temp_folder, exist_ok=True)
    all_details = []

    print(f"\n📊 Pulling Template 41 for {len(unit_ids)} units...")

    for i, unit_id in enumerate(unit_ids, 1):
        print(f"  [{i}/{len(unit_ids)}] Unit ID {unit_id}...", end=' ')
        path = os.path.join(temp_folder, f"unit_{unit_id}.xlsx")

        if api.execute_report(unit_id, DETAIL_TEMPLATE_ID, path, None):
            df = pd.read_excel(path, sheet_name='Live Data')
            df['unit_id'] = unit_id  # ⚡ Add unit ID column to link back to summary
            all_details.append(df)
            print(f"✓ ({len(df)} events)")
        else:
            print("✗ Failed")

        time.sleep(0.4)

    return pd.concat(all_details, ignore_index=True) if all_details else pd.DataFrame()


def merge_harsh_brake_reports(summary_path, details_path, dest_path, api):
    """Merge summary with detail reports and fill Event text in summary."""
    s = pd.read_excel(summary_path, sheet_name='Live Data')
    print(f"Summary rows (raw): {len(s)}")

    s_count = find_column(s, ['count', 'cnt', 'total'])
    s_unit = choose_unit_column(s)

    # EARLY FILTER: keep rows with Count >= 3
    s[s_count] = pd.to_numeric(s[s_count], errors='coerce').fillna(0).astype(int)
    s = s[s[s_count] >= 3].copy()
    print(f"Summary rows after Count>=3 filter: {len(s)}")

    # Extract unit IDs from JSON and filter
    unit_name_to_id = extract_unit_ids_from_json(summary_path)
    filtered_unit_ids = [
        unit_name_to_id[u]
        for u in s[s_unit].astype(str).str.strip()
        if u in unit_name_to_id
    ]
    print(f"Units to pull details for: {len(filtered_unit_ids)}")

    # Pull detailed reports
    temp_folder = os.path.join(os.path.dirname(summary_path), "temp_unit_details")
    details_df = pull_details_for_all_units(api, filtered_unit_ids, temp_folder)
    
    # Remove unwanted columns from details
    cols_to_remove = []
    for col in details_df.columns:
        col_lower = str(col).lower().strip()
        if col_lower in ['event type', 'notification text', 'eventtype', 'notificationtext']:
            cols_to_remove.append(col)
    
    if cols_to_remove:
        details_df = details_df.drop(columns=cols_to_remove)
        print(f"  ✓ Removed columns from details: {cols_to_remove}")
    
    details_df.to_excel(details_path, index=False)

    # Fill Event text in summary using unit_id mapping
    detail_event_col = None
    for col in details_df.columns:
        if 'event' in col.lower() and 'time' not in col.lower():
            detail_event_col = col
            break
    if detail_event_col is None:
        print("⚠ Could not find Event text column in details! Using first column as fallback.")
        detail_event_col = details_df.columns[0]

    s['Event text'] = ''

    # Build a map: first Event text per unit_id
    first_event_map = {}
    for _, row in details_df.iterrows():
        uid = row.get('unit_id')
        event_text = str(row.get(detail_event_col, '')).strip()
        if uid not in first_event_map and event_text:
            first_event_map[uid] = event_text

    # Apply map to summary
    filled = 0
    for idx, row in s.iterrows():
        unit_name = str(row.get(s_unit, '')).strip()
        uid = unit_name_to_id.get(unit_name)
        if uid in first_event_map:
            s.at[idx, 'Event text'] = first_event_map[uid]
            filled += 1

    print(f"✓ Filled Event text for {filled}/{len(s)} summary rows")
    
    # Remove unwanted columns from summary
    cols_to_remove_summary = []
    for col in s.columns:
        col_lower = str(col).lower().strip()
        if col_lower in ['event type', 'notification text', 'eventtype', 'notificationtext']:
            cols_to_remove_summary.append(col)
    
    if cols_to_remove_summary:
        s = s.drop(columns=cols_to_remove_summary)
        print(f"  ✓ Removed columns from summary: {cols_to_remove_summary}")

    # Save final summary
    s.to_excel(dest_path, index=False)
    print(f"✓ Final summary saved: {dest_path}")

    # Cleanup temp folder
    import shutil
    shutil.rmtree(temp_folder, ignore_errors=True)

    return True


def process_harsh_brake_detail(df, template_id, api):
    """Placeholder - not used in this workflow."""
    return df