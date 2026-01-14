"""Main script to pull violation reports from Wialon and append to OVERALL excel.

This script pulls four types of violation reports for the TRANSIT_ALL_TRUCKS group:
1. Speed Violation (85+ km/h)
2. Harsh Brake Violations (with consolidated summary + details)
3. Idling Violations
4. Night Driving Violations

Then appends the data to OVERALL VIOLATIONS REPORT excel file.
"""

import os
import sys
import time
import shutil
import glob

# Add processors directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'processors'))

from wialon_api import WialonAPI
from utils import get_timestamp_string

# Import processors
from speed_violation import process_speed_violation, TEMPLATE_ID as SPEED_TEMPLATE_ID
from harsh_brake import (
    merge_harsh_brake_reports,
    SUMMARY_TEMPLATE_ID,
    DETAIL_TEMPLATE_ID
)
from idling import process_idling, TEMPLATE_ID as IDLING_TEMPLATE_ID
from night_driving import process_night_driving, TEMPLATE_ID as NIGHT_TEMPLATE_ID


# Configuration
TARGET_GROUP = "TRANSIT_ALL_TRUCKS"
DEFAULT_OUTPUT_FOLDER = r"C:\Users\SAMA\Downloads\OVERALL VIOLATION"


def pull_violation_reports(output_folder=None, group_name=None):
    """Pull all violation reports for specified group.
    
    Args:
        output_folder: Output directory path
        group_name: Target group name
        
    Returns:
        List of dicts with downloaded file info and raw folder path
    """
    if output_folder is None:
        output_folder = DEFAULT_OUTPUT_FOLDER
    if group_name is None:
        group_name = TARGET_GROUP
    
    os.makedirs(output_folder, exist_ok=True)
    
    # Create raw folder for all reports
    raw_folder = os.path.join(output_folder, "raw")
    os.makedirs(raw_folder, exist_ok=True)

    api = WialonAPI()
    if not api.login():
        print("✗ Failed to login to Wialon")
        return [], raw_folder
    
    downloaded = []
    timestamp = get_timestamp_string()
    
    try:
        print(f"\n{'='*60}")
        print(f"PULLING VIOLATION REPORTS FOR GROUP: {group_name}")
        print(f"{'='*60}\n")
        
        group_id = api.find_group_id(group_name)
        if not group_id:
            print(f"✗ Group not found: {group_name}")
            return [], raw_folder
        
        print(f"✓ Found group ID: {group_id}\n")
        
        # 1. Speed Violation Report
        print("📊 [1/4] Pulling Speed Violation Report...")
        speed_path = os.path.join(raw_folder, f"{group_name}_SPEED_VIOLATION_{timestamp}.xlsx")
        json_folder = output_folder
        success = api.execute_report(
            group_id,
            SPEED_TEMPLATE_ID,
            speed_path,
            processor_func=lambda df, template_id, api: process_speed_violation(
                df, template_id, api, json_folder=json_folder
            )
        )
        if success:
            downloaded.append({"type": "SPEED_VIOLATION", "path": speed_path, "template_id": SPEED_TEMPLATE_ID})
        time.sleep(1)
        
        # 2. Idling Report
        print("\n📊 [2/4] Pulling Idling Violations Report...")
        idling_path = os.path.join(raw_folder, f"{group_name}_IDLING_{timestamp}.xlsx")
        success = api.execute_report(group_id, IDLING_TEMPLATE_ID, idling_path, processor_func=process_idling)
        if success:
            downloaded.append({"type": "IDLING", "path": idling_path, "template_id": IDLING_TEMPLATE_ID})
        time.sleep(1)
        
        # 3. Night Driving Report
        print("\n📊 [3/4] Pulling Night Driving Report...")
        night_path = os.path.join(raw_folder, f"{group_name}_NIGHT_DRIVING_{timestamp}.xlsx")
        success = api.execute_report(group_id, NIGHT_TEMPLATE_ID, night_path, processor_func=process_night_driving)
        if success:
            downloaded.append({"type": "NIGHT_DRIVING", "path": night_path, "template_id": NIGHT_TEMPLATE_ID})
        time.sleep(1)
        
        # 4. Harsh Brake Report
        print("\n📊 [4/4] Pulling Harsh Brake Violations Report...")
        summary_path = os.path.join(raw_folder, f"{group_name}_HARSH_BRAKE_SUMMARY_{timestamp}.xlsx")
        success_summary = api.execute_report(group_id, SUMMARY_TEMPLATE_ID, summary_path)
        time.sleep(1)
        if success_summary:
            downloaded.append({"type": "HARSH_BRAKE_SUMMARY", "path": summary_path, "template_id": SUMMARY_TEMPLATE_ID})
        
        details_path = os.path.join(raw_folder, f"{group_name}_HARSH_BRAKE_DETAIL_{timestamp}.xlsx")
        if success_summary:
            print("  → Extracting units and pulling detailed reports...")
            consolidated_path = os.path.join(raw_folder, f"{group_name}_HARSH_BRAKE_CONSOLIDATED_{timestamp}.xlsx")
            merge_success = merge_harsh_brake_reports(summary_path, details_path, consolidated_path, api=api)
            if merge_success:
                downloaded.append({"type": "HARSH_BRAKE_DETAIL", "path": details_path, "template_id": DETAIL_TEMPLATE_ID})
                downloaded.append({"type": "HARSH_BRAKE_CONSOLIDATED", "path": consolidated_path, "template_id": f"{SUMMARY_TEMPLATE_ID}+{DETAIL_TEMPLATE_ID}"})
                try:
                    shutil.copyfile(consolidated_path, summary_path)
                    print(f"  ✓ Updated summary file with enriched data")
                except Exception as e:
                    print(f"  ⚠ Could not overwrite summary: {e}")
    
    finally:
        api.logout()
    
    return downloaded, raw_folder


def print_summary(downloaded_files):
    """Print summary of downloaded reports."""
    print(f"\n{'='*60}")
    print("DOWNLOAD SUMMARY")
    print(f"{'='*60}")
    print(f"✓ Downloaded {len(downloaded_files)} reports:\n")
    for file_info in downloaded_files:
        print(f"  • {file_info['type']}")
        print(f"    Template ID: {file_info['template_id']}")
        print(f"    Path: {file_info['path']}\n")


if __name__ == "__main__":
    print("\n" + "="*60)
    print("WIALON VIOLATION REPORTS PULLER")
    print("="*60)
    
    output_dir = sys.argv[1] if len(sys.argv) > 1 else None
    group = sys.argv[2] if len(sys.argv) > 2 else None
    
    files, raw_folder = pull_violation_reports(output_dir, group)
    print_summary(files)
    
    if files:
        print("✓ All reports downloaded successfully!")
        print("\n" + "="*60)
        print("STEP 2: APPENDING TO OVERALL VIOLATIONS REPORT")
        print("="*60)
        
        try:
            from append_to_overall import append_violations_to_overall
            overall_folder = DEFAULT_OUTPUT_FOLDER
            overall_files = glob.glob(os.path.join(overall_folder, "OVERALL VIOLATIONS REPORT *.xlsx"))
            latest_overall = max(overall_files, key=os.path.getmtime) if overall_files else None
            
            if latest_overall:
                backup_folder = os.path.join(overall_folder, "backup")
                os.makedirs(backup_folder, exist_ok=True)
                shutil.copy(latest_overall, os.path.join(backup_folder, os.path.basename(latest_overall)))
                print(f"✓ Backed up latest overall to {backup_folder}")
                
                append_success = append_violations_to_overall(raw_folder, overall_folder)
                if append_success:
                    print("\n✓ Data successfully appended to OVERALL excel!")
                else:
                    print("\n⚠ Failed to append to OVERALL excel (check errors above)")
                    print("  Raw reports are still available in:", raw_folder)
            else:
                print("⚠ No OVERALL file to append. Only raw reports are saved.")
        
        except ImportError as e:
            print(f"\n⚠ Could not import append_to_overall module: {e}")
            print("  Raw reports are available in:", raw_folder)
        except Exception as e:
            print(f"\n⚠ Error during append operation: {e}")
            print("  Raw reports are still available in:", raw_folder)
    else:
        print("✗ No reports were downloaded.")
    
    print("="*60 + "\n")
