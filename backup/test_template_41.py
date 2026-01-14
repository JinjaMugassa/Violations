"""Test Template 41 with a specific unit ID to see if it has speed data."""

import os
import sys
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'processors'))

from wialon_api import WialonAPI

TEST_UNIT_ID = 600172549  # T107EBX
TEMPLATE_41 = 41  # Unit-wise harsh brake template


def main():
    """Test Template 41 with specific unit ID."""
    
    output_folder = os.path.join(os.path.dirname(__file__), "test_template_41")
    os.makedirs(output_folder, exist_ok=True)
    
    api = WialonAPI()
    if not api.login():
        print("✗ Failed to login")
        return
    
    try:
        timestamp = datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
        output_path = os.path.join(output_folder, f"TEMPLATE_41_UNIT_{timestamp}.xlsx")
        
        print(f"\n{'='*60}")
        print(f"TESTING TEMPLATE 41 WITH UNIT ID")
        print(f"{'='*60}")
        print(f"Template ID: {TEMPLATE_41}")
        print(f"Unit ID: {TEST_UNIT_ID}")
        print(f"Output: {output_path}\n")
        
        success = api.execute_report(
            TEST_UNIT_ID,  # Use unit ID instead of group ID
            TEMPLATE_41,
            output_path,
            processor_func=None
        )
        
        if success:
            # Load and analyze
            df = pd.read_excel(output_path, sheet_name='Live Data')
            
            print(f"\n{'='*60}")
            print("🎉 SUCCESS! TEMPLATE 41 WORKS WITH UNIT ID")
            print(f"{'='*60}")
            print(f"✓ Rows: {len(df)}")
            print(f"✓ Columns ({len(df.columns)}): {list(df.columns)}\n")
            
            # Check for speed columns
            speed_cols = []
            for col in df.columns:
                col_lower = str(col).lower()
                if any(keyword in col_lower for keyword in ['speed', 'km/h', 'kph', 'velocity', 'km']):
                    speed_cols.append(col)
            
            if speed_cols:
                print(f"🎉🎉🎉 SPEED COLUMNS FOUND: {speed_cols}\n")
                
                # Show sample data
                for col in speed_cols:
                    print(f"Column: '{col}'")
                    sample = df[col].dropna().head(5).tolist()
                    print(f"  Sample values: {sample}\n")
            else:
                print(f"⚠ No dedicated speed columns found\n")
            
            # Check Event text
            if 'Event text' in df.columns:
                print("✓ Event text column exists!")
                
                # Get first non-empty event text
                non_empty = df[df['Event text'].notna() & (df['Event text'] != '')]
                if len(non_empty) > 0:
                    sample = non_empty['Event text'].iloc[0]
                    print(f"\nSample Event text:")
                    print(f"  {sample}")
                    
                    # Check if speed is in the text
                    if 'speed' in str(sample).lower() and ('km/h' in str(sample) or 'kph' in str(sample)):
                        print("\n🎉🎉🎉 SPEED IS IN EVENT TEXT! THIS IS THE TEMPLATE WE NEED! 🎉🎉🎉")
                        print("\nWe can use Template 41 by:")
                        print("  1. Get all units in the group")
                        print("  2. Pull Template 41 for each unit")
                        print("  3. Combine the results")
                    else:
                        print("\n✗ Event text exists but no speed in it")
                else:
                    print("\n⚠ Event text column exists but all rows are empty")
            else:
                print("✗ No Event text column\n")
            
            # Show all columns with samples
            print(f"\n{'='*60}")
            print("ALL COLUMN DETAILS")
            print(f"{'='*60}")
            for col in df.columns:
                non_null = df[col].notna().sum()
                if non_null > 0:
                    sample = df[col].dropna().iloc[0]
                    print(f"\n{col}:")
                    print(f"  Non-null: {non_null}/{len(df)}")
                    print(f"  Sample: {str(sample)[:150]}")
                else:
                    print(f"\n{col}: (all empty)")
            
            print(f"\n{'='*60}")
            print(f"Full Excel saved: {output_path}")
            print("Open it to see all the data!")
            print(f"{'='*60}\n")
            
        else:
            print("\n✗ Template 41 failed even with unit ID")
            print("This template might not exist or we don't have access\n")
    
    finally:
        api.logout()


if __name__ == "__main__":
    main()