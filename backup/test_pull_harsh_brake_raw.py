"""Test script to pull harsh brake and reconstruct Event text with speed."""

import os
import sys
import json
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'processors'))

from wialon_api import WialonAPI

TARGET_GROUP = "TRANSIT_ALL_TRUCKS"
DETAIL_TEMPLATE_ID = 67


def get_speed_from_unit_position(api, unit_name, event_timestamp):
    """Get speed by loading unit position data at event timestamp."""
    try:
        # Find unit
        unit_id = api.find_unit_id(unit_name)
        if not unit_id:
            print(f"      Unit not found: {unit_name}")
            return None
        
        # Load messages around event time (±5 minutes)
        from_ts = event_timestamp - 300
        to_ts = event_timestamp + 300
        
        params = {
            "svc": "messages/load_interval",
            "params": json.dumps({
                "itemId": int(unit_id),
                "timeFrom": int(from_ts),
                "timeTo": int(to_ts),
                "flags": 0x0001,  # Load with parameters
                "flagsMask": 0xFFFF
            }),
            "sid": api.sid,
        }
        
        resp = api.requests.post(api.api_url, params=params, timeout=30)
        data = resp.json()
        
        if "messages" not in data:
            print(f"      No messages returned (might be error {data.get('error', 'unknown')})")
            return None
        
        messages = data["messages"]
        
        if not messages:
            print(f"      No messages in time window")
            return None
        
        # Find message closest to event timestamp
        closest_msg = min(messages, key=lambda m: abs(m.get("t", 0) - event_timestamp))
        
        # Debug: save first message structure
        if len(messages) > 0:
            debug_path = f"debug_msg_structure_{unit_name}.json"
            with open(debug_path, "w") as f:
                json.dump(closest_msg, f, indent=2)
            print(f"      Saved message structure to {debug_path}")
        
        # Extract speed from closest message
        speed = None
        
        # Try different speed field locations
        if "s" in closest_msg:
            speed = closest_msg["s"]
        elif "pos" in closest_msg and isinstance(closest_msg["pos"], dict):
            if "s" in closest_msg["pos"]:
                speed = closest_msg["pos"]["s"]
        elif "p" in closest_msg and isinstance(closest_msg["p"], dict):
            for key in ["speed", "s", "spd", "velocity"]:
                if key in closest_msg["p"]:
                    speed = closest_msg["p"][key]
                    break
        
        # Validate speed
        if speed is not None:
            try:
                speed_val = float(speed)
                if 0 <= speed_val < 400:
                    print(f"      ✓ Found speed: {int(speed_val)} km/h")
                    return f"{int(speed_val)} km/h"
            except (ValueError, TypeError):
                pass
        
        print(f"      ✗ No speed in message (keys: {list(closest_msg.keys())})")
        
    except Exception as e:
        print(f"      Exception: {e}")
        import traceback
        traceback.print_exc()
    
    return None


def process_harsh_brake_with_speed(df, api):
    """Process harsh brake data and add speed to event text."""
    
    # Find columns
    unit_col = None
    time_col = None
    location_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'group' in col_lower and not unit_col:
            unit_col = col
        elif 'event time' in col_lower and not time_col:
            time_col = col
        elif 'location' in col_lower and not location_col:
            location_col = col
    
    if not unit_col or not time_col:
        print("  ✗ Missing required columns")
        return df
    
    print(f"  Using columns: unit={unit_col}, time={time_col}, location={location_col}")
    
    # Build event text for each row
    event_texts = []
    
    for idx, row in df.iterrows():
        unit = str(row.get(unit_col, '')).strip()
        event_time_data = row.get(time_col, '')
        location = str(row.get(location_col, '')).strip() if location_col else ''
        
        if not unit:
            event_texts.append('')
            continue
        
        # Extract timestamp and formatted time
        event_ts = None
        time_str = ''
        
        if isinstance(event_time_data, dict):
            # It's the object from Wialon: {"t": "...", "v": timestamp, ...}
            event_ts = event_time_data.get('v')
            time_str = event_time_data.get('t', '')
        else:
            # It's a string, parse it
            try:
                time_dt = pd.to_datetime(event_time_data, errors='coerce')
                if pd.notna(time_dt):
                    time_str = time_dt.strftime('%d.%m.%Y %H:%M:%S')
                    event_ts = int(time_dt.timestamp())
                else:
                    time_str = str(event_time_data)
            except:
                time_str = str(event_time_data)
        
        # Also handle location if it's an object
        if isinstance(location, dict):
            location = location.get('t', '')
        
        # Get speed
        print(f"  [{idx+1}/{len(df)}] Getting speed for {unit} at {time_str}...")
        speed = None
        
        if event_ts:
            speed = get_speed_from_unit_position(api, unit, event_ts)
        else:
            print(f"      ✗ No timestamp available")
        
        # Build sentence
        parts = ["Harsh braking has been activated."]
        
        if time_str:
            parts.append(f"At {time_str}")
        
        if speed:
            parts.append(f"it moved with speed {speed}")
        else:
            print(f"      ⚠ Using sentence without speed")
        
        if location:
            parts.append(f"near '{location}'")
        
        sentence = ' '.join(parts)
        if not sentence.endswith('.'):
            sentence += '.'
        
        event_text = f"{unit}: {sentence}"
        event_texts.append(event_text)
    
    # Add Event text column
    df['Event text'] = event_texts
    
    return df


def test_pull_with_speed():
    """Pull harsh brake report and add speed via API."""
    
    output_folder = os.path.join(os.path.dirname(__file__), "test_raw_output")
    os.makedirs(output_folder, exist_ok=True)
    
    api = WialonAPI()
    if not api.login():
        print("✗ Failed to login")
        return
    
    # Store requests module reference
    import requests
    api.requests = requests
    
    try:
        group_id = api.find_group_id(TARGET_GROUP)
        if not group_id:
            print(f"✗ Group not found: {TARGET_GROUP}")
            return
        
        print(f"✓ Found group ID: {group_id}")
        
        timestamp = datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
        
        # Pull raw data first
        raw_path = os.path.join(output_folder, f"HARSH_BRAKE_RAW_{timestamp}.xlsx")
        
        print(f"\n📊 Step 1: Pulling raw harsh brake data...")
        print(f"   Template ID: {DETAIL_TEMPLATE_ID}")
        
        success = api.execute_report(
            group_id,
            DETAIL_TEMPLATE_ID,
            raw_path,
            processor_func=None
        )
        
        if not success:
            print("✗ Failed to pull report")
            return
        
        print(f"✓ Raw data saved: {raw_path}")
        
        # Load and process
        print(f"\n📊 Step 2: Adding speed to event text...")
        df = pd.read_excel(raw_path, sheet_name='Live Data')
        
        print(f"  Columns: {list(df.columns)}")
        print(f"  Rows: {len(df)}")
        
        # Process first 5 rows as test
        df_test = df.head(5).copy()
        df_processed = process_harsh_brake_with_speed(df_test, api)
        
        # Save processed
        processed_path = os.path.join(output_folder, f"HARSH_BRAKE_WITH_SPEED_{timestamp}.xlsx")
        with pd.ExcelWriter(processed_path, engine='openpyxl') as writer:
            df_processed.to_excel(writer, sheet_name='Live Data', index=False)
        
        print(f"\n{'='*60}")
        print("✓ TEST COMPLETE")
        print(f"{'='*60}")
        print(f"\nRaw file: {raw_path}")
        print(f"Processed file: {processed_path}")
        print("\nCheck the processed file to see if Event text has speed!")
        print(f"{'='*60}\n")
        
    finally:
        api.logout()


if __name__ == "__main__":
    print("\n" + "="*60)
    print("HARSH BRAKE SPEED EXTRACTION TEST")
    print("="*60 + "\n")
    
    test_pull_with_speed()