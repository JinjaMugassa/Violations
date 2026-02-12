"""Night driving report processor."""
 
import pandas as pd
import pytz

TEMPLATE_ID = 6
TEMPLATE_NAME = "01_RPT_NIGHT DRIVING REPORT (GROUP)"

# Specific locations to include
BORDER_LOCATIONS = [
    'Great North Road, Nakonde',
    'T1, Nakonde',
    'NAKONDE ZM SIDE',
    'TUNDUMA BORDER TZ SIDE',
    'MOKAMBO DRC SIDE',
    'MOKAMBO ZM SIDE',
    'KASUMBALESA DRC SIDE',
    'KASUMBALESA ZAMBBIA SIDE',
    'TUNDUMA PARKING',
    'KASUMBALESA 114 PARKING',
    'GALCO ICD UDA',
    'GALCO ICD MBOZI',
    'TICS',
    'TPA PORT',
    'DP WORLD',
    'KIOO',
    'CHANGOMBE',
    'KITOPENI',
    'MWENGE YARD',
    'MBAGALA',
    'DAR ES SALAAM',
    'GL KURASINI YARD',
    'MOFED',
]

# Keywords to search for in locations
LOCATION_KEYWORDS = ['PARKING', 'MINE', 'MINES']


def get_tanzania_timezone():
    """Returns Tanzania timezone object (UTC+3)."""
    return pytz.timezone('Africa/Dar_es_Salaam')


def is_relevant_location(location_str):
    """Check if location matches our criteria (borders, parking, mines)."""
    if pd.isna(location_str):
        return False
    
    location_upper = str(location_str).upper().strip()
    
    # Check exact matches with border locations
    for border in BORDER_LOCATIONS:
        if border.upper() in location_upper:
            return True
    
    # Check for keywords (PARKING, MINE, MINES)
    for keyword in LOCATION_KEYWORDS:
        if keyword in location_upper:
            return True
    
    return False


def process_night_driving(df, template_id, api):
    """Process night driving report DataFrame.
    
    - Filters by time windows (20:30-23:59 or 04:30-05:40)
    - Filters by location (borders, parking, mines)
    - Removes unwanted columns (Off-time next, Max speed, Driver)
    
    Args:
        df: pandas DataFrame
        template_id: Template ID being processed
        api: WialonAPI instance (unused but kept for consistency)
        
    Returns:
        Processed DataFrame
    """
    if int(template_id) != TEMPLATE_ID:
        return df
    
    try:
        # Debug: Print all columns
        print(f"\n  📋 Night Driving Report Columns:")
        for i, col in enumerate(df.columns):
            print(f"    [{i}] '{col}'")
        
        tz_tz = get_tanzania_timezone()
        
        # Find Begin and End columns
        begin_cols = []
        end_cols = []
        
        for col in list(df.columns):
            lc = str(col).lower()
            if 'begin' in lc or 'start' in lc or lc == 'beginning':
                begin_cols.append(col)
            if 'end' in lc or 'finish' in lc:
                end_cols.append(col)
        
        # Format End columns
        for col in end_cols:
            try:
                parsed_end = pd.to_datetime(df[col], errors='coerce', dayfirst=False)
                if parsed_end.isna().all():
                    parsed_end = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
                
                def format_end(x):
                    if pd.isna(x):
                        return ''
                    if getattr(x, 'tzinfo', None) is None:
                        return (x + pd.Timedelta(hours=3)).strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        return x.tz_convert(tz_tz).strftime('%Y-%m-%d %H:%M:%S')
                
                df[col] = parsed_end.apply(lambda x: format_end(x) if not pd.isna(x) else '')
            except Exception:
                pass
        
        # Process Begin column and filter by time
        if begin_cols:
            bcol = begin_cols[0]
            
            try:
                parsed_begin = pd.to_datetime(df[bcol], errors='coerce', dayfirst=False)
                if parsed_begin.isna().all():
                    parsed_begin = pd.to_datetime(df[bcol], errors='coerce', dayfirst=True)
                
                # Apply timezone correction
                def format_begin(x):
                    if pd.isna(x):
                        return pd.NaT
                    if getattr(x, 'tzinfo', None) is None:
                        return x + pd.Timedelta(hours=3)
                    else:
                        return x.tz_convert(tz_tz)
                
                parsed_fixed = parsed_begin.apply(lambda x: format_begin(x) if not pd.isna(x) else pd.NaT)

                # Filter: keep if time is between 20:30-23:59 OR 04:30-05:40
                def is_night_time(dt):
                    if pd.isna(dt):
                        return False
                    hh = dt.hour
                    mm = dt.minute
                    
                    # Evening window (20:30-23:59)
                    if hh > 20 and hh <= 23:
                        return True
                    if hh == 20 and mm >= 30:
                        return True
                    
                    # Morning window (04:30-05:40)
                    if hh == 4 and mm >= 30:
                        return True
                    if hh == 5 and mm <= 40:
                        return True
                    
                    return False

                time_mask = parsed_fixed.apply(is_night_time)
                original_count = len(df)
                df = df[time_mask.fillna(False)].copy()
                
                print(f"  ✓ Time filtered: {original_count} -> {len(df)} rows (night windows)")

                # Format the begin column to ISO string
                def to_iso_str(x):
                    if pd.isna(x):
                        return ''
                    return x.strftime('%Y-%m-%d %H:%M:%S')
                
                df[bcol] = parsed_fixed[time_mask].apply(to_iso_str).values
            except Exception as e:
                print(f"  ⚠ Warning: Time filtering failed: {e}")
        
        # Filter by Initial location - MORE FLEXIBLE SEARCH
        location_col = None
        
        # Try to find location column with various patterns
        for col in df.columns:
            col_lower = str(col).lower().strip()
            # Check for various location column names
            if any(keyword in col_lower for keyword in [
                'initial location',
                'initial_location', 
                'location',
                'begin location',
                'start location',
                'place',
                'address'
            ]):
                # Avoid time/date columns
                if 'time' not in col_lower and 'date' not in col_lower:
                    location_col = col
                    break
        
        if location_col:
            print(f"  📍 Using location column: '{location_col}'")
            
            # Show sample locations before filtering
            non_empty_locs = df[df[location_col].notna() & (df[location_col] != '')]
            if len(non_empty_locs) > 0:
                print(f"  Sample locations before filter:")
                for i, loc in enumerate(non_empty_locs[location_col].head(5)):
                    print(f"    - {loc}")
            
            before_location = len(df)
            df['_location_match'] = df[location_col].apply(is_relevant_location)
            
            # Show how many matched
            matched_count = df['_location_match'].sum()
            print(f"  Found {matched_count} rows to REMOVE (borders/parking/mines)")
            
            df = df[~df['_location_match']].copy()
            df = df.drop(columns=['_location_match'])
            
            print(f"  ✓ Location filtered: {before_location} -> {len(df)} rows (borders/parking/mines)")
            
            # Show sample locations after filtering
            if len(df) > 0:
                print(f"  Sample locations after filter:")
                for i, loc in enumerate(df[location_col].head(5)):
                    print(f"    - {loc}")
        else:
            print(f"  ⚠ Warning: Could not find location column")
            print(f"  Available columns: {list(df.columns)}")
        
        # Remove unwanted columns: Off-time next, Max speed, Driver
        cols_to_remove = []
        for col in df.columns:
            col_lower = str(col).lower().strip()
            # Match various forms of the column names
            if any(keyword in col_lower for keyword in [
                'off-time next', 'offtime next', 'off time next',
                'max speed', 'maxspeed', 'maximum speed',
                'driver'
            ]):
                cols_to_remove.append(col)
        
        if cols_to_remove:
            df = df.drop(columns=cols_to_remove)
            print(f"  ✓ Removed columns: {cols_to_remove}")

            # --------------------------------------------------
        # Filter out durations less than 10 minutes
        # --------------------------------------------------
        duration_col = None
        for col in df.columns:
            if 'duration' in str(col).lower():
                duration_col = col
                break

        if duration_col:
            before_duration = len(df)

            # Convert duration like "0:16:40" to timedelta
            duration_td = pd.to_timedelta(df[duration_col], errors='coerce')

            # Keep only durations >= 20 minutes
            df = df[duration_td >= pd.Timedelta(minutes=20)].copy()

            print(
                f"  ✓ Duration filtered (<20 min removed): "
                f"{before_duration} -> {len(df)} rows"
            )
        else:
            print("  ⚠ Warning: Duration column not found")

    
    except Exception as e:
        print(f"  ⚠ Warning: Night driving processing failed: {e}")
        import traceback
        traceback.print_exc()
    
    return df 