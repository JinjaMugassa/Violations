"""Shared utilities for violation reports processing."""

import pandas as pd
import pytz
from datetime import datetime, timezone, timedelta


def get_tanzania_timezone():
    """Returns Tanzania timezone object (UTC+3)."""
    return pytz.timezone('Africa/Dar_es_Salaam')


def get_local_timezone_offset():
    """Returns Tanzania timezone offset in seconds."""
    return 10800  # 3 hours in seconds


def get_yesterday_interval():
    """Returns tuple of (from_timestamp, to_timestamp) for yesterday in Tanzania timezone."""
    tz_local = timezone(timedelta(hours=3))
    now_dt = datetime.now(tz_local)
    yesterday_date = (now_dt.date() - timedelta(days=1))
    start_dt = datetime(
        yesterday_date.year, 
        yesterday_date.month, 
        yesterday_date.day, 
        0, 0, 0, 
        tzinfo=tz_local
    )
    end_dt = start_dt + timedelta(days=1) - timedelta(seconds=1)
    return int(start_dt.timestamp()), int(end_dt.timestamp())


def get_timestamp_string():
    """Returns formatted timestamp string for file naming."""
    tz = timezone(timedelta(hours=3))
    return datetime.now(tz).strftime("%d.%m.%Y_%H-%M-%S")


def format_time_value(tv):
    """Format a time value to ISO string in Tanzania timezone.
    
    Args:
        tv: Time value (string, datetime, or timestamp)
        
    Returns:
        Formatted time string or empty string if invalid
    """
    try:
        parsed = pd.to_datetime(tv, errors='coerce', dayfirst=False)
        if pd.isna(parsed):
            parsed = pd.to_datetime(tv, errors='coerce', dayfirst=True)
        if pd.isna(parsed):
            s_tv = str(tv).strip()
            if s_tv.lower() in ('nan', 'none', ''):
                return ''
            return s_tv
            
        if getattr(parsed, 'tz', None) is None:
            return parsed.strftime('%Y-%m-%d %H:%M:%S')
        else:
            tz_tz = get_tanzania_timezone()
            parsed = parsed.tz_convert(tz_tz)
            return parsed.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        s_tv = str(tv).strip()
        if s_tv.lower() in ('nan', 'none', ''):
            return ''
        return s_tv


def convert_timestamps_to_tanzania(df):
    """Convert datetime columns in DataFrame to Tanzania timezone.
    
    Args:
        df: pandas DataFrame
        
    Returns:
        Modified DataFrame
    """
    try:
        tz_tz = get_tanzania_timezone()
        for col in list(df.columns):
            lc = col.lower()
            if 'time' in lc or 'date' in lc or 'last' in lc:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce', utc=True)
                    df[col] = df[col].dt.tz_convert(tz_tz)
                    df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    try:
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                        df[col] = df[col].dt.tz_localize(pytz.UTC).dt.tz_convert(tz_tz)
                        df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                    except Exception:
                        pass
        print('✓ Converted timestamps to Tanzania time where applicable')
    except Exception:
        pass
    return df


def find_column(df, keywords):
    """Find a column in DataFrame by matching keywords.
    
    Args:
        df: pandas DataFrame
        keywords: List of keywords to search for
        
    Returns:
        Column name or None
    """
    if df is None or df.empty:
        return None
    
    import re
    for c in df.columns:
        try:
            lc = str(c).lower()
        except Exception:
            lc = ''
        norm = re.sub(r"[^a-z0-9]", "", lc)
        for k in keywords:
            if k in norm or k in lc:
                return c
    return None


def choose_unit_column(df):
    """Choose the best unit/vehicle identifier column from DataFrame.
    
    Args:
        df: pandas DataFrame
        
    Returns:
        Column name or None
    """
    candidates = ['unit', 'vehicle', 'truck', 'name', 'group', 'grouping', '№']
    col = find_column(df, candidates)
    if col:
        return col
    
    # Fallback: choose a column with mostly non-numeric strings
    best = None
    best_score = -1
    for c in df.columns:
        sample = df[c].dropna().astype(str).head(50)
        if sample.empty:
            continue
        non_numeric = sum(1 for v in sample if any(ch.isalpha() for ch in v))
        score = non_numeric / len(sample)
        if score > best_score:
            best_score = score
            best = c
    return best


def first_non_empty(series):
    """Return first non-empty value from a pandas Series.
    
    Args:
        series: pandas Series
        
    Returns:
        First non-empty value or empty string
    """
    for v in series:
        if pd.notna(v) and str(v).strip() and str(v).strip().lower() not in ('nan', 'none'):
            return str(v).strip()
    return ''


def extract_speed_from_text(text):
    """Extract speed value from text string.
    
    Args:
        text: String containing speed information
        
    Returns:
        Speed string like "65 km/h" or None
    """
    import re
    m = re.search(r"(\d{1,3})\s*(km/h|kph|kmh)", text, re.IGNORECASE)
    if m:
        return f"{m.group(1)} {m.group(2)}"
    
    m2 = re.search(r"speed\s*[:\-]?\s*(\d{1,3})", text, re.IGNORECASE)
    if m2:
        return f"{m2.group(1)} km/h"
    
    return None


def save_debug_json(data, output_path, suffix):
    """Save debug JSON file.
    
    Args:
        data: Data to save
        output_path: Base output path
        suffix: Suffix for debug filename
    """
    import json
    debug_path = output_path.replace(".xlsx", f"_{suffix}.json")
    try:
        with open(debug_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass