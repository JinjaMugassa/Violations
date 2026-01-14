"""Idling report processor."""

import pandas as pd


TEMPLATE_ID = 11
TEMPLATE_NAME = "01_RPT_IDLING VIOLATIONS REPORT (GROUP)"


def choose_unit_column(df):
    """Choose the best unit/vehicle identifier column from DataFrame."""
    candidates = ['unit', 'vehicle', 'truck', 'name', 'group', 'grouping', '№']
    for cand in candidates:
        for c in df.columns:
            if cand in str(c).lower():
                return c
    for c in df.columns:
        sample = df[c].dropna().astype(str).head(20)
        if sample.empty:
            continue
        non_numeric = sum(1 for v in sample if any(ch.isalpha() for ch in v))
        if non_numeric > len(sample) * 0.5:
            return c
    return df.columns[0]


def first_non_empty(series):
    """Return first non-empty value from a pandas Series."""
    for v in series:
        if pd.notna(v) and str(v).strip() and str(v).strip().lower() not in ('nan', 'none'):
            return str(v).strip()
    return ''


def process_idling(df, template_id, api):
    """Process idling report DataFrame.
    
    - Keeps time format as YYYY-MM-DD HH:MM:SS
    - Filters Count=1,2 (keeps >=3)
    - Replaces =&gt; with =>
    """
    if int(template_id) != TEMPLATE_ID:
        return df
    
    try:
        unit_col = choose_unit_column(df)
        if unit_col is None:
            unit_col = df.columns[0]
        
        event_col = None
        for cand in ['Event text', 'Notification text', 'Event type', 'Event']:
            for c in df.columns:
                if cand.lower() == str(c).lower():
                    event_col = c
                    break
            if event_col:
                break
        
        if event_col is None:
            for c in df.columns:
                lc = str(c).lower()
                if 'text' in lc or 'message' in lc or 'notification' in lc or 'event' in lc:
                    event_col = c
                    break
        
        if event_col is None:
            df['Event text'] = ''
            event_col = 'Event text'
        
        df['_unit_str'] = df[unit_col].astype(str).fillna('').map(lambda x: x.strip())
        counts = df['_unit_str'].value_counts().to_dict()
        
        event_type_col = None
        for c in df.columns:
            lc = str(c).lower()
            if 'event type' in lc or lc.replace(' ', '') == 'eventtype':
                event_type_col = c
                break
        
        cols_to_agg = [c for c in list(df.columns) if c not in ('_unit_str',)]
        if event_type_col is not None and event_type_col in cols_to_agg:
            cols_to_agg.remove(event_type_col)
        
        agg_funcs = {c: first_non_empty for c in cols_to_agg}
        grouped = df.groupby('_unit_str').agg(agg_funcs).reset_index()
        
        try:
            grouped[unit_col] = grouped['_unit_str']
        except Exception:
            pass
        
        grouped['Count'] = grouped[unit_col].map(lambda u: int(counts.get(u, 0)))
        
        original_cols = [c for c in list(df.columns) if c not in ('_unit_str',)]
        if event_type_col is not None and event_type_col in original_cols:
            original_cols = [c for c in original_cols if c != event_type_col]
        
        final_cols = []
        if unit_col in original_cols:
            final_cols.append(unit_col)
        
        for c in original_cols:
            if c == unit_col:
                continue
            if event_type_col is not None and c == event_type_col:
                continue
            if c in grouped.columns and c not in final_cols:
                final_cols.append(c)
        
        if 'Count' not in final_cols:
            final_cols.append('Count')
        
        # Filter: Remove Count=1,2 (keep >=3)
        try:
            if 'Count' in grouped.columns:
                original_count = len(grouped)
                grouped = grouped[grouped['Count'] >= 3]
                print(f"  Filtered idling: {original_count} -> {len(grouped)} rows (kept Count>=3)")
        except Exception:
            pass
        
        # Replace =&gt; with => in Event text column
        try:
            if event_col in grouped.columns:
                grouped[event_col] = grouped[event_col].astype(str).str.replace('=&gt;', '=>', regex=False)
                grouped[event_col] = grouped[event_col].str.replace('&gt;', '>', regex=False)
                grouped[event_col] = grouped[event_col].str.replace('&lt;', '<', regex=False)
                grouped[event_col] = grouped[event_col].str.replace('&amp;', '&', regex=False)
        except Exception:
            pass
        
        # Keep original time format (YYYY-MM-DD HH:MM:SS) - no changes needed
        
        final_cols = [c for c in final_cols if c in grouped.columns]
        grouped = grouped[final_cols]
        
        return grouped
    
    except Exception as e:
        print(f"  Warning: Idling processing failed: {e}")
        import traceback
        traceback.print_exc()
        return df