"""Pull violation reports from Template 9 and generate 4 raw files for OVERALL append.

This workflow uses:
- Template 9 / table 0 (Events)
- Template 9 / table 1 (Trips) for night mileage estimation

Generated raw files preserve legacy naming so append_to_overall.py can keep working.
"""

import glob
import html
import os
import re
import shutil
import sys
import time
from datetime import datetime, timezone, timedelta

import pandas as pd
from utils import get_timestamp_string
from wialon_api import WialonAPI


TARGET_GROUP = "TRANSIT_ALL_TRUCKS"
DEFAULT_OUTPUT_FOLDER = r"C:\Users\arksecurity\Downloads\OVERALL VIOLATION"
ALL_VIOLATIONS_TEMPLATE_ID = 9
NIGHT_TEMPLATE_ID = 6

NIGHT_LOCATION_EXCLUDE_TERMS = [
    "UBUNGO",
    "KIMARA",
    "DAR ES SALAAM",
    "TICS",
    "DP WORLD",
    "TEMEKE",
    "MBAGALA",
    "BUGURUNI",
    "CHANGOMBE",
    "KASUMBALESA",
    "NAKONDE",
    "TUNDUMA",
    "SAKANIA",
    "MOKAMBO",
]


def _parse_speed_kmh(text):
    if pd.isna(text):
        return None
    m = re.search(r"speed\s+(\d+(?:\.\d+)?)\s*km/h", str(text), flags=re.IGNORECASE)
    if not m:
        return None
    try:
        return float(m.group(1))
    except Exception:
        return None


def _extract_event_type(raw_text):
    text = html.unescape(str(raw_text or "")).upper()
    if "HARSH BRAKING HAS BEEN ACTIVATED" in text:
        return "HARSH_BRAKE"
    if "=> SPEEDING" in text or "=>SPEEDING" in text:
        return "SPEEDING"
    if "=> IDLING" in text or "=>IDLING" in text:
        return "IDLING"
    if "=> NIGHT DRIVING" in text or "=>NIGHT DRIVING" in text:
        return "NIGHT_DRIVING"
    return "OTHER"


def _first_non_empty(values):
    for value in values:
        if pd.notna(value) and str(value).strip():
            return value
    return ""


def _to_datetime(series):
    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().all():
        dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return dt


def _format_duration(delta):
    if pd.isna(delta):
        return ""
    secs = int(max(0, delta.total_seconds()))
    hh = secs // 3600
    mm = (secs % 3600) // 60
    ss = secs % 60
    return f"{hh}:{mm:02d}:{ss:02d}"


def _parse_mileage_km(value):
    if pd.isna(value):
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?", str(value))
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def _duration_hours(begin, end):
    if pd.isna(begin) or pd.isna(end) or end <= begin:
        return 0.0
    return max(0.0, (end - begin).total_seconds() / 3600.0)


def _estimate_mileage_from_speed_profile(unit, begin, end, trips, t6):
    duration_h = _duration_hours(begin, end)
    if duration_h <= 0:
        return 0.0

    speed_candidates = []

    # Unit-specific speeds from Trips
    if trips is not None and not trips.empty:
        ut = trips[trips["Grouping"].astype(str).str.strip() == str(unit).strip()].copy()
        ut = ut[ut["Begin dt"].notna() & ut["End dt"].notna() & ut["Mileage km"].notna()].copy()
        if not ut.empty:
            trip_hours = (ut["End dt"] - ut["Begin dt"]).dt.total_seconds() / 3600.0
            rates = (ut["Mileage km"] / trip_hours).replace([pd.NA, pd.NaT, float("inf"), -float("inf")], pd.NA).dropna()
            rates = rates[(rates > 1) & (rates < 120)]
            if not rates.empty:
                speed_candidates.append(float(rates.median()))

    # Unit-specific speeds from Template 6
    if t6 is not None and not t6.empty:
        ut6 = t6[t6["Grouping"].astype(str).str.strip() == str(unit).strip()].copy()
        ut6 = ut6[ut6["Begin dt"].notna() & ut6["End dt"].notna() & ut6["Mileage km"].notna()].copy()
        if not ut6.empty:
            h = (ut6["End dt"] - ut6["Begin dt"]).dt.total_seconds() / 3600.0
            rates = (ut6["Mileage km"] / h).replace([pd.NA, pd.NaT, float("inf"), -float("inf")], pd.NA).dropna()
            rates = rates[(rates > 1) & (rates < 120)]
            if not rates.empty:
                speed_candidates.append(float(rates.median()))

    # Global fallback speed profile
    global_speed = None
    if trips is not None and not trips.empty:
        all_t = trips[trips["Begin dt"].notna() & trips["End dt"].notna() & trips["Mileage km"].notna()].copy()
        if not all_t.empty:
            h = (all_t["End dt"] - all_t["Begin dt"]).dt.total_seconds() / 3600.0
            rates = (all_t["Mileage km"] / h).replace([pd.NA, pd.NaT, float("inf"), -float("inf")], pd.NA).dropna()
            rates = rates[(rates > 1) & (rates < 120)]
            if not rates.empty:
                global_speed = float(rates.median())

    if speed_candidates:
        speed = float(pd.Series(speed_candidates).median())
    elif global_speed is not None:
        speed = global_speed
    else:
        speed = 35.0

    est = speed * duration_h
    return max(0.5, round(est, 2))


def _estimate_mileage_from_duration_text(duration_value, default_speed_kmh=35.0):
    try:
        td = pd.to_timedelta(duration_value, errors="coerce")
        if pd.isna(td):
            return 1.0
        hours = max(0.0, td.total_seconds() / 3600.0)
        if hours <= 0:
            return 1.0
        return max(0.5, round(default_speed_kmh * hours, 2))
    except Exception:
        return 1.0


def _build_speed_report(events):
    speed = events[events["event_type"] == "SPEEDING"].copy()
    if speed.empty:
        return pd.DataFrame(columns=["№", "Grouping", "Time", "Max speed", "Location", "Speed limit", "Count"])

    speed["speed_num"] = speed["Event text clean"].apply(_parse_speed_kmh)
    speed["rpt_date"] = speed["Event time dt"].dt.date
    speed = speed[speed["rpt_date"].notna()].copy()

    rows = []
    for _, grp in speed.groupby(["Grouping", "rpt_date"], sort=True):
        grp_sorted = grp.sort_values("Event time dt")
        max_speed = pd.to_numeric(grp_sorted["speed_num"], errors="coerce").max()
        best_row = grp_sorted.loc[grp_sorted["speed_num"].fillna(-1).idxmax()] if grp_sorted["speed_num"].notna().any() else grp_sorted.iloc[0]
        rows.append({
            "Grouping": grp_sorted.iloc[0]["Grouping"],
            "Time": grp_sorted.iloc[0]["Event time dt"],
            "Max speed": f"{int(round(max_speed))} km/h" if pd.notna(max_speed) else "",
            "Location": best_row.get("Location", "") or _first_non_empty(grp_sorted["Location"]),
            "Speed limit": "81 km/h",
            "Count": int(len(grp_sorted)),
        })

    out = pd.DataFrame(rows).sort_values(["Grouping", "Time"]).reset_index(drop=True)
    out.insert(0, "№", range(1, len(out) + 1))
    return out


def _build_idling_report(events):
    idle = events[events["event_type"] == "IDLING"].copy()
    if idle.empty:
        return pd.DataFrame(columns=["Grouping", "№", "Event time", "Time received", "Event text", "Location", "Count"])

    idle["rpt_date"] = idle["Event time dt"].dt.date
    idle = idle[idle["rpt_date"].notna()].copy()

    rows = []
    for _, grp in idle.groupby(["Grouping", "rpt_date"], sort=True):
        grp_sorted = grp.sort_values("Event time dt")
        rows.append({
            "Grouping": grp_sorted.iloc[0]["Grouping"],
            "Event time": grp_sorted.iloc[0]["Event time dt"],
            "Time received": grp_sorted["Time received dt"].min() if grp_sorted["Time received dt"].notna().any() else grp_sorted.iloc[0].get("Time received", ""),
            "Event text": grp_sorted.iloc[0]["Event text clean"],
            "Location": _first_non_empty(grp_sorted["Location"]),
            "Count": int(len(grp_sorted)),
        })

    out = pd.DataFrame(rows)
    out = out[out["Count"] >= 3].copy()
    out = out.sort_values(["Grouping", "Event time"]).reset_index(drop=True)
    out.insert(1, "№", range(1, len(out) + 1))
    return out


def _build_harsh_report(events):
    harsh = events[events["event_type"] == "HARSH_BRAKE"].copy()
    if harsh.empty:
        return pd.DataFrame(columns=["№", "Grouping", "Event time", "Time received", "Event text", "Location", "Count"])

    harsh["rpt_date"] = harsh["Event time dt"].dt.date
    harsh = harsh[harsh["rpt_date"].notna()].copy()

    rows = []
    for _, grp in harsh.groupby(["Grouping", "rpt_date"], sort=True):
        grp_sorted = grp.sort_values("Event time dt")
        rows.append({
            "Grouping": grp_sorted.iloc[0]["Grouping"],
            "Event time": grp_sorted.iloc[0]["Event time dt"],
            "Time received": grp_sorted["Time received dt"].min() if grp_sorted["Time received dt"].notna().any() else grp_sorted.iloc[0].get("Time received", ""),
            "Event text": grp_sorted.iloc[0]["Event text clean"],
            "Location": _first_non_empty(grp_sorted["Location"]),
            "Count": int(len(grp_sorted)),
        })

    out = pd.DataFrame(rows)
    out = out[out["Count"] >= 3].copy()
    out = out.sort_values(["Grouping", "Event time"]).reset_index(drop=True)
    out.insert(0, "№", range(1, len(out) + 1))
    return out


def _estimate_night_mileage_km(trips, unit, begin, end):
    if trips.empty or pd.isna(begin) or pd.isna(end) or end <= begin:
        return None

    t = trips[trips["Grouping"].astype(str).str.strip() == str(unit).strip()].copy()
    if t.empty:
        return None

    t = t[t["Begin dt"].notna() & t["End dt"].notna()].copy()
    if t.empty:
        return None

    overlap_sum = 0.0
    found_overlap = False

    for _, row in t.iterrows():
        tb = row["Begin dt"]
        te = row["End dt"]
        mileage_km = row.get("Mileage km")
        if pd.isna(tb) or pd.isna(te) or te <= tb or pd.isna(mileage_km):
            continue

        overlap_start = max(tb, begin)
        overlap_end = min(te, end)
        if overlap_end <= overlap_start:
            continue

        trip_seconds = (te - tb).total_seconds()
        overlap_seconds = (overlap_end - overlap_start).total_seconds()
        if trip_seconds <= 0 or overlap_seconds <= 0:
            continue

        overlap_sum += float(mileage_km) * (overlap_seconds / trip_seconds)
        found_overlap = True

    if not found_overlap:
        return None
    return max(0.0, overlap_sum)


def _build_night_report(events):
    night = events[events["event_type"] == "NIGHT_DRIVING"].copy()
    if night.empty:
        return pd.DataFrame(columns=["№", "Grouping", "Beginning", "Initial location", "End", "Final location", "Duration", "Mileage"])

    night["rpt_date"] = night["Event time dt"].dt.date
    night = night[night["rpt_date"].notna()].copy()

    rows = []
    for _, grp in night.groupby(["Grouping", "rpt_date"], sort=True):
        grp_sorted = grp.sort_values("Event time dt")
        begin = grp_sorted.iloc[0]["Event time dt"]
        end = grp_sorted.iloc[-1]["Event time dt"]
        if pd.isna(begin) or pd.isna(end):
            continue

        hour = begin.hour
        minute = begin.minute
        is_early_start = (hour == 4) or (hour == 5 and minute <= 39)
        is_night_driving = (hour > 19) or (hour == 19 and minute >= 30)

        # Ignore anything outside requested windows.
        if not (is_early_start or is_night_driving):
            continue

        duration_delta = end - begin
        if duration_delta < timedelta(minutes=30):
            continue

        first_loc = str(grp_sorted.iloc[0].get("Location", "") or "").upper()
        last_loc = str(grp_sorted.iloc[-1].get("Location", "") or "").upper()
        combined_loc = f"{first_loc} {last_loc}"

        # Ignore common local/border/parking locations.
        if "PARKING" in combined_loc:
            continue
        if any(term in combined_loc for term in NIGHT_LOCATION_EXCLUDE_TERMS):
            continue

        if is_night_driving:
            # If start is before 20:30, require end >= 21:00.
            if (hour < 20) or (hour == 20 and minute < 30):
                if (end.hour < 21):
                    continue

        rows.append({
            "Grouping": grp_sorted.iloc[0]["Grouping"],
            "Beginning": begin,
            "Initial location": grp_sorted.iloc[0].get("Location", ""),
            "End": end,
            "Final location": grp_sorted.iloc[-1].get("Location", ""),
            "Duration": _format_duration(duration_delta),
            "Mileage": "",
        })

    out = pd.DataFrame(rows).sort_values(["Grouping", "Beginning"]).reset_index(drop=True)
    out.insert(0, "№", range(1, len(out) + 1))
    return out


def _enrich_night_mileage_from_template6(night_df, template6_df, trips=None):
    if night_df.empty or template6_df is None or template6_df.empty:
        return night_df

    t6 = template6_df.copy()
    t6["Begin dt"] = _to_datetime(t6.get("Beginning"))
    t6["End dt"] = _to_datetime(t6.get("End"))
    t6["Mileage km"] = t6.get("Mileage", pd.Series(dtype=object)).apply(_parse_mileage_km)
    t6 = t6[t6["Grouping"].notna() & t6["Begin dt"].notna() & t6["End dt"].notna()].copy()
    if t6.empty:
        return night_df

    out = night_df.copy()
    out["Begin dt"] = _to_datetime(out.get("Beginning"))
    out["End dt"] = _to_datetime(out.get("End"))
    missing_mask = out["Mileage"].isna() | out["Mileage"].astype(str).str.strip().isin(["", "nan", "None", "NONE"])
    if not missing_mask.any():
        return out.drop(columns=["Begin dt", "End dt"], errors="ignore")

    for idx in out[missing_mask].index:
        unit = str(out.at[idx, "Grouping"]).strip()
        b = out.at[idx, "Begin dt"]
        e = out.at[idx, "End dt"]
        if pd.isna(b) or pd.isna(e):
            est = _estimate_mileage_from_duration_text(out.at[idx, "Duration"])
            out.at[idx, "Mileage"] = f"{float(est):.2f} km"
            continue

        cand = t6[t6["Grouping"].astype(str).str.strip() == unit].copy()
        if cand.empty:
            continue

        # Prefer interval overlap match.
        cand["overlap_sec"] = cand.apply(
            lambda r: max(
                0.0,
                (min(r["End dt"], e) - max(r["Begin dt"], b)).total_seconds(),
            ),
            axis=1,
        )
        ov = cand[cand["overlap_sec"] > 0].sort_values("overlap_sec", ascending=False)
        chosen = None
        if not ov.empty:
            chosen = ov.iloc[0]
        else:
            # Fallback to closest beginning time within 12h.
            cand["begin_diff_sec"] = (cand["Begin dt"] - b).abs().dt.total_seconds()
            near = cand[cand["begin_diff_sec"] <= 12 * 3600].sort_values("begin_diff_sec")
            if not near.empty:
                chosen = near.iloc[0]

        if chosen is not None and pd.notna(chosen.get("Mileage km")):
            out.at[idx, "Mileage"] = f"{float(chosen['Mileage km']):.2f} km"
            continue

        # Final fallback: estimate from unit speed profile to avoid blanks.
        est = _estimate_mileage_from_speed_profile(unit, b, e, trips=trips, t6=t6)
        out.at[idx, "Mileage"] = f"{float(est):.2f} km"

    # Hard guard: never leave mileage blank.
    final_missing = out["Mileage"].isna() | out["Mileage"].astype(str).str.strip().isin(["", "nan", "None", "NONE"])
    for idx in out[final_missing].index:
        est = _estimate_mileage_from_duration_text(out.at[idx, "Duration"])
        out.at[idx, "Mileage"] = f"{float(est):.2f} km"

    return out.drop(columns=["Begin dt", "End dt"], errors="ignore")


def _load_template9_raw(events_path, trips_path=None):
    events = pd.read_excel(events_path, sheet_name="Live Data")
    events["Event text clean"] = events.get("Event text", "").map(lambda x: html.unescape(str(x or "")))
    events["event_type"] = events["Event text clean"].apply(_extract_event_type)
    events["Event time dt"] = _to_datetime(events.get("Event time"))
    events["Time received dt"] = _to_datetime(events.get("Time received"))

    if trips_path and os.path.exists(trips_path):
        trips = pd.read_excel(trips_path, sheet_name="Live Data")
    else:
        trips = pd.DataFrame(columns=["Grouping", "Beginning", "End", "Mileage"])
    trips["Begin dt"] = _to_datetime(trips.get("Beginning"))
    trips["End dt"] = _to_datetime(trips.get("End"))
    trips["Mileage km"] = trips.get("Mileage", pd.Series(dtype=object)).apply(_parse_mileage_km)
    return events, trips


def _save_report(df, path, sheet_name):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def _generate_derived_reports(events_path, trips_path, raw_folder, group_name, timestamp, night_fallback_path=None):
    events, trips = _load_template9_raw(events_path, trips_path)

    speed_df = _build_speed_report(events)
    idling_df = _build_idling_report(events)
    harsh_df = _build_harsh_report(events)
    night_df = _build_night_report(events)
    if night_fallback_path and os.path.exists(night_fallback_path):
        try:
            night_fallback_df = pd.read_excel(night_fallback_path, sheet_name="Live Data")
            night_df = _enrich_night_mileage_from_template6(night_df, night_fallback_df, trips=trips)
        except Exception as e:
            print(f"  ⚠ Could not apply Template 6 mileage fallback: {e}")
    elif not night_df.empty:
        # Ensure non-empty mileage even without Template 6.
        night_df["Begin dt"] = _to_datetime(night_df.get("Beginning"))
        night_df["End dt"] = _to_datetime(night_df.get("End"))
        for idx in night_df.index:
            if str(night_df.at[idx, "Mileage"]).strip():
                continue
            unit = str(night_df.at[idx, "Grouping"]).strip()
            b = night_df.at[idx, "Begin dt"]
            e = night_df.at[idx, "End dt"]
            est = _estimate_mileage_from_speed_profile(unit, b, e, trips=trips, t6=None)
            night_df.at[idx, "Mileage"] = f"{float(est):.2f} km"
        night_df = night_df.drop(columns=["Begin dt", "End dt"], errors="ignore")

    speed_path = os.path.join(raw_folder, f"{group_name}_SPEED_VIOLATION_{timestamp}.xlsx")
    idling_path = os.path.join(raw_folder, f"{group_name}_IDLING_{timestamp}.xlsx")
    harsh_path = os.path.join(raw_folder, f"{group_name}_HARSH_BRAKE_SUMMARY_{timestamp}.xlsx")
    night_path = os.path.join(raw_folder, f"{group_name}_NIGHT_DRIVING_{timestamp}.xlsx")

    _save_report(speed_df, speed_path, "Live Data")
    _save_report(idling_df, idling_path, "Live Data")
    _save_report(harsh_df, harsh_path, "Sheet1")
    _save_report(night_df, night_path, "Live Data")

    print(f"  ✓ Derived SPEED rows: {len(speed_df)}")
    print(f"  ✓ Derived IDLING rows: {len(idling_df)}")
    print(f"  ✓ Derived HARSH rows: {len(harsh_df)}")
    print(f"  ✓ Derived NIGHT rows: {len(night_df)}")

    missing_mileage = int((night_df["Mileage"].astype(str).str.strip() == "").sum()) if not night_df.empty else 0
    if missing_mileage:
        print(f"  ⚠ Night rows without estimated mileage: {missing_mileage}")

    return [
        {"type": "SPEED_VIOLATION", "path": speed_path, "template_id": "9/events"},
        {"type": "IDLING", "path": idling_path, "template_id": "9/events"},
        {"type": "NIGHT_DRIVING", "path": night_path, "template_id": "9/events+trips"},
        {"type": "HARSH_BRAKE_SUMMARY", "path": harsh_path, "template_id": "9/events"},
    ]


def _parse_date_arg(value):
    if not value:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            continue
    raise ValueError(f"Unsupported date format: {value} (use YYYY-MM-DD or DD.MM.YYYY)")


def _date_to_interval(report_date):
    tz_local = timezone(timedelta(hours=3))
    start_dt = datetime(
        report_date.year,
        report_date.month,
        report_date.day,
        0, 0, 0,
        tzinfo=tz_local,
    )
    end_dt = start_dt + timedelta(days=1) - timedelta(seconds=1)
    return int(start_dt.timestamp()), int(end_dt.timestamp())


def _daterange(start_date, end_date):
    current = start_date
    while current <= end_date:
        yield current
        current += timedelta(days=1)


def pull_violation_reports(output_folder=None, group_name=None, report_date=None, raw_subfolder=None):
    if output_folder is None:
        output_folder = DEFAULT_OUTPUT_FOLDER
    if group_name is None:
        group_name = TARGET_GROUP

    os.makedirs(output_folder, exist_ok=True)
    if raw_subfolder is None and report_date is not None:
        raw_subfolder = report_date.strftime("%Y-%m-%d")
    raw_folder = os.path.join(output_folder, "raw", raw_subfolder) if raw_subfolder else os.path.join(output_folder, "raw")
    os.makedirs(raw_folder, exist_ok=True)

    api = WialonAPI()
    if not api.login():
        print("✗ Failed to login to Wialon")
        return [], raw_folder

    downloaded = []
    timestamp = get_timestamp_string()

    try:
        print(f"\n{'='*60}")
        if report_date:
            print(f"PULLING VIOLATION REPORTS FOR GROUP: {group_name} | DATE: {report_date.isoformat()}")
        else:
            print(f"PULLING VIOLATION REPORTS FOR GROUP: {group_name}")
        print(f"{'='*60}\n")

        group_id = api.find_group_id(group_name)
        if not group_id:
            print(f"✗ Group not found: {group_name}")
            return [], raw_folder

        print(f"✓ Found group ID: {group_id}\n")
        print("📊 [1/1] Pulling Template 9 report tables (Events + Trips)...")

        events_path = os.path.join(raw_folder, f"{group_name}_ALL_VIOLATIONS_T9_{timestamp}.xlsx")
        trips_path = os.path.join(raw_folder, f"{group_name}_ALL_VIOLATIONS_T9_TRIPS_{timestamp}.xlsx")

        interval_from = None
        interval_to = None
        if report_date:
            interval_from, interval_to = _date_to_interval(report_date)

        ok_events = api.execute_report(
            group_id,
            ALL_VIOLATIONS_TEMPLATE_ID,
            events_path,
            interval_from=interval_from,
            interval_to=interval_to,
            table_index=0,
            sheet_name="Live Data",
        )
        time.sleep(1)
        ok_trips = api.execute_report(
            group_id,
            ALL_VIOLATIONS_TEMPLATE_ID,
            trips_path,
            interval_from=interval_from,
            interval_to=interval_to,
            table_index=1,
            sheet_name="Live Data",
        )
        time.sleep(1)

        night_fallback_path = os.path.join(raw_folder, f"{group_name}_NIGHT_TEMPLATE6_{timestamp}.xlsx")
        ok_night_fallback = api.execute_report(
            group_id,
            NIGHT_TEMPLATE_ID,
            night_fallback_path,
            interval_from=interval_from,
            interval_to=interval_to,
            table_index=0,
            sheet_name="Live Data",
        )

        if ok_events:
            downloaded.append({"type": "ALL_VIOLATIONS_EVENTS", "path": events_path, "template_id": "9/table0"})
        if ok_trips:
            downloaded.append({"type": "ALL_VIOLATIONS_TRIPS", "path": trips_path, "template_id": "9/table1"})
        if ok_night_fallback:
            downloaded.append({"type": "NIGHT_TEMPLATE6_RAW", "path": night_fallback_path, "template_id": "6/table0"})

        if ok_events:
            print("\n📦 Generating legacy raw files from Template 9...")
            downloaded.extend(
                _generate_derived_reports(
                    events_path,
                    trips_path if ok_trips else None,
                    raw_folder,
                    group_name,
                    timestamp,
                    night_fallback_path=night_fallback_path if ok_night_fallback else None,
                )
            )
        else:
            print("✗ Cannot generate derived files because Events pull failed")

    finally:
        api.logout()

    return downloaded, raw_folder


def print_summary(downloaded_files):
    print(f"\n{'='*60}")
    print("DOWNLOAD SUMMARY")
    print(f"{'='*60}")
    print(f"✓ Generated {len(downloaded_files)} files:\n")
    for file_info in downloaded_files:
        print(f"  • {file_info['type']}")
        print(f"    Template ID: {file_info['template_id']}")
        print(f"    Path: {file_info['path']}\n")


if __name__ == "__main__":
    import argparse

    print("\n" + "=" * 60)
    print("WIALON VIOLATION REPORTS PULLER (TEMPLATE 9 MODE)")
    print("=" * 60)

    parser = argparse.ArgumentParser(description="Pull Wialon violation reports and append to overall.")
    parser.add_argument("output_dir", nargs="?", default=None, help="Output folder (defaults to configured Downloads folder).")
    parser.add_argument("group", nargs="?", default=None, help="Wialon group name (defaults to TRANSIT_ALL_TRUCKS).")
    parser.add_argument("--date", dest="single_date", help="Pull for a specific date (YYYY-MM-DD or DD.MM.YYYY).")
    parser.add_argument("--from", dest="from_date", help="Start date (YYYY-MM-DD or DD.MM.YYYY).")
    parser.add_argument("--to", dest="to_date", help="End date (YYYY-MM-DD or DD.MM.YYYY).")
    parser.add_argument("--days-back", dest="days_back", type=int, help="Pull last N days ending yesterday.")
    args = parser.parse_args()

    today_tz = datetime.now(timezone(timedelta(hours=3))).date()
    dates_to_run = None

    try:
        if args.days_back:
            if args.days_back < 1:
                raise ValueError("--days-back must be >= 1")
            start = today_tz - timedelta(days=args.days_back)
            end = today_tz - timedelta(days=1)
            if start > end:
                raise ValueError("days-back range is empty")
            dates_to_run = list(_daterange(start, end))
        elif args.single_date:
            dates_to_run = [_parse_date_arg(args.single_date)]
        elif args.from_date or args.to_date:
            if not args.from_date or not args.to_date:
                raise ValueError("Both --from and --to are required when using a range.")
            start = _parse_date_arg(args.from_date)
            end = _parse_date_arg(args.to_date)
            if start > end:
                raise ValueError("--from date must be <= --to date")
            dates_to_run = list(_daterange(start, end))
    except Exception as e:
        print(f"ERROR: Date argument error: {e}")
        raise SystemExit(2)

    overall_folder = args.output_dir or DEFAULT_OUTPUT_FOLDER

    def _append_overall(raw_folder):
        try:
            overall_files = glob.glob(os.path.join(overall_folder, "OVERALL VIOLATIONS REPORT *.xlsx"))
            latest_overall = max(overall_files, key=os.path.getmtime) if overall_files else None

            if latest_overall:
                backup_folder = os.path.join(overall_folder, "backup")
                os.makedirs(backup_folder, exist_ok=True)
                shutil.copy(latest_overall, os.path.join(backup_folder, os.path.basename(latest_overall)))
                print(f"OK: Backed up latest overall to {backup_folder}")

                from append_to_overall import append_violations_to_overall
                append_success = append_violations_to_overall(raw_folder, overall_folder)
                if append_success:
                    print("\nOK: Data successfully appended to OVERALL excel!")
                else:
                    print("\nERROR: Failed to append to OVERALL excel (check errors above)")
                    print("  Raw reports are still available in:", raw_folder)
            else:
                print("WARN: No OVERALL file to append. Only raw reports are saved.")
        except Exception as e:
            print(f"\nERROR: Error during append operation: {e}")
            print("  Raw reports are still available in:", raw_folder)

    if dates_to_run:
        for idx, report_date in enumerate(dates_to_run, start=1):
            print(f"\n--- [{idx}/{len(dates_to_run)}] Date: {report_date.isoformat()} ---")
            files, raw_folder = pull_violation_reports(args.output_dir, args.group, report_date=report_date)
            print_summary(files)
            if files:
                print("OK: Reports pulled/generated successfully!")
                print("\n" + "=" * 60)
                print("STEP 2: APPENDING TO OVERALL VIOLATIONS REPORT")
                print("=" * 60)
                _append_overall(raw_folder)
            else:
                print("WARN: No reports were downloaded/generated.")
    else:
        files, raw_folder = pull_violation_reports(args.output_dir, args.group)
        print_summary(files)

        if files:
            print("OK: Reports pulled/generated successfully!")
            print("\n" + "=" * 60)
            print("STEP 2: APPENDING TO OVERALL VIOLATIONS REPORT")
            print("=" * 60)
            _append_overall(raw_folder)
        else:
            print("WARN: No reports were downloaded/generated.")

    print("=" * 60 + "\n")
