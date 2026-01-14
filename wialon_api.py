"""Wialon API client for fetching reports and data."""

import os
import json
import time
import requests
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

WIALON_TOKEN = os.getenv("WIALON_TOKEN")
WIALON_API_URL = os.getenv("WIALON_API_URL")


def get_local_timezone_offset():
    """Returns Tanzania timezone offset in seconds."""
    return 10800  # 3 hours in seconds


def get_yesterday_interval():
    """Returns tuple of (from_timestamp, to_timestamp) for yesterday in Tanzania timezone."""
    from datetime import datetime, timezone, timedelta
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


def convert_timestamps_to_tanzania(df):
    """Convert datetime columns in DataFrame to Tanzania timezone."""
    import pytz
    try:
        tz_tz = pytz.timezone('Africa/Dar_es_Salaam')
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


def save_debug_json(data, output_path, suffix):
    """Save debug JSON file."""
    debug_path = output_path.replace(".xlsx", f"_{suffix}.json")
    try:
        with open(debug_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


class WialonAPI:
    """Wialon API client for authentication and data retrieval."""
    
    # Default resource ID for reports
    RESOURCE_ID = 22504459
    
    def __init__(self):
        self.token = WIALON_TOKEN
        self.api_url = WIALON_API_URL
        self.sid = None

    def login(self):
        """Login to Wialon API and establish session."""
        params = {
            "svc": "token/login",
            "params": json.dumps({"token": self.token})
        }
        resp = requests.post(self.api_url, params=params)
        data = resp.json()
        
        if "eid" in data:
            self.sid = data["eid"]
            print("✓ Logged in to Wialon successfully")
            return True
        
        print("✗ Login failed:", data)
        return False

    def logout(self):
        """Logout from Wialon API."""
        if self.sid:
            requests.post(
                self.api_url, 
                params={"svc": "core/logout", "sid": self.sid}
            )
            print("✓ Logged out from Wialon")

    def find_group_id(self, group_name):
        """Find a unit group ID by name."""
        params = {
            "svc": "core/search_items",
            "params": json.dumps({
                "spec": {
                    "itemsType": "avl_unit_group",
                    "propName": "sys_name",
                    "propValueMask": group_name,
                    "sortType": "sys_name",
                },
                "force": 1,
                "flags": 1,
                "from": 0,
                "to": 0,
            }),
            "sid": self.sid,
        }
        
        resp = requests.post(self.api_url, params=params)
        data = resp.json()
        items = data.get("items") or []
        
        if items:
            return items[0].get("id")
        return None

    def find_unit_id(self, unit_name):
        """Find a unit ID by its system name."""
        params = {
            "svc": "core/search_items",
            "params": json.dumps({
                "spec": {
                    "itemsType": "avl_unit",
                    "propName": "sys_name",
                    "propValueMask": unit_name,
                    "sortType": "sys_name",
                },
                "force": 1,
                "flags": 1,
                "from": 0,
                "to": 0,
            }),
            "sid": self.sid,
        }
        
        try:
            resp = requests.post(self.api_url, params=params, timeout=15)
            data = resp.json()
            items = data.get("items") or []
            if items:
                return items[0].get("id")
        except Exception:
            pass
        return None

    def get_unit_speed_at(self, unit_name, approx_ts=None):
        """Retrieve speed value for a unit at approximate timestamp.
        
        Args:
            unit_name: Name of the unit/vehicle
            approx_ts: Unix timestamp (searches ±2 minutes around this time)
            
        Returns:
            String like "52 km/h" or None if not found
        """
        if not self.sid:
            print(f"    ✗ No active session for speed lookup")
            return None
        
        try:
            # Find unit ID
            unit_id = self.find_unit_id(unit_name)
            if not unit_id:
                print(f"    ✗ Unit not found: {unit_name}")
                return None

            # Set time window (±2 minutes around event)
            now_ts = int(time.time())
            if approx_ts is None:
                to_ts = now_ts
                from_ts = max(0, to_ts - 3600)
            else:
                to_ts = int(approx_ts) + 120
                from_ts = max(0, int(approx_ts) - 120)

            # Load messages in time window
            params = {
                "svc": "messages/load_interval",
                "params": json.dumps({
                    "itemId": int(unit_id),
                    "timeFrom": int(from_ts),
                    "timeTo": int(to_ts),
                    "flags": 0x0000,  # Basic message data
                    "flagsMask": 0xFFFFFFFF
                }),
                "sid": self.sid,
            }
            
            resp = requests.post(self.api_url, params=params, timeout=30)
            data = resp.json()
            
            # Debug: save response
            # print(f"    DEBUG: Messages response keys: {data.keys()}")
            
            messages = data.get("messages", [])
            if not messages:
                print(f"    ⚠ No messages found for {unit_name} in time window")
                return None
            
            # Find message closest to target timestamp
            if approx_ts:
                messages_sorted = sorted(messages, key=lambda m: abs(m.get("t", 0) - approx_ts))
            else:
                messages_sorted = messages
            
            # Extract speed from messages
            for msg in messages_sorted:
                speed = None
                
                # Method 1: Direct speed field in message
                if "s" in msg:
                    speed = msg["s"]
                
                # Method 2: Speed in position data
                elif "pos" in msg and isinstance(msg["pos"], dict):
                    if "s" in msg["pos"]:
                        speed = msg["pos"]["s"]
                
                # Method 3: Speed in parameters
                elif "p" in msg and isinstance(msg["p"], dict):
                    if "speed" in msg["p"]:
                        speed = msg["p"]["speed"]
                    elif "s" in msg["p"]:
                        speed = msg["p"]["s"]
                
                # Validate and format speed
                if speed is not None:
                    try:
                        speed_val = float(speed)
                        if 0 <= speed_val < 400:  # Reasonable speed range
                            return f"{int(speed_val)} km/h"
                    except (ValueError, TypeError):
                        pass
            
            # If no direct speed field found, try text extraction as last resort
            import re
            for msg in messages_sorted[:5]:  # Check first 5 messages only
                msg_str = str(msg)
                
                # Look for "52 km/h" pattern
                match = re.search(r"(\d{1,3})\s*km[/\s]*h", msg_str, re.IGNORECASE)
                if match:
                    return f"{match.group(1)} km/h"
                
                # Look for "speed: 52" pattern
                match2 = re.search(r"speed\s*[:\-]?\s*(\d{1,3})", msg_str, re.IGNORECASE)
                if match2:
                    return f"{match2.group(1)} km/h"
            
            print(f"    ⚠ Found {len(messages)} messages but no speed data for {unit_name}")
            
        except Exception as e:
            print(f"    ✗ Speed lookup exception for {unit_name}: {e}")
            import traceback
            traceback.print_exc()
        
        return None

    def execute_report(self, group_id, template_id, output_path, 
                       interval_from=None, interval_to=None, 
                       processor_func=None):
        """Execute a Wialon report and save to Excel."""
        tz_offset = get_local_timezone_offset()
        
        if interval_from is None and interval_to is None:
            interval_from, interval_to = get_yesterday_interval()
        
        params = {
            "svc": "report/exec_report",
            "params": json.dumps({
                "reportResourceId": self.RESOURCE_ID,
                "reportTemplateId": template_id,
                "reportObjectId": group_id,
                "reportObjectSecId": 0,
                "interval": {
                    "from": int(interval_from),
                    "to": int(interval_to),
                    "flags": 0
                },
                "tzOffset": tz_offset,
            }),
            "sid": self.sid,
        }

        resp = requests.post(self.api_url, params=params, timeout=30)
        data = resp.json()
        
        if "reportResult" not in data:
            print("✗ Report execution failed:", data)
            return False

        save_debug_json(data, output_path, "report_response")

        report_result = data["reportResult"]
        tables = report_result.get("tables", [])
        
        if not tables:
            print("✗ No tables in report result")
            return False

        table = tables[0]
        headers = table.get("header", [])
        row_count = table.get("rows", 0)

        if row_count <= 0:
            print("✗ Report has zero rows")
            return False

        rows_list = self._fetch_report_rows(row_count, output_path)
        
        if not rows_list:
            print("✗ No rows extracted")
            return False

        parsed_rows = self._parse_rows(rows_list)
        
        if not parsed_rows:
            print("✗ No parsed rows extracted")
            return False

        try:
            df = pd.DataFrame(parsed_rows, columns=headers if headers else None)
        except Exception:
            df = pd.DataFrame(parsed_rows)

        df = convert_timestamps_to_tanzania(df)

        if processor_func:
            df = processor_func(df, template_id, self)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Live Data", index=False)

        print(f"✓ Report saved: {output_path} ({len(parsed_rows)} rows)")
        return True

    def _fetch_report_rows(self, row_count, output_path):
        """Fetch rows from executed report."""
        def fetch_rows(start, end):
            p = {
                "svc": "report/select_result_rows",
                "params": json.dumps({
                    "tableIndex": 0,
                    "config": {
                        "type": "range",
                        "data": {"from": start, "to": end, "level": 0}
                    }
                }),
                "sid": self.sid,
            }
            r = requests.post(self.api_url, params=p, timeout=60)
            try:
                return r.json()
            except Exception:
                return None

        rows_resp = fetch_rows(0, max(0, row_count - 1))
        rows_list = []

        save_debug_json(rows_resp, output_path, "rows_debug")

        if isinstance(rows_resp, dict) and rows_resp.get("error") is not None:
            chunk_size = 200
            for s in range(0, row_count, chunk_size):
                e = min(s + chunk_size - 1, row_count - 1)
                part = fetch_rows(s, e)
                save_debug_json(part, output_path, f"rows_debug_{s}_{e}")
                if isinstance(part, list):
                    rows_list.extend(part)
        elif isinstance(rows_resp, list):
            rows_list = rows_resp
        else:
            print("✗ Unexpected rows response type", type(rows_resp))

        return rows_list

    def _parse_rows(self, rows_list):
        """Parse raw rows into tabular format."""
        parsed_rows = []
        
        for row in rows_list:
            row_cells = []
            if isinstance(row, dict) and "c" in row:
                for cell in row["c"]:
                    if isinstance(cell, dict):
                        row_cells.append(cell.get("t", ""))
                    else:
                        row_cells.append(str(cell) if cell is not None else "")
            elif isinstance(row, list):
                for cell in row:
                    if isinstance(cell, dict):
                        row_cells.append(cell.get("t", ""))
                    else:
                        row_cells.append(str(cell) if cell is not None else "")
            else:
                continue
            parsed_rows.append(row_cells)

        return parsed_rows