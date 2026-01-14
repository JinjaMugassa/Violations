"""Configuration constants for violation reports."""

# Wialon API Configuration
# These should be set in your .env file:
# WIALON_TOKEN=your_token_here
# WIALON_API_URL=https://hst-api.wialon.com/wialon/ajax.html

# Default resource ID for reports
WIALON_RESOURCE_ID = 22504459

# Default timezone offset for Tanzania (UTC+3)
TANZANIA_TIMEZONE_OFFSET = 10800  # seconds

# Target Groups
GROUPS = {
    "TRANSIT_ALL_TRUCKS": "TRANSIT_ALL_TRUCKS",
    "COPPER_LOADED_TRUCKS": "COPPER LOADED TRUCKS",
    "TRANSIT_GOING_TRIP": "TRANSIT- GOING TRIP",
    "TRANSIT_EMPTY_TRUCKS": "TRANSIT EMPTY TRUCKS",
    "LOCAL_MGS": "LOCAL _ MGS",
}

# Report Templates Configuration
TEMPLATES = {
    # Live Status Report (for night reports)
    "LIVE_STATUS": {
        "id": 8,
        "name": "Live Status Report"
    },
    
    # Speed Violation
    "SPEED_VIOLATION": {
        "id": 3,
        "name": "01_80 KPH_RPT_SPEED VIOLATION REPORT",
        "min_speed_threshold": 85  # km/h
    },
    
    # Harsh Brake
    "HARSH_BRAKE_SUMMARY": {
        "id": 89,
        "name": "RPT_HARSH BRAKE VIOLATIONS REPORT _GROUP WITH COUNT ONLY- TRANSIT"
    },
    "HARSH_BRAKE_DETAIL": {
        "id": 42,
        "name": "RPT_HARSH BRAKE VIOLATIONS REPORT _GROUP"
    },
    
    # Idling
    "IDLING": {
        "id": 11,
        "name": "01_RPT_IDLING VIOLATIONS REPORT (GROUP)",
        "exclude_counts": [1, 2]  # Omit rows with these counts
    },
    
    # Night Driving
    "NIGHT_DRIVING": {
        "id": 6,
        "name": "01_RPT_NIGHT DRIVING REPORT (GROUP)",
        "evening_window": {
            "start": {"hour": 20, "minute": 30},
            "end": {"hour": 23, "minute": 59}
        },
        "morning_window": {
            "start": {"hour": 4, "minute": 30},
            "end": {"hour": 5, "minute": 40}
        }
    }
}

# Excel Output Configuration
EXCEL_SHEET_NAME = "Live Data"
EXCEL_ENGINE = "openpyxl"

# Debug Configuration
SAVE_DEBUG_JSON = True  # Save intermediate JSON responses for debugging