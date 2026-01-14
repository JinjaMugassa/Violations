# Wialon Violation Reports System

A modular Python system for pulling and processing violation reports from Wialon tracking system.

## 📁 Project Structure

```
violations/
│
├── README.md                    # This file
├── requirements.txt             # Python dependencies
├── .env                        # Environment variables (not in repo)
│
├── config.py                   # Configuration constants
├── utils.py                    # Shared utility functions
├── wialon_api.py              # Wialon API client
├── run_pull_violation.py      # Main runner script
│
└── processors/                 # Report processors
    ├── __init__.py
    ├── speed_violation.py      # Speed violation processor
    ├── harsh_brake.py          # Harsh brake processor
    ├── idling.py               # Idling violations processor
    └── night_driving.py        # Night driving processor
```

## 🚀 Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment

Create a `.env` file in the project root:

```env
WIALON_TOKEN=your_wialon_api_token_here
WIALON_API_URL=https://hst-api.wialon.com/wialon/ajax.html
```

### 3. Run Reports

```bash
# Pull all violation reports for default group (TRANSIT_ALL_TRUCKS)
python run_pull_violation.py

# Specify custom output folder
python run_pull_violation.py /path/to/output

# Specify custom group
python run_pull_violation.py /path/to/output "YOUR_GROUP_NAME"
```

## 📊 Report Types

### 1. Speed Violation Report
- **Template ID**: 3
- **Processor**: `speed_violation.py`
- **Filters**: Only speeds ≥ 85 km/h
- **Output**: Single Excel file with filtered violations

### 2. Harsh Brake Violations
- **Template IDs**: 89 (summary), 42 (details)
- **Processor**: `harsh_brake.py`
- **Features**:
  - Merges summary counts with detailed event text
  - Includes speed information when available
  - Attempts Wialon API speed lookup if needed
  - Filters out Count = 1 or 2
- **Output**: Three files (summary, details, consolidated)

### 3. Idling Violations
- **Template ID**: 11
- **Processor**: `idling.py`
- **Features**:
  - Aggregates events per unit
  - Filters out Count = 1 or 2
  - One row per unit with total count
- **Output**: Single aggregated Excel file

### 4. Night Driving Report
- **Template ID**: 6
- **Processor**: `night_driving.py`
- **Features**:
  - Filters events within night windows:
    - Evening: 20:30 - 23:59
    - Morning: 04:30 - 05:40
  - Formats timestamps to Tanzania timezone (UTC+3)
- **Output**: Single filtered Excel file

## 🔧 Configuration

All configuration is centralized in `config.py`:

- **Wialon Resource ID**: 22504459
- **Timezone**: Africa/Dar_es_Salaam (UTC+3)
- **Template IDs**: All report template configurations
- **Processing Rules**: Speed thresholds, time windows, etc.

## 📝 Module Details

### `wialon_api.py`
Core Wialon API client handling:
- Authentication and session management
- Group and unit ID lookup
- Report execution and data fetching
- Speed value retrieval from message history

### `utils.py`
Shared utility functions:
- Timezone handling (Tanzania UTC+3)
- Timestamp formatting and conversion
- Column detection in DataFrames
- Speed extraction from text
- Debug JSON saving

### Processors (`processors/`)
Each processor implements:
- `process_*()` function that takes DataFrame, template_id, and API instance
- Report-specific filtering and transformation logic
- Data aggregation and formatting

## 🔍 Data Flow

```
1. run_pull_violation.py
   ↓
2. WialonAPI.login()
   ↓
3. WialonAPI.find_group_id(group_name)
   ↓
4. For each report type:
   ↓
5. WialonAPI.execute_report()
   ├── Fetch raw data from Wialon
   ├── Parse rows into DataFrame
   ├── Convert timestamps to Tanzania timezone
   ├── Apply processor function
   └── Save to Excel
   ↓
6. For harsh brake: merge_harsh_brake_reports()
   ├── Read summary and details
   ├── Match units between files
   ├── Enrich with speed data if available
   └── Create consolidated report
   ↓
7. Print summary and logout
```

## 🛠️ Adding New Report Types

To add a new report type:

1. **Create processor** in `processors/`:
   ```python
   # processors/new_report.py
   TEMPLATE_ID = your_template_id
   TEMPLATE_NAME = "Your Report Name"
   
   def process_new_report(df, template_id, api):
       if int(template_id) != TEMPLATE_ID:
           return df
       
       # Your processing logic here
       
       return df
   ```

2. **Update config.py**:
   ```python
   TEMPLATES["NEW_REPORT"] = {
       "id": your_template_id,
       "name": "Your Report Name"
   }
   ```

3. **Import in runner**:
   ```python
   from new_report import process_new_report, TEMPLATE_ID as NEW_TEMPLATE_ID
   ```

4. **Add to pull sequence** in `run_pull_violation.py`

## 📋 Output Files

All reports are saved as Excel files with:
- **Sheet name**: "Live Data"
- **Format**: `.xlsx` (OpenPyXL engine)
- **Naming**: `{GROUP}_{TYPE}_{TIMESTAMP}.xlsx`
- **Timestamp format**: `DD.MM.YYYY_HH-MM-SS`

Example filenames:
- `TRANSIT_ALL_TRUCKS_SPEED_VIOLATION_08.01.2026_14-30-45.xlsx`
- `TRANSIT_ALL_TRUCKS_HARSH_BRAKE_CONSOLIDATED_08.01.2026_14-31-02.xlsx`

## 🐛 Debugging

Debug JSON files are automatically saved alongside Excel outputs:
- `*_report_response.json` - Raw report execution response
- `*_rows_debug.json` - Raw rows data
- `*_rows_debug_X_Y.json` - Chunked rows (if applicable)

To disable debug files, set in `config.py`:
```python
SAVE_DEBUG_JSON = False
```

## ⚠️ Important Notes

1. **Timezone**: All timestamps are converted to Tanzania time (UTC+3)
2. **Rate Limiting**: 1-second delay between API calls
3. **Filtering**: Some reports filter out low-count violations (Count=1,2)
4. **Speed Data**: Harsh brake report attempts multiple methods to get speed:
   - Direct speed column
   - Text extraction from event fields
   - Wialon message history lookup (requires additional API call)

## 🔐 Security

- Never commit `.env` file to repository
- Keep Wialon token secure
- Rotate tokens periodically
- Use read-only tokens when possible

## 📄 License

[Your License Here]

## 👥 Contributors

[Your Name/Team]

## 📞 Support

For issues or questions, contact [your contact info]