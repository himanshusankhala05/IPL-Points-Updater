# IPL Fantasy Points Updater

A Python automation script that synchronizes IPL fantasy cricket player statistics from Excel files into Google Sheets. The tool updates player points, dot-ball statistics, team performance metrics, daily points gained, and Top-12 team scores while maintaining detailed execution logs.

## Features

* 📊 Read player statistics from Excel spreadsheets
* ☁️ Authenticate and connect to Google Sheets using a Service Account
* 🔄 Update player fantasy points automatically
* 🎯 Update bowler dot-ball statistics
* 📈 Track daily points gained for each team
* 🏆 Update Top 12 team scores
* 📝 Detailed logging with timestamped log files
* ⚡ Batch updates for improved Google Sheets performance
* 🔍 Detect unchanged records and skip unnecessary updates
* 🚨 Track and report players that were not updated

---

## Project Structure

```text
.
├── IPLPointsData.xlsx          # Source player data
├── service_account.json        # Google API credentials
├── teams.txt                   # Team information
├── logs/
│   └── ipl_updater_*.log       # Generated log files
└── pythonScript_curr.py        # Main script
```

---

## Prerequisites

### Python Version

* Python 3.10+

### Required Packages

Install dependencies:

```bash
pip install pandas gspread google-auth openpyxl
```

---

## Google Sheets Setup

### 1. Create a Google Cloud Project

1. Open Google Cloud Console.
2. Create a new project.
3. Enable:

   * Google Sheets API
   * Google Drive API

### 2. Create a Service Account

1. Navigate to **IAM & Admin → Service Accounts**
2. Create a Service Account
3. Generate a JSON Key
4. Download the JSON file
5. Save it as:

```text
service_account.json
```

### 3. Share the Spreadsheet

Share your Google Spreadsheet with the Service Account email address and grant **Editor** access.

---

## Configuration

Update the following constants in the script:

```python
EXCEL_FILE = "path_to_IPLPointsData.xlsx"
SERVICE_ACCOUNT = "path_to_service_account.json"
SHEET_NAME = "YourGoogleSpreadsheetName"
LOG_FILE_PATH = "path_to_log_file.log"
```

---

## Google Sheet Requirements

The script expects the following worksheets:

### Points

Stores player data and fantasy points.

### Home

Stores team summary data:

| Range | Purpose             |
| ----- | ------------------- |
| F2:G6 | Top 12 Team Points  |
| K2:L6 | Daily Points Gained |

### Data

Stores historical tracking data:

* Daily team points
* Top 12 team scores
* Date-wise analytics

---

## Excel Requirements

### Sheet1

Contains player information.

Example:

| Name     | Team Name | Points | Dot Balls |
| -------- | --------- | ------ | --------- |
| Player A | Team X    | 120    | 10        |

### Sheet2

Contains player fantasy points and dot-ball data used during updates.

---

## Workflow

### Step 1

Load player data from Excel.

### Step 2

Authenticate with Google Sheets.

### Step 3

Open required worksheets:

* Points
* Home
* Data

### Step 4

(Optional) Update previous player points.

### Step 5

Process player points and dot-ball updates.

### Step 6

Update:

* Daily points gained
* Top 12 team points

### Step 7

Generate execution logs.

---

## Running the Script

```bash
python pythonScript_curr.py
```

The script provides interactive prompts:

```text
Do you want to update previous points? (y/n)

Do you want to update past top 12 points? (y/n)

Do you want to process player points and dot balls updates? (y/n)

Do you want to update points gained today and top 12 points? (y/n)
```

---

## Logging

Logs are automatically generated with timestamps.

Example:

```text
ipl_updater_20250415_213012.log
```

Log entries include:

* Authentication status
* Sheet updates
* Player processing
* Errors and exceptions
* Summary statistics

---

## Error Handling

The script handles:

* Missing worksheets
* Authentication failures
* Invalid data formats
* Missing players
* Missing date columns
* Google Sheets API issues

All exceptions are logged for troubleshooting.

---

## Team Mapping

| Full Team Name    | Abbreviation |
| ----------------- | ------------ |
| Himanshu Warriors | HW           |
| Mehul Challengers | MC           |
| Onkar Legends     | OL           |
| Abhishek Strikers | AS           |
| Pranav Astra      | PA           |

---

## Future Improvements

* Configuration via `.env` file
* Automatic date management
* Retry mechanism for Google API limits
* Command-line arguments
* Email notifications
* Automated scheduling using Task Scheduler or Cron

---

## Author

Developed for IPL Fantasy League data management and Google Sheets automation.

---

## License

This project is available for personal and educational use.
