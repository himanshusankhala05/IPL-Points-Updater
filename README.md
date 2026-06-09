# IPL Fantasy League Management System

An end-to-end IPL Fantasy League platform built using **UiPath**, **Python**, **Excel**, and **Google Sheets** to manage player auctions, scrape live IPL statistics, calculate fantasy points, track team performance, and maintain league analytics.

---

## Project Overview

This project was created to automate and manage a private IPL Fantasy League.

The complete workflow consists of:

1. **IPL Auction System**

   * Conduct player auctions.
   * Allocate players to team owners.
   * Manage team budgets and squad formation.

2. **UiPath Data Scraping**

   * Collect player statistics from official IPL websites.
   * Extract Fantasy IPL points.
   * Export data into Excel.

3. **Python Automation**

   * Process player statistics.
   * Calculate points and dot-ball bonuses.
   * Update Google Sheets dashboards.

4. **League Analytics**

   * Track team standings.
   * Monitor daily points gained.
   * Maintain historical performance records.
   * Generate Top 12 team statistics.

---

## Related Repository

### IPL Auction System

Repository used to conduct the fantasy league auction and allocate players to teams:

**GitHub Repository**

`https://github.com/himanshusankhala05/IPL_Auction2025`

### Auction Features

* Player bidding system
* Team purse management
* Player allocation
* Squad creation
* Auction history tracking
* Team roster generation

The output of the auction system becomes the foundation for this fantasy league management platform.

---

## Data Sources

### Official IPL Statistics

https://www.iplt20.com/stats/2026

Used for:

* Batting statistics
* Bowling statistics
* Match performance data
* Dot-ball statistics
* Player performance metrics

### Fantasy IPL

https://fantasy.iplt20.com

Used for:

* Fantasy points
* Player rankings
* Fantasy scoring information

---

## System Architecture

```text
                 IPL Auction System
                          │
                          ▼
                Team & Player Allocation
                          │
                          ▼
      IPL Website + Fantasy IPL Website
                          │
                          ▼
                  UiPath Automation
                          │
                          ▼
                    Excel Dataset
                          │
                          ▼
                  Python Processing
                          │
                          ▼
                 Google Sheets Database
                          │
                          ▼
          League Dashboard & Analytics
```

---

## Features

### Auction Management

* Fantasy player auction
* Team creation
* Budget management
* Squad tracking

### Data Scraping (UiPath)

* Automated IPL statistics extraction
* Fantasy IPL points scraping
* Structured Excel exports
* Reduced manual effort

### Fantasy Points Processing

* Player points updates
* Dot-ball bonus tracking
* Previous points comparison
* Daily score calculations
* Automatic data synchronization

### Team Analytics

* Team rankings
* Daily points gained
* Historical score tracking
* Top 12 player/team analysis
* League performance monitoring

### Dashboard Management

* Google Sheets integration
* Automated updates
* Historical reporting
* League summaries

---

## Technology Stack

### UiPath

Used for:

* Web scraping
* Browser automation
* Data extraction

### Python

Libraries:

```bash
pip install pandas gspread google-auth openpyxl
```

Main Packages:

* pandas
* gspread
* google-auth
* openpyxl
* logging
* datetime

### Google Sheets API

Used for:

* Data storage
* League dashboard
* Analytics reporting

---

## Project Structure

```text
.
├── UiPath/
│   └── IPL.xaml
│
├── Python/
│   └── pythonScript_curr.py
│
├── Data/
│   ├── IPLPointsData.xlsx
│   ├── service_account.json
│   └── teams.txt
│
├── Logs/
│   └── ipl_updater_*.log
│
└── README.md
```

---

## Google Sheets Structure

### Points Sheet

Stores:

* Player Name
* Fantasy Points
* Dot Balls
* Previous Points
* Updated Points

### Home Sheet

Displays:

* Team rankings
* Daily points gained
* Top 12 scores
* League summary

### Data Sheet

Stores historical records:

* Daily points
* Team performance trends
* Historical rankings
* Top 12 statistics

---

## Workflow

### Step 1: Auction

Players are allocated to fantasy teams through the IPL Auction System.

### Step 2: Data Collection

UiPath scrapes:

* IPL statistics
* Fantasy IPL points

and exports them into Excel.

### Step 3: Data Processing

Python:

* Reads Excel files
* Validates data
* Calculates updates
* Identifies changes

### Step 4: Dashboard Update

Google Sheets is automatically updated with:

* Player points
* Dot-ball bonuses
* Team standings
* Daily performance metrics

### Step 5: Analytics

League dashboards provide:

* Team rankings
* Daily gains
* Historical trends
* Top 12 performance tracking

---

## Running the Project

### Install Dependencies

```bash
pip install pandas gspread google-auth openpyxl
```

### Configure

Update:

```python
EXCEL_FILE
SERVICE_ACCOUNT
SHEET_NAME
LOG_FILE_PATH
```

inside the Python script.

### Execute

Run UiPath workflow to collect data.

Then execute:

```bash
python pythonScript_curr.py
```

---

## Logging

The system generates timestamped log files.

Example:

```text
ipl_updater_20260509_214530.log
```

Logs include:

* Authentication status
* Sheet updates
* Player updates
* Errors and warnings
* Execution summaries

---

## Future Enhancements

* Fully automated scheduling
* Live score integration
* Telegram/WhatsApp notifications
* Power BI dashboards
* Web-based fantasy portal
* Automated league reports

---

## Author

Himanshu Sankhala

Passionate about automation, data analytics, fantasy sports management, and process optimization using UiPath and Python.

---

## License

This project is intended for educational and personal use.
