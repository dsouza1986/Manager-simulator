# Automation Script

This script automates the daily workflow of downloading a report from a web portal, copying it to a network drive, running an Excel macro, and finally publishing the resulting files to a SharePoint location.

> **Important:** All element selectors, URLs, and file paths are placeholders. You must update them to match your environment before the script will work.

## Prerequisites

1. **Python 3.10+**
2. **Google Chrome / Chromium** installed, plus the matching version of **ChromeDriver** available on your `$PATH`.
3. **Python packages** (create your virtual environment and run):

```bash
pip install -r requirements.txt
```

4. **Windows-only:** If you need to execute the Excel macro step, you must run the script on Windows with Microsoft Excel installed. Uncomment the `pywin32` line in `requirements.txt` and install it.

5. Ensure that the network drive and SharePoint folders are mounted locally (e.g. `/mnt/shared` and `/mnt/sharepoint`). The script simply treats them as regular file paths.

## Configuration

Open `automation.py` and adjust the **Configuration** section at the top:

```python
WEB_URL = "https://example.com/your-report-page"
PROFILE_NAME = "My Profile"
DATE_FROM = "2024-01-01"
DATE_TO = "2024-12-31"
DOWNLOAD_DIR = Path.home() / "Downloads"
NEW_FILE_NAME = "DailyReport.xlsx"
NETWORK_DRIVE_FILE = Path("/mnt/shared/reports/DailyReport.xlsx")
MACRO_WORKBOOK_PATH = Path("/mnt/shared/reports/ProcessMacro.xlsm")
MACRO_NAME = "AutoProcess"
SHAREPOINT_DIR = Path("/mnt/sharepoint/Shared Documents/Monthly Reports")
FILES_TO_DEPLOY = [
    NETWORK_DRIVE_FILE.with_name("Result1.xlsx"),
    NETWORK_DRIVE_FILE.with_name("Result2.xlsx"),
]
```

## Running the script

```bash
python automation.py
```

When run headless, Chrome will not pop up a visible window. Check the console output for a success ✅ message.

## Customising selectors

Depending on the target web page, you may need to replace the XPaths and CSS selectors in `download_excel_report()` so Selenium can locate the correct elements.

## Error handling

All major steps raise `ReportAutomationError` when something goes wrong. The script will also print warnings if you attempt to run the Excel macro step on a non-Windows system.