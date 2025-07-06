import os
import shutil
import sys
import time
import platform
from datetime import datetime

from pathlib import Path

# Optional dependency – only available on Windows
try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # noqa: N816  # keep camelCase for clarity

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

###############################################################
# Configuration – EDIT THESE VALUES TO MATCH YOUR ENVIRONMENT #
###############################################################

WEB_URL = "https://example.com/your-report-page"  # URL to visit
PROFILE_NAME = "My Profile"  # Visible text of the profile to select
DATE_FROM = "2024-01-01"  # yyyy-mm-dd
DATE_TO = "2024-12-31"  # yyyy-mm-dd

# Where to store the file after initial download
DOWNLOAD_DIR = Path.home() / "Downloads"
# What to rename the download to *before* copying it to the network drive
NEW_FILE_NAME = "DailyReport.xlsx"

# Network/share drive paths (assumes already mounted / accessible)
NETWORK_DRIVE_FILE = Path("/mnt/shared/reports/DailyReport.xlsx")
MACRO_WORKBOOK_PATH = Path("/mnt/shared/reports/ProcessMacro.xlsm")
MACRO_NAME = "AutoProcess"  # Name of the macro sub to run inside the workbook

# SharePoint folder mounted locally (e.g. via rclone, OneDrive, or smb)
SHAREPOINT_DIR = Path("/mnt/sharepoint/Shared Documents/Monthly Reports")
FILES_TO_DEPLOY = [
    NETWORK_DRIVE_FILE.with_name("Result1.xlsx"),
    NETWORK_DRIVE_FILE.with_name("Result2.xlsx"),
]
###############################################################


class ReportAutomationError(Exception):
    """Custom exception for automation-specific failures."""


class ReportAutomation:
    def __init__(self):
        self.driver = None
        self._configure_driver()

    def _configure_driver(self) -> None:
        """Set up headless Chrome with automatic download directory."""
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        prefs = {
            "download.default_directory": str(DOWNLOAD_DIR),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.set_window_size(1920, 1080)

    # ------------------------------------------------------
    #                WEB INTERACTION METHODS
    # ------------------------------------------------------

    def download_excel_report(self) -> Path:
        """Automate the browser to download the Excel report.

        Returns the path to the downloaded file once download is complete.
        """
        assert self.driver is not None, "Driver not initialised"
        drv = self.driver
        wait = WebDriverWait(drv, 30)

        drv.get(WEB_URL)

        # Select profile – assumes it's a simple dropdown where you can click by visible text
        try:
            profile_dropdown = wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//select[./option[text()='{PROFILE_NAME}']]"))
            )
            profile_dropdown.click()
            option = profile_dropdown.find_element(By.XPATH, f".//option[normalize-space(text())='{PROFILE_NAME}']")
            option.click()
        except Exception as exc:
            raise ReportAutomationError("Unable to select profile") from exc

        # Input the date range – replace XPaths/selectors as appropriate
        try:
            from_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input#fromDate")))
            to_input = drv.find_element(By.CSS_SELECTOR, "input#toDate")
            from_input.clear()
            from_input.send_keys(DATE_FROM)
            to_input.clear()
            to_input.send_keys(DATE_TO)
        except Exception as exc:
            raise ReportAutomationError("Unable to set date range") from exc

        # Click export button
        try:
            export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]")))
            export_btn.click()
        except Exception as exc:
            raise ReportAutomationError("Export button not clickable") from exc

        # Wait for the .crdownload file to disappear signalling completion
        downloaded_file = self._wait_for_download_completion()
        return downloaded_file

    def _wait_for_download_completion(self, timeout: int = 120) -> Path:
        """Monitor the download directory for a new file and wait until it's fully downloaded."""
        before = {f for f in DOWNLOAD_DIR.iterdir()}
        start_time = time.time()
        while time.time() - start_time < timeout:
            current = {f for f in DOWNLOAD_DIR.iterdir()}
            new_files = current - before
            if new_files:
                newest = max(new_files, key=lambda p: p.stat().st_mtime)
                # Check if .crdownload exists (Chrome tmp extension)
                if newest.suffix == ".crdownload":
                    # Wait until it disappears
                    while newest.exists():
                        time.sleep(1)
                    # After completion, Chrome typically renames without .crdownload
                    # Fetch the actual file that matches the basename
                    completed = newest.with_suffix("")
                    if completed.exists():
                        return completed
                else:
                    # Non-crdownload – assume complete
                    return newest
            time.sleep(1)
        raise ReportAutomationError("Timed out waiting for download to complete")

    # ------------------------------------------------------
    #                       UTILITIES
    # ------------------------------------------------------

    @staticmethod
    def rename_file(src: Path, new_name: str) -> Path:
        dst = src.with_name(new_name)
        src.rename(dst)
        return dst

    @staticmethod
    def replace_file(src: Path, dst: Path) -> None:
        dst.parent.mkdir(parents=True, exist_ok=True)
        # Use os.replace to atomically replace if exists
        os.replace(src, dst)

    # ------------------------------------------------------
    #                    EXCEL MACRO RUNNER
    # ------------------------------------------------------

    def run_excel_macro(self, workbook_path: Path, macro_name: str) -> None:
        if platform.system().lower() != "windows":
            print("Warning: Running Excel macros is only supported on Windows with Excel installed.")
            return
        if win32com is None:
            raise ReportAutomationError(
                "pywin32 is not installed. Install it via 'pip install pywin32' on Windows."
            )
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        wb = excel.Workbooks.Open(str(workbook_path))
        try:
            excel.Application.Run(f"'{workbook_path.name}'!{macro_name}")
            time.sleep(2)  # Give macro some time – adjust as needed
        finally:
            # Close without saving as per requirement
            wb.Close(SaveChanges=False)
            excel.Quit()

    # ------------------------------------------------------
    #                      FILE DEPLOYMENT
    # ------------------------------------------------------

    @staticmethod
    def deploy_files_to_sharepoint(files: list[Path], target_dir: Path) -> None:
        target_dir.mkdir(parents=True, exist_ok=True)
        for src in files:
            dst = target_dir / src.name
            os.replace(src, dst)

    # ------------------------------------------------------
    #                           MAIN
    # ------------------------------------------------------

    def run(self) -> None:
        try:
            downloaded = self.download_excel_report()
            renamed = self.rename_file(downloaded, NEW_FILE_NAME)
            self.replace_file(renamed, NETWORK_DRIVE_FILE)
            self.run_excel_macro(MACRO_WORKBOOK_PATH, MACRO_NAME)
            self.deploy_files_to_sharepoint(FILES_TO_DEPLOY, SHAREPOINT_DIR)
            print("Automation completed successfully ✔")
        finally:
            if self.driver:
                self.driver.quit()


def main() -> None:
    automation = ReportAutomation()
    automation.run()


if __name__ == "__main__":
    main()