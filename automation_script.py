#!/usr/bin/env python3
"""
Comprehensive Automation Script
Handles web automation, file operations, Excel macros, and SharePoint integration
"""

import os
import time
import shutil
import smtplib
from datetime import datetime, timedelta
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
from logging.handlers import RotatingFileHandler
from typing import Optional, List

# Web automation
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Excel operations
import xlwings as xw
# Alternative: import openpyxl

# SharePoint operations
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Import configurations
from config import (
    WEB_CONFIG, FILE_CONFIG, DATE_CONFIG, SHAREPOINT_CONFIG, 
    EXCEL_CONFIG, LOGIN_CONFIG, BROWSER_CONFIG, LOGGING_CONFIG, 
    EMAIL_CONFIG
)

class AutomationScript:
    def __init__(self):
        self.setup_logging()
        self.logger = logging.getLogger(__name__)
        self.driver = None
        
        # Create necessary directories
        os.makedirs(FILE_CONFIG['local_temp_folder'], exist_ok=True)
        os.makedirs(FILE_CONFIG['download_folder'], exist_ok=True)
        
        # Setup Chrome options
        self.chrome_options = Options()
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        self.chrome_options.add_argument('--disable-gpu')
        self.chrome_options.add_argument('--disable-extensions')
        self.chrome_options.add_argument('--disable-web-security')
        self.chrome_options.add_argument('--allow-running-insecure-content')
        
        if BROWSER_CONFIG['headless']:
            self.chrome_options.add_argument('--headless')
        
        # Set window size
        self.chrome_options.add_argument(f'--window-size={BROWSER_CONFIG["window_size"][0]},{BROWSER_CONFIG["window_size"][1]}')
        
        # Set download preferences
        prefs = {
            'download.default_directory': FILE_CONFIG['download_folder'],
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        }
        self.chrome_options.add_experimental_option('prefs', prefs)

    def setup_logging(self):
        """Setup logging configuration"""
        # Create logs directory if it doesn't exist
        os.makedirs('logs', exist_ok=True)
        
        # Setup rotating file handler
        log_file = os.path.join('logs', LOGGING_CONFIG['log_file'])
        handler = RotatingFileHandler(
            log_file,
            maxBytes=LOGGING_CONFIG['max_log_size'],
            backupCount=LOGGING_CONFIG['backup_count']
        )
        
        # Setup console handler
        console_handler = logging.StreamHandler()
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Setup logger
        logger = logging.getLogger()
        logger.setLevel(getattr(logging, LOGGING_CONFIG['log_level']))
        logger.addHandler(handler)
        logger.addHandler(console_handler)

    def send_email_notification(self, subject: str, body: str, is_success: bool = True):
        """Send email notification"""
        if not EMAIL_CONFIG['send_notifications']:
            return
            
        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_CONFIG['sender_email']
            msg['To'] = ', '.join(EMAIL_CONFIG['recipient_emails'])
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Connect to server and send email
            server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
            server.starttls()
            server.login(EMAIL_CONFIG['sender_email'], EMAIL_CONFIG['sender_password'])
            server.send_message(msg)
            server.quit()
            
            self.logger.info("Email notification sent successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to send email notification: {str(e)}")

    def setup_webdriver(self) -> webdriver.Chrome:
        """Initialize and return Chrome WebDriver"""
        self.logger.info("Setting up WebDriver...")
        try:
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.implicitly_wait(BROWSER_CONFIG['implicit_wait'])
            driver.set_page_load_timeout(BROWSER_CONFIG['page_load_timeout'])
            return driver
        except Exception as e:
            self.logger.error(f"Failed to setup WebDriver: {str(e)}")
            raise

    def perform_login(self, driver: webdriver.Chrome):
        """Perform login if required"""
        if not LOGIN_CONFIG['use_login']:
            return
            
        try:
            self.logger.info("Performing login...")
            
            # Find and fill username
            username_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, WEB_CONFIG['username_field']))
            )
            username_field.clear()
            username_field.send_keys(LOGIN_CONFIG['username'])
            
            # Find and fill password
            password_field = driver.find_element(By.CSS_SELECTOR, WEB_CONFIG['password_field'])
            password_field.clear()
            password_field.send_keys(LOGIN_CONFIG['password'])
            
            # Click login button
            login_button = driver.find_element(By.CSS_SELECTOR, WEB_CONFIG['login_button'])
            login_button.click()
            
            # Wait for login to complete
            WebDriverWait(driver, 10).until(
                EC.url_changes(driver.current_url)
            )
            
            self.logger.info("Login completed successfully")
            
        except Exception as e:
            self.logger.error(f"Login failed: {str(e)}")
            raise

    def web_automation_and_export(self, driver: webdriver.Chrome) -> str:
        """
        Perform web automation: navigate, select profile, date range, and export
        Returns the path to the downloaded file
        """
        self.logger.info("Starting web automation...")
        
        try:
            # Navigate to the web address
            self.logger.info(f"Navigating to {WEB_CONFIG['web_url']}")
            driver.get(WEB_CONFIG['web_url'])
            
            # Wait for page to load
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # Perform login if required
            self.perform_login(driver)
            
            # Select profile
            if WEB_CONFIG['profile_selector']:
                self.logger.info("Selecting profile...")
                profile_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, WEB_CONFIG['profile_selector']))
                )
                profile_element.click()
                time.sleep(1)  # Wait for profile to load
            
            # Select date range
            if WEB_CONFIG['date_range_selector']:
                self.logger.info("Setting date range...")
                date_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, WEB_CONFIG['date_range_selector']))
                )
                date_element.click()
                
                # Set start date
                if WEB_CONFIG['start_date_selector']:
                    start_date_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, WEB_CONFIG['start_date_selector']))
                    )
                    start_date_input.clear()
                    start_date_input.send_keys(DATE_CONFIG['start_date'])
                
                # Set end date
                if WEB_CONFIG['end_date_selector']:
                    end_date_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, WEB_CONFIG['end_date_selector']))
                    )
                    end_date_input.clear()
                    end_date_input.send_keys(DATE_CONFIG['end_date'])
                
                time.sleep(1)  # Wait for date range to be processed
            
            # Click export button
            self.logger.info("Clicking export button...")
            export_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, WEB_CONFIG['export_button_selector']))
            )
            export_button.click()
            
            # Wait for download to complete
            self.logger.info("Waiting for download to complete...")
            downloaded_file = self.wait_for_download()
            
            return downloaded_file
            
        except TimeoutException as e:
            self.logger.error(f"Timeout during web automation: {str(e)}")
            raise
        except NoSuchElementException as e:
            self.logger.error(f"Element not found during web automation: {str(e)}")
            raise
        except Exception as e:
            self.logger.error(f"Error during web automation: {str(e)}")
            raise

    def wait_for_download(self, timeout: int = None) -> str:
        """Wait for file download to complete and return the file path"""
        if timeout is None:
            timeout = WEB_CONFIG['download_timeout']
            
        download_folder = Path(FILE_CONFIG['download_folder'])
        
        # Get initial file list
        initial_files = set(download_folder.glob(FILE_CONFIG['exported_file_pattern']))
        
        # Wait for new file
        start_time = time.time()
        while time.time() - start_time < timeout:
            current_files = set(download_folder.glob(FILE_CONFIG['exported_file_pattern']))
            new_files = current_files - initial_files
            
            if new_files:
                # Get the newest file
                newest_file = max(new_files, key=lambda f: f.stat().st_mtime)
                
                # Check if file is fully downloaded (not being written to)
                if self.is_file_ready(newest_file):
                    self.logger.info(f"Downloaded file: {newest_file}")
                    return str(newest_file)
            
            time.sleep(1)
        
        raise TimeoutError(f"Download timeout after {timeout} seconds - file not found")

    def is_file_ready(self, file_path: Path) -> bool:
        """Check if file is fully downloaded and ready to use"""
        try:
            # Check if file size is stable
            size1 = file_path.stat().st_size
            time.sleep(0.5)
            size2 = file_path.stat().st_size
            
            # File is ready if size is stable and greater than 0
            return size1 == size2 and size1 > 0
        except:
            return False

    def rename_and_copy_to_network(self, source_file: str) -> str:
        """Rename file and copy to network drive"""
        self.logger.info("Renaming and copying file to network drive...")
        
        try:
            # Ensure network drive path exists
            network_path = Path(FILE_CONFIG['network_drive_path'])
            network_path.mkdir(parents=True, exist_ok=True)
            
            # Create target path
            target_path = network_path / FILE_CONFIG['target_file_name']
            
            # Copy and rename
            shutil.copy2(source_file, target_path)
            self.logger.info(f"File copied to: {target_path}")
            
            return str(target_path)
            
        except Exception as e:
            self.logger.error(f"Error copying file to network drive: {str(e)}")
            raise

    def run_excel_macro(self, macro_file_path: str):
        """Open Excel file and run macro"""
        self.logger.info(f"Running Excel macro from: {macro_file_path}")
        
        app = None
        try:
            # Check if macro file exists
            if not Path(macro_file_path).exists():
                raise FileNotFoundError(f"Macro file not found: {macro_file_path}")
            
            # Start Excel application
            app = xw.App(visible=EXCEL_CONFIG['visible'])
            
            # Open the macro file
            wb = app.books.open(macro_file_path)
            
            # Run the macro
            macro_name = EXCEL_CONFIG['macro_name']
            self.logger.info(f"Executing macro: {macro_name}")
            wb.macro(macro_name)()
            
            # Close without saving if specified
            if not EXCEL_CONFIG['save_before_close']:
                wb.close()
            else:
                wb.save()
                wb.close()
            
            app.quit()
            self.logger.info("Macro executed successfully")
            
        except Exception as e:
            self.logger.error(f"Error running macro: {str(e)}")
            # Ensure Excel is closed
            if app:
                try:
                    app.quit()
                except:
                    pass
            raise

    def copy_files_to_sharepoint(self, files_to_copy: List[str]):
        """Copy files to SharePoint folder"""
        self.logger.info("Copying files to SharePoint...")
        
        try:
            # Authenticate with SharePoint
            ctx_auth = AuthenticationContext(SHAREPOINT_CONFIG['sharepoint_url'])
            ctx_auth.acquire_token_for_user(
                SHAREPOINT_CONFIG['username'],
                SHAREPOINT_CONFIG['password']
            )
            
            # Create client context
            ctx = ClientContext(SHAREPOINT_CONFIG['sharepoint_url'], ctx_auth)
            
            # Upload files
            for local_file_path in files_to_copy:
                if not Path(local_file_path).exists():
                    self.logger.warning(f"File not found, skipping: {local_file_path}")
                    continue
                    
                file_name = Path(local_file_path).name
                self.logger.info(f"Uploading {file_name} to SharePoint...")
                
                # Read file content
                with open(local_file_path, 'rb') as file_content:
                    file_data = file_content.read()
                
                # Upload to SharePoint
                target_folder = ctx.web.get_folder_by_server_relative_url(
                    SHAREPOINT_CONFIG['target_folder']
                )
                
                # Check if file exists and handle replacement
                if SHAREPOINT_CONFIG['replace_existing']:
                    try:
                        # Try to delete existing file first
                        existing_file = target_folder.get_file(file_name)
                        existing_file.delete_object()
                        ctx.execute_query()
                        self.logger.info(f"Replaced existing file: {file_name}")
                    except:
                        # File doesn't exist, which is fine
                        pass
                
                # Upload the file
                target_file = target_folder.upload_file(file_name, file_data)
                ctx.execute_query()
                self.logger.info(f"Successfully uploaded {file_name} to SharePoint")
                
        except Exception as e:
            self.logger.error(f"Error copying to SharePoint: {str(e)}")
            raise

    def cleanup_temp_files(self):
        """Clean up temporary files"""
        self.logger.info("Cleaning up temporary files...")
        try:
            temp_folder = Path(FILE_CONFIG['local_temp_folder'])
            if temp_folder.exists():
                shutil.rmtree(temp_folder)
                self.logger.info("Temporary files cleaned up")
        except Exception as e:
            self.logger.warning(f"Error cleaning up temp files: {str(e)}")

    def run_full_automation(self):
        """Run the complete automation process"""
        self.logger.info("Starting full automation process...")
        start_time = datetime.now()
        
        try:
            # Step 1: Web automation and export
            self.logger.info("Step 1: Web automation and file export")
            self.driver = self.setup_webdriver()
            downloaded_file = self.web_automation_and_export(self.driver)
            
            # Step 2: Rename and copy to network drive
            self.logger.info("Step 2: Copying file to network drive")
            network_file = self.rename_and_copy_to_network(downloaded_file)
            
            # Step 3: Run Excel macro
            self.logger.info("Step 3: Running Excel macro")
            macro_file_path = Path(FILE_CONFIG['network_drive_path']) / FILE_CONFIG['macro_file_name']
            self.run_excel_macro(str(macro_file_path))
            
            # Step 4: Copy files to SharePoint
            self.logger.info("Step 4: Copying files to SharePoint")
            files_to_copy = [network_file]
            
            # Add output files from macro execution
            for output_file in FILE_CONFIG['output_files']:
                file_path = Path(FILE_CONFIG['network_drive_path']) / output_file
                if file_path.exists():
                    files_to_copy.append(str(file_path))
                else:
                    self.logger.warning(f"Output file not found: {file_path}")
            
            self.copy_files_to_sharepoint(files_to_copy)
            
            # Success notification
            end_time = datetime.now()
            duration = end_time - start_time
            
            success_message = f"""
            Automation completed successfully!
            
            Duration: {duration}
            Files processed: {len(files_to_copy)}
            
            Steps completed:
            1. Web automation and file export ✓
            2. File copied to network drive ✓
            3. Excel macro executed ✓
            4. Files uploaded to SharePoint ✓
            
            Log file: {os.path.join('logs', LOGGING_CONFIG['log_file'])}
            """
            
            self.logger.info(success_message)
            self.send_email_notification(
                EMAIL_CONFIG['subject_success'],
                success_message,
                is_success=True
            )
            
        except Exception as e:
            # Failure notification
            end_time = datetime.now()
            duration = end_time - start_time
            
            error_message = f"""
            Automation failed!
            
            Duration: {duration}
            Error: {str(e)}
            
            Please check the log file for more details: {os.path.join('logs', LOGGING_CONFIG['log_file'])}
            """
            
            self.logger.error(error_message)
            self.send_email_notification(
                EMAIL_CONFIG['subject_failure'],
                error_message,
                is_success=False
            )
            raise
        
        finally:
            # Cleanup
            if self.driver:
                self.driver.quit()
            self.cleanup_temp_files()

if __name__ == "__main__":
    # Create and run automation
    automation = AutomationScript()
    automation.run_full_automation()