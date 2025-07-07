"""
Configuration file for the automation script
Modify the values below to match your specific requirements
"""

from datetime import datetime, timedelta
from pathlib import Path

# Web automation configuration
WEB_CONFIG = {
    'web_url': 'https://example.com',  # Replace with the actual web address
    'profile_selector': '#profile-dropdown',  # CSS selector for profile selection
    'date_range_selector': '#date-range-picker',  # CSS selector for date range
    'export_button_selector': '#export-btn',  # CSS selector for export button
    'start_date_selector': '#start-date',  # CSS selector for start date input
    'end_date_selector': '#end-date',  # CSS selector for end date input
    'username_field': '#username',  # If login is required
    'password_field': '#password',  # If login is required
    'login_button': '#login-btn',  # If login is required
    'download_timeout': 60,  # Seconds to wait for download
}

# File paths and naming
FILE_CONFIG = {
    'download_folder': str(Path.home() / 'Downloads'),
    'local_temp_folder': str(Path.home() / 'temp_automation'),
    'network_drive_path': '/path/to/network/drive',  # Replace with actual network drive path
    'sharepoint_folder': '/sites/yoursite/Shared Documents/YourFolder',  # SharePoint folder path
    'exported_file_pattern': 'export*.xlsx',  # Pattern to match downloaded files
    'target_file_name': 'processed_data.xlsx',  # Name for the processed file
    'macro_file_name': 'macro_processor.xlsm',  # Name of the macro file
    'output_files': [  # Files to copy to SharePoint after macro execution
        'output_file1.xlsx',
        'output_file2.xlsx',
    ],
}

# Date range configuration
DATE_CONFIG = {
    'start_date': (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d'),
    'end_date': datetime.now().strftime('%Y-%m-%d'),
    'date_format': '%Y-%m-%d',  # Format for date inputs
}

# SharePoint configuration
SHAREPOINT_CONFIG = {
    'sharepoint_url': 'https://yourcompany.sharepoint.com/sites/yoursite',
    'username': 'your_username@yourcompany.com',
    'password': 'your_password',  # Consider using environment variables for security
    'target_folder': '/sites/yoursite/Shared Documents/YourFolder',
    'replace_existing': True,  # Whether to replace existing files
}

# Excel macro configuration
EXCEL_CONFIG = {
    'macro_name': 'Main',  # Name of the macro to run
    'visible': False,  # Whether to show Excel during execution
    'save_before_close': False,  # Whether to save the file before closing
    'close_timeout': 30,  # Seconds to wait for macro completion
}

# Login credentials (if required for web automation)
LOGIN_CONFIG = {
    'username': 'your_web_username',
    'password': 'your_web_password',
    'use_login': False,  # Set to True if login is required
}

# Browser configuration
BROWSER_CONFIG = {
    'headless': False,  # Set to True to run browser in background
    'window_size': (1920, 1080),  # Browser window size
    'implicit_wait': 10,  # Seconds to wait for elements
    'page_load_timeout': 30,  # Seconds to wait for page load
}

# Logging configuration
LOGGING_CONFIG = {
    'log_file': 'automation.log',
    'log_level': 'INFO',  # DEBUG, INFO, WARNING, ERROR, CRITICAL
    'max_log_size': 10 * 1024 * 1024,  # 10MB
    'backup_count': 5,  # Number of log files to keep
}

# Email notification configuration (optional)
EMAIL_CONFIG = {
    'send_notifications': False,  # Set to True to enable email notifications
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'sender_email': 'your_email@gmail.com',
    'sender_password': 'your_app_password',
    'recipient_emails': ['recipient1@company.com', 'recipient2@company.com'],
    'subject_success': 'Automation Script - Success',
    'subject_failure': 'Automation Script - Failure',
}