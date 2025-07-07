# Web Automation & SharePoint Integration Script

A comprehensive automation program that performs web browser automation, Excel file processing, macro execution, and SharePoint integration.

## Features

- **Web Automation**: Automated browser interactions using Selenium
- **File Processing**: Download, rename, and copy files between locations
- **Excel Integration**: Run macros and process Excel files
- **SharePoint Integration**: Upload and manage files in SharePoint
- **Email Notifications**: Get notified of success/failure
- **Robust Error Handling**: Comprehensive logging and error recovery
- **Cross-Platform Support**: Works on Windows, macOS, and Linux

## Workflow

The automation performs the following steps:

1. **Web Automation**: Opens browser, navigates to specified URL, selects profile and date range
2. **File Export**: Downloads Excel file from web application
3. **File Management**: Renames and copies file to network drive
4. **Macro Execution**: Runs Excel macro on the processed file
5. **SharePoint Upload**: Uploads processed files to SharePoint folder
6. **Cleanup**: Removes temporary files and closes applications

## Prerequisites

- Python 3.8 or higher
- Google Chrome browser
- Microsoft Excel (for macro execution)
- Access to SharePoint (with valid credentials)
- Network drive access (if applicable)

## Installation

### Option 1: Automatic Setup (Recommended)

1. Clone or download this repository
2. Run the setup script:
   ```bash
   python setup.py
   ```

### Option 2: Manual Setup

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Install Chrome WebDriver:
   ```bash
   pip install webdriver-manager
   ```

3. Create necessary directories:
   ```bash
   mkdir logs downloads temp config
   ```

## Configuration

1. Copy the sample configuration file:
   ```bash
   cp config_sample.py config.py
   ```

2. Edit `config.py` with your specific settings:

### Web Configuration
```python
WEB_CONFIG = {
    'web_url': 'https://your-website.com',
    'profile_selector': '#profile-dropdown',
    'date_range_selector': '#date-range-picker',
    'export_button_selector': '#export-button',
    'start_date_selector': '#start-date',
    'end_date_selector': '#end-date',
    'download_timeout': 60,
}
```

### File Configuration
```python
FILE_CONFIG = {
    'download_folder': str(Path.home() / 'Downloads'),
    'network_drive_path': '/path/to/your/network/drive',
    'sharepoint_folder': '/sites/yoursite/Documents/YourFolder',
    'target_file_name': 'processed_data.xlsx',
    'macro_file_name': 'your_macro.xlsm',
    'output_files': ['output1.xlsx', 'output2.xlsx'],
}
```

### SharePoint Configuration
```python
SHAREPOINT_CONFIG = {
    'sharepoint_url': 'https://yourcompany.sharepoint.com/sites/yoursite',
    'username': 'your_username@yourcompany.com',
    'password': 'your_password',
    'target_folder': '/sites/yoursite/Documents/YourFolder',
    'replace_existing': True,
}
```

## Usage

### Basic Usage
```bash
python automation_script.py
```

### Running with Virtual Environment
```bash
python -m venv automation_env
source automation_env/bin/activate  # On Windows: automation_env\Scripts\activate
pip install -r requirements.txt
python automation_script.py
```

### Scheduling (Optional)
You can schedule the automation to run automatically:

#### Windows (Task Scheduler)
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (daily, weekly, etc.)
4. Set action to run `python automation_script.py`

#### Linux/macOS (Cron)
```bash
# Edit crontab
crontab -e

# Run every day at 9 AM
0 9 * * * /usr/bin/python3 /path/to/automation_script.py
```

## Configuration Details

### CSS Selectors
You'll need to identify the correct CSS selectors for your web application. Use browser developer tools:

1. Right-click on the element → Inspect
2. Right-click on the HTML element → Copy → Copy selector
3. Use the copied selector in your configuration

### Excel Macro Setup
1. Ensure your Excel macro file is in the network drive location
2. The macro should be named according to `EXCEL_CONFIG['macro_name']`
3. Test the macro manually first to ensure it works

### SharePoint Authentication
- Use your organization's SharePoint credentials
- For enhanced security, consider using environment variables:
  ```python
  import os
  SHAREPOINT_CONFIG = {
      'username': os.getenv('SHAREPOINT_USERNAME'),
      'password': os.getenv('SHAREPOINT_PASSWORD'),
  }
  ```

## Logging

The script creates detailed logs in the `logs/` directory:
- `automation.log` - Main log file
- Log rotation keeps 5 backup files
- Configurable log levels (DEBUG, INFO, WARNING, ERROR)

## Email Notifications

Configure email notifications for success/failure:

```python
EMAIL_CONFIG = {
    'send_notifications': True,
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'sender_email': 'your_email@gmail.com',
    'sender_password': 'your_app_password',
    'recipient_emails': ['recipient@company.com'],
}
```

For Gmail, use an App Password instead of your regular password.

## Troubleshooting

### Common Issues

#### 1. Chrome WebDriver Issues
```bash
# Update Chrome WebDriver
pip install --upgrade webdriver-manager
```

#### 2. Excel/Macro Issues
- Ensure Excel is installed and accessible
- Check macro security settings
- Verify macro file path and name

#### 3. SharePoint Connection Issues
- Verify URL format: `https://company.sharepoint.com/sites/sitename`
- Check folder path format: `/sites/sitename/Documents/FolderName`
- Ensure credentials are correct

#### 4. File Permission Issues
- Check network drive permissions
- Ensure write access to download folder
- Verify SharePoint upload permissions

### Debug Mode

Enable debug logging:
```python
LOGGING_CONFIG = {
    'log_level': 'DEBUG',
}
```

Run with visible browser for troubleshooting:
```python
BROWSER_CONFIG = {
    'headless': False,
}
```

## File Structure

```
automation_project/
├── automation_script.py      # Main automation script
├── config.py                 # Configuration file (create from sample)
├── config_sample.py          # Sample configuration
├── requirements.txt          # Python dependencies
├── setup.py                  # Setup script
├── README.md                 # This file
├── logs/                     # Log files
├── downloads/               # Downloaded files
└── temp/                    # Temporary files
```

## Security Considerations

1. **Credentials**: Never commit passwords to version control
2. **Network Access**: Ensure proper network security
3. **File Permissions**: Set appropriate file permissions
4. **Logging**: Be careful not to log sensitive information

## Performance Tips

1. **Headless Mode**: Use headless browser for better performance
2. **File Cleanup**: Regular cleanup of temporary files
3. **Network Optimization**: Ensure good network connectivity
4. **Resource Management**: Monitor memory and CPU usage

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with proper testing
4. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review log files for error details
3. Create an issue in the repository

## Version History

- v1.0.0 - Initial release with basic automation features
- v1.1.0 - Added SharePoint integration
- v1.2.0 - Added email notifications and improved error handling

## Dependencies

- selenium - Web automation
- xlwings - Excel integration
- Office365-REST-Python-Client - SharePoint integration
- requests - HTTP requests
- pandas - Data processing
- pathlib2 - Path handling

For a complete list, see `requirements.txt`.

## Environment Variables

You can use environment variables for sensitive configuration:

```bash
export SHAREPOINT_USERNAME="your_username"
export SHAREPOINT_PASSWORD="your_password"
export WEB_USERNAME="your_web_username"
export WEB_PASSWORD="your_web_password"
```

Then reference them in your config:
```python
import os
SHAREPOINT_CONFIG = {
    'username': os.getenv('SHAREPOINT_USERNAME'),
    'password': os.getenv('SHAREPOINT_PASSWORD'),
}
```

---

**Note**: This automation script is designed for legitimate business use cases. Ensure you have proper permissions and comply with your organization's policies before running automated scripts against web applications.