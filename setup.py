#!/usr/bin/env python3
"""
Setup script for the automation program
Installs all required dependencies and sets up the environment
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

def run_command(command, description=""):
    """Run a shell command and handle errors"""
    print(f"\n{'='*50}")
    print(f"Running: {description or command}")
    print(f"{'='*50}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {command}")
        print(f"Error output: {e.stderr}")
        return False

def check_python_version():
    """Check if Python version is compatible"""
    print("Checking Python version...")
    if sys.version_info < (3, 8):
        print("ERROR: Python 3.8 or higher is required")
        sys.exit(1)
    print(f"Python version: {sys.version}")

def install_chrome_driver():
    """Install Chrome WebDriver"""
    print("\nInstalling Chrome WebDriver...")
    
    system = platform.system().lower()
    
    if system == "linux":
        # Install Chrome if not present
        commands = [
            "wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | apt-key add -",
            "echo 'deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main' > /etc/apt/sources.list.d/google-chrome.list",
            "apt-get update",
            "apt-get install -y google-chrome-stable"
        ]
        
        for cmd in commands:
            if not run_command(cmd, f"Installing Chrome on Linux"):
                print("Note: Chrome installation may have failed. Please install Chrome manually.")
                break
    
    elif system == "darwin":  # macOS
        if not run_command("brew install --cask google-chrome", "Installing Chrome on macOS"):
            print("Note: Chrome installation may have failed. Please install Chrome manually.")
    
    elif system == "windows":
        print("Please install Chrome manually from https://www.google.com/chrome/")
    
    # Install WebDriver Manager
    run_command("pip install webdriver-manager", "Installing WebDriver Manager")

def create_directories():
    """Create necessary directories"""
    directories = [
        'logs',
        'downloads',
        'temp',
        'config'
    ]
    
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"Created directory: {directory}")

def install_python_dependencies():
    """Install Python dependencies"""
    print("\nInstalling Python dependencies...")
    
    # Upgrade pip first
    run_command("pip install --upgrade pip", "Upgrading pip")
    
    # Install requirements
    if Path("requirements.txt").exists():
        run_command("pip install -r requirements.txt", "Installing requirements from requirements.txt")
    else:
        print("requirements.txt not found. Installing dependencies manually...")
        
        dependencies = [
            "selenium==4.15.2",
            "webdriver-manager==4.0.1",
            "xlwings==0.31.4",
            "openpyxl==3.1.2",
            "Office365-REST-Python-Client==2.5.3",
            "requests==2.31.0",
            "pandas==2.1.4",
            "numpy==1.24.3",
            "pathlib2==2.3.7",
            "cryptography==41.0.7",
            "msal==1.24.1"
        ]
        
        for dep in dependencies:
            run_command(f"pip install {dep}", f"Installing {dep}")

def setup_excel_support():
    """Setup Excel support based on operating system"""
    system = platform.system().lower()
    
    if system == "windows":
        print("\nSetting up Excel support for Windows...")
        run_command("pip install pywin32", "Installing pywin32 for Windows Excel support")
    
    elif system == "darwin":  # macOS
        print("\nSetting up Excel support for macOS...")
        print("xlwings should work with Excel for Mac. Please ensure Excel is installed.")
    
    elif system == "linux":
        print("\nSetting up Excel alternatives for Linux...")
        print("Note: xlwings requires Excel which is not available on Linux.")
        print("Consider using LibreOffice or run this on a Windows/macOS system.")

def create_sample_config():
    """Create a sample configuration file"""
    config_content = '''
# Sample configuration file
# Copy this to config.py and modify the values

from datetime import datetime, timedelta
from pathlib import Path

# IMPORTANT: Modify these values for your specific use case
WEB_CONFIG = {
    'web_url': 'https://your-website.com',
    'profile_selector': '#profile-dropdown',
    'date_range_selector': '#date-range-picker',
    'export_button_selector': '#export-button',
    'start_date_selector': '#start-date',
    'end_date_selector': '#end-date',
    'download_timeout': 60,
}

FILE_CONFIG = {
    'download_folder': str(Path.home() / 'Downloads'),
    'local_temp_folder': str(Path.home() / 'temp_automation'),
    'network_drive_path': '/path/to/your/network/drive',
    'sharepoint_folder': '/sites/yoursite/Documents/YourFolder',
    'exported_file_pattern': 'export*.xlsx',
    'target_file_name': 'processed_data.xlsx',
    'macro_file_name': 'your_macro.xlsm',
    'output_files': ['output1.xlsx', 'output2.xlsx'],
}

SHAREPOINT_CONFIG = {
    'sharepoint_url': 'https://yourcompany.sharepoint.com/sites/yoursite',
    'username': 'your_username@yourcompany.com',
    'password': 'your_password',
    'target_folder': '/sites/yoursite/Documents/YourFolder',
    'replace_existing': True,
}

# Add other configurations as needed...
'''
    
    with open('config_sample.py', 'w') as f:
        f.write(config_content)
    
    print("Created config_sample.py - copy this to config.py and modify values")

def main():
    """Main setup function"""
    print("="*60)
    print("AUTOMATION SCRIPT SETUP")
    print("="*60)
    
    # Check Python version
    check_python_version()
    
    # Create directories
    print("\nCreating directories...")
    create_directories()
    
    # Install Python dependencies
    install_python_dependencies()
    
    # Install Chrome WebDriver
    install_chrome_driver()
    
    # Setup Excel support
    setup_excel_support()
    
    # Create sample config
    create_sample_config()
    
    print("\n" + "="*60)
    print("SETUP COMPLETE!")
    print("="*60)
    print("\nNext steps:")
    print("1. Copy config_sample.py to config.py")
    print("2. Modify config.py with your specific settings")
    print("3. Test the automation by running: python automation_script.py")
    print("\nFor more information, see the README.md file")

if __name__ == "__main__":
    main()