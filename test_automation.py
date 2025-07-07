#!/usr/bin/env python3
"""
Test script for the automation program
Tests individual components to help with troubleshooting
"""

import sys
import os
from pathlib import Path
import logging

# Add current directory to path for imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from automation_script import AutomationScript
    from config import (
        WEB_CONFIG, FILE_CONFIG, SHAREPOINT_CONFIG, 
        EXCEL_CONFIG, BROWSER_CONFIG
    )
except ImportError as e:
    print(f"Import error: {e}")
    print("Please ensure config.py exists and is properly configured")
    sys.exit(1)

class AutomationTester:
    def __init__(self):
        self.automation = AutomationScript()
        self.logger = logging.getLogger(__name__)
        self.test_results = {}

    def test_dependencies(self):
        """Test if all required dependencies are installed"""
        print("Testing dependencies...")
        
        dependencies = [
            ('selenium', 'selenium'),
            ('xlwings', 'xlwings'),
            ('Office365-REST-Python-Client', 'office365'),
            ('requests', 'requests'),
            ('pandas', 'pandas'),
            ('numpy', 'numpy'),
        ]
        
        failed_imports = []
        for dep_name, import_name in dependencies:
            try:
                __import__(import_name)
                print(f"✓ {dep_name} - OK")
            except ImportError:
                print(f"✗ {dep_name} - MISSING")
                failed_imports.append(dep_name)
        
        if failed_imports:
            print(f"\nMissing dependencies: {', '.join(failed_imports)}")
            print("Run: pip install -r requirements.txt")
            return False
        
        print("All dependencies installed successfully!")
        return True

    def test_browser_setup(self):
        """Test browser and WebDriver setup"""
        print("\nTesting browser setup...")
        
        try:
            driver = self.automation.setup_webdriver()
            driver.get("https://www.google.com")
            
            if "Google" in driver.title:
                print("✓ Browser setup - OK")
                driver.quit()
                return True
            else:
                print("✗ Browser setup - Failed to load Google")
                driver.quit()
                return False
                
        except Exception as e:
            print(f"✗ Browser setup failed: {str(e)}")
            return False

    def test_directories(self):
        """Test if required directories exist and are writable"""
        print("\nTesting directories...")
        
        directories = [
            FILE_CONFIG['download_folder'],
            FILE_CONFIG['local_temp_folder'],
            'logs'
        ]
        
        success = True
        for directory in directories:
            path = Path(directory)
            
            try:
                path.mkdir(parents=True, exist_ok=True)
                
                # Test write permissions
                test_file = path / 'test_write.txt'
                test_file.write_text('test')
                test_file.unlink()
                
                print(f"✓ {directory} - OK")
                
            except Exception as e:
                print(f"✗ {directory} - Error: {str(e)}")
                success = False
        
        return success

    def test_network_drive(self):
        """Test network drive access"""
        print("\nTesting network drive access...")
        
        network_path = Path(FILE_CONFIG['network_drive_path'])
        
        try:
            # Check if path exists
            if not network_path.exists():
                print(f"✗ Network drive path does not exist: {network_path}")
                return False
            
            # Test write access
            test_file = network_path / 'test_write.txt'
            test_file.write_text('test')
            test_file.unlink()
            
            print(f"✓ Network drive access - OK")
            return True
            
        except Exception as e:
            print(f"✗ Network drive access failed: {str(e)}")
            return False

    def test_excel_integration(self):
        """Test Excel/xlwings integration"""
        print("\nTesting Excel integration...")
        
        try:
            import xlwings as xw
            
            # Try to start Excel
            app = xw.App(visible=False)
            
            # Create a simple workbook
            wb = app.books.add()
            wb.sheets[0].range('A1').value = 'Test'
            
            # Close without saving
            wb.close()
            app.quit()
            
            print("✓ Excel integration - OK")
            return True
            
        except Exception as e:
            print(f"✗ Excel integration failed: {str(e)}")
            print("Note: Excel must be installed for macro functionality")
            return False

    def test_sharepoint_connection(self):
        """Test SharePoint connection"""
        print("\nTesting SharePoint connection...")
        
        try:
            from office365.runtime.auth.authentication_context import AuthenticationContext
            from office365.sharepoint.client_context import ClientContext
            
            # Try to authenticate
            ctx_auth = AuthenticationContext(SHAREPOINT_CONFIG['sharepoint_url'])
            
            # Note: This is a basic connection test
            # Actual authentication would require valid credentials
            print("✓ SharePoint libraries - OK")
            print("Note: Actual connection test requires valid credentials")
            return True
            
        except Exception as e:
            print(f"✗ SharePoint connection failed: {str(e)}")
            return False

    def test_configuration(self):
        """Test configuration validity"""
        print("\nTesting configuration...")
        
        issues = []
        
        # Check web configuration
        if not WEB_CONFIG['web_url'] or WEB_CONFIG['web_url'] == 'https://example.com':
            issues.append("Web URL not configured")
        
        if not WEB_CONFIG['export_button_selector']:
            issues.append("Export button selector not configured")
        
        # Check file configuration
        if not FILE_CONFIG['network_drive_path'] or FILE_CONFIG['network_drive_path'] == '/path/to/network/drive':
            issues.append("Network drive path not configured")
        
        # Check SharePoint configuration
        if not SHAREPOINT_CONFIG['sharepoint_url'] or 'yourcompany' in SHAREPOINT_CONFIG['sharepoint_url']:
            issues.append("SharePoint URL not configured")
        
        if not SHAREPOINT_CONFIG['username'] or 'your_username' in SHAREPOINT_CONFIG['username']:
            issues.append("SharePoint username not configured")
        
        if issues:
            print("✗ Configuration issues found:")
            for issue in issues:
                print(f"  - {issue}")
            return False
        
        print("✓ Configuration - OK")
        return True

    def test_web_selectors(self):
        """Test web selectors by loading the target page"""
        print("\nTesting web selectors...")
        
        if not WEB_CONFIG['web_url'] or WEB_CONFIG['web_url'] == 'https://example.com':
            print("✗ Web URL not configured, skipping selector test")
            return False
        
        try:
            driver = self.automation.setup_webdriver()
            driver.get(WEB_CONFIG['web_url'])
            
            # Test if page loads
            if driver.title:
                print(f"✓ Page loaded: {driver.title}")
            else:
                print("✗ Page failed to load properly")
                driver.quit()
                return False
            
            # Test selectors
            selectors_to_test = [
                ('Profile selector', WEB_CONFIG['profile_selector']),
                ('Date range selector', WEB_CONFIG['date_range_selector']),
                ('Export button selector', WEB_CONFIG['export_button_selector']),
            ]
            
            for name, selector in selectors_to_test:
                if selector:
                    try:
                        elements = driver.find_elements("css selector", selector)
                        if elements:
                            print(f"✓ {name} found")
                        else:
                            print(f"✗ {name} not found")
                    except Exception as e:
                        print(f"✗ {name} error: {str(e)}")
                else:
                    print(f"- {name} not configured")
            
            driver.quit()
            return True
            
        except Exception as e:
            print(f"✗ Web selector test failed: {str(e)}")
            return False

    def run_all_tests(self):
        """Run all tests and provide a summary"""
        print("="*60)
        print("AUTOMATION SYSTEM TESTS")
        print("="*60)
        
        tests = [
            ("Dependencies", self.test_dependencies),
            ("Browser Setup", self.test_browser_setup),
            ("Directories", self.test_directories),
            ("Network Drive", self.test_network_drive),
            ("Excel Integration", self.test_excel_integration),
            ("SharePoint Connection", self.test_sharepoint_connection),
            ("Configuration", self.test_configuration),
            ("Web Selectors", self.test_web_selectors),
        ]
        
        results = {}
        for test_name, test_func in tests:
            try:
                results[test_name] = test_func()
            except Exception as e:
                print(f"✗ {test_name} failed with exception: {str(e)}")
                results[test_name] = False
        
        # Summary
        print("\n" + "="*60)
        print("TEST SUMMARY")
        print("="*60)
        
        passed = sum(1 for result in results.values() if result)
        total = len(results)
        
        for test_name, result in results.items():
            status = "PASS" if result else "FAIL"
            print(f"{test_name:<20} - {status}")
        
        print(f"\nTotal: {passed}/{total} tests passed")
        
        if passed == total:
            print("✓ All tests passed! The automation system is ready to use.")
        else:
            print("✗ Some tests failed. Please address the issues above.")
        
        return passed == total

def main():
    """Main test function"""
    if len(sys.argv) > 1:
        test_name = sys.argv[1].lower()
        tester = AutomationTester()
        
        test_functions = {
            'deps': tester.test_dependencies,
            'browser': tester.test_browser_setup,
            'dirs': tester.test_directories,
            'network': tester.test_network_drive,
            'excel': tester.test_excel_integration,
            'sharepoint': tester.test_sharepoint_connection,
            'config': tester.test_configuration,
            'selectors': tester.test_web_selectors,
        }
        
        if test_name in test_functions:
            test_functions[test_name]()
        else:
            print(f"Unknown test: {test_name}")
            print(f"Available tests: {', '.join(test_functions.keys())}")
    else:
        # Run all tests
        tester = AutomationTester()
        tester.run_all_tests()

if __name__ == "__main__":
    main()