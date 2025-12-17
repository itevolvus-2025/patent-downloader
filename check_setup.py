"""
Check if the system is properly set up to run the Patent Downloader
"""

import sys
import platform

# Set console encoding for Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

def check_python_version():
    """Check Python version"""
    version = sys.version_info
    print(f"[OK] Python {version.major}.{version.minor}.{version.micro} detected")
    
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("  WARNING: Python 3.8+ recommended")
        return False
    return True

def check_packages():
    """Check if required packages are installed"""
    required = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'selenium': 'selenium',
        'requests': 'requests',
        'tkinter': 'tkinter (for GUI)',
    }
    
    missing = []
    
    for module, name in required.items():
        try:
            __import__(module)
            print(f"[OK] {name} installed")
        except ImportError:
            print(f"[X] {name} NOT installed")
            missing.append(module)
    
    return missing

def check_chrome():
    """Check if Chrome browser is available"""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        
        driver = webdriver.Chrome(options=options)
        driver.quit()
        print("[OK] Chrome and ChromeDriver working")
        return True
    except Exception as e:
        print(f"[X] Chrome/ChromeDriver issue: {e}")
        return False

def main():
    print("=" * 60)
    print("  Patent Downloader - System Check")
    print("=" * 60)
    print()
    
    print("1. Checking Python...")
    python_ok = check_python_version()
    print()
    
    print("2. Checking required packages...")
    missing = check_packages()
    print()
    
    print("3. Checking Chrome browser...")
    chrome_ok = check_chrome()
    print()
    
    print("=" * 60)
    if python_ok and not missing and chrome_ok:
        print("[SUCCESS] ALL CHECKS PASSED!")
        print("  You can run: run.bat")
    else:
        print("[ERROR] ISSUES FOUND:")
        if missing:
            print(f"  Missing packages: {', '.join(missing)}")
            print("  Run: install_requirements.bat")
        if not chrome_ok:
            print("  Chrome/ChromeDriver needs to be installed")
            print("  Install: pip install webdriver-manager")
    print("=" * 60)
    print()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
