"""
Google Patent PDF Downloader - GUI Version
A user-friendly graphical interface for downloading patents from Google Patents
Compatible with Python 3.8+
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import os
from pathlib import Path
import threading
import logging

# Set console encoding for Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('patent_download_gui.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Setup separate logger for failed downloads
failed_logger = logging.getLogger('failed_patents')
failed_logger.setLevel(logging.INFO)
failed_handler = logging.FileHandler('failed_patents.log')
failed_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
failed_logger.addHandler(failed_handler)
failed_logger.propagate = False  # Don't propagate to root logger


class PatentDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patent PDF Downloader")
        self.root.geometry("1000x850")
        self.root.resizable(True, True)
        
        # Set minimum window size to ensure buttons are always visible
        self.root.minsize(900, 750)
        
        # Modern color scheme
        self.colors = {
            'primary': '#1a73e8',      # Modern blue
            'primary_dark': '#1557b0',  # Darker blue for hover
            'success': '#34a853',       # Green
            'danger': '#ea4335',        # Red
            'background': '#f8f9fa',    # Light gray background
            'surface': '#ffffff',       # White
            'text': '#202124',          # Dark gray text
            'text_secondary': '#5f6368',# Medium gray
            'border': '#dadce0',        # Light border
            'header': '#1f1f1f',        # Dark header
            'accent': '#fbbc04'         # Yellow/Orange accent
        }
        
        # Set window background
        self.root.configure(bg=self.colors['background'])
        
        # Variables
        self.excel_file = tk.StringVar()
        self.output_dir = "downloaded_patents"
        self.is_downloading = False
        self.driver = None
        self.failed_patents = []  # Track failed patents
        self.direct_download_first = tk.BooleanVar(value=True)  # Try direct download without Chrome first
        
        # Create GUI
        self.create_widgets()
        
    def create_widgets(self):
        """Create all GUI widgets"""
        
        # Title Frame with gradient effect
        title_frame = tk.Frame(self.root, bg=self.colors['header'], pady=25)
        title_frame.pack(fill=tk.X, side=tk.TOP)
        
        # Icon + Title
        title_container = tk.Frame(title_frame, bg=self.colors['header'])
        title_container.pack()
        
        title_label = tk.Label(
            title_container,
            text="📄 Patent Downloader",
            font=("Segoe UI", 24, "bold"),
            bg=self.colors['header'],
            fg="white"
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            title_frame,
            text="Download patents automatically from Google Patents",
            font=("Segoe UI", 11),
            bg=self.colors['header'],
            fg="#b8b8b8"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Main Frame with modern background
        main_frame = tk.Frame(
            self.root, 
            padx=30, 
            pady=20,
            bg=self.colors['background']
        )
        main_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        
        # IMPORTANT: Pack bottom buttons FIRST with side=BOTTOM to ensure they're always visible
        # Bottom Frame with modern buttons - PACK THIS FIRST!
        bottom_frame = tk.Frame(main_frame, bg=self.colors['background'], pady=8)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, expand=False)
        
        # Left side buttons
        left_buttons = tk.Frame(bottom_frame, bg=self.colors['background'])
        left_buttons.pack(side=tk.LEFT)
        
        self.stop_btn = tk.Button(
            left_buttons,
            text="⏹️ Stop Download",
            command=self.stop_download,
            bg=self.colors['danger'],
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=15,
            pady=8,
            state=tk.DISABLED,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#c5221f",
            activeforeground="white",
            disabledforeground="white"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        open_folder_btn = tk.Button(
            left_buttons,
            text="📂 Open Downloads Folder",
            command=self.open_output_folder,
            bg="#5f6368",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#4a4d50",
            activeforeground="white",
            disabledforeground="white"
        )
        open_folder_btn.pack(side=tk.LEFT)
        
        # Right side buttons - Log files
        right_buttons = tk.Frame(bottom_frame, bg=self.colors['background'])
        right_buttons.pack(side=tk.RIGHT)
        
        open_main_log_btn = tk.Button(
            right_buttons,
            text="📄 Main Log",
            command=self.open_main_log,
            bg=self.colors['primary'],
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground=self.colors['primary_dark'],
            activeforeground="white"
        )
        open_main_log_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        open_failed_log_btn = tk.Button(
            right_buttons,
            text="❌ Failed Patents Log",
            command=self.open_failed_log,
            bg=self.colors['accent'],
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#e5a903",
            activeforeground="white"
        )
        open_failed_log_btn.pack(side=tk.LEFT)
        
        # NOW pack everything else from top to bottom
        # File Selection Section - Modern card style
        file_frame = tk.Frame(
            main_frame,
            bg=self.colors['surface'],
            highlightbackground=self.colors['border'],
            highlightthickness=1,
            padx=15,
            pady=12
        )
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Section header
        section_label = tk.Label(
            file_frame,
            text="📁 Select Excel File",
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['surface'],
            fg=self.colors['text'],
            anchor=tk.W
        )
        section_label.pack(fill=tk.X, pady=(0, 8))
        
        file_path_frame = tk.Frame(file_frame, bg=self.colors['surface'])
        file_path_frame.pack(fill=tk.X)
        
        self.file_entry = tk.Entry(
            file_path_frame,
            textvariable=self.excel_file,
            font=("Segoe UI", 10),
            state="readonly",
            bg="#f1f3f4",
            fg=self.colors['text'],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary']
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=6)
        
        browse_btn = tk.Button(
            file_path_frame,
            text="📂 Browse Files",
            command=self.browse_file,
            bg=self.colors['primary'],
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=6,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground=self.colors['primary_dark'],
            activeforeground="white"
        )
        browse_btn.pack(side=tk.RIGHT)
        
        info_label = tk.Label(
            file_frame,
            text="💡 Excel file must have a 'Display Key' column with patent numbers",
            font=("Segoe UI", 8),
            fg=self.colors['text_secondary'],
            bg=self.colors['surface']
        )
        info_label.pack(anchor=tk.W, pady=(6, 0))
        
        # Download Button Section (Direct download enabled by default)
        button_frame = tk.Frame(main_frame, bg=self.colors['background'])
        button_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.download_btn = tk.Button(
            button_frame,
            text="🚀 Start Download",
            command=self.start_download,
            bg=self.colors['success'],
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=40,
            pady=10,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#2d8f47",
            activeforeground="white"
        )
        self.download_btn.pack()
        
        # Progress Section - Modern card style
        progress_frame = tk.Frame(
            main_frame,
            bg=self.colors['surface'],
            highlightbackground=self.colors['border'],
            highlightthickness=1,
            padx=15,
            pady=12
        )
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        # Section header
        progress_label = tk.Label(
            progress_frame,
            text="📊 Download Progress",
            font=("Segoe UI", 12, "bold"),
            bg=self.colors['surface'],
            fg=self.colors['text'],
            anchor=tk.W
        )
        progress_label.pack(fill=tk.X, pady=(0, 10))
        
        # Progress bar with style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor='#e8eaed',
            background=self.colors['primary'],
            bordercolor=self.colors['border'],
            lightcolor=self.colors['primary'],
            darkcolor=self.colors['primary']
        )
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 12), ipady=8)
        
        # Status label with icon
        self.status_label = tk.Label(
            progress_frame,
            text="⏳ Ready to download",
            font=("Segoe UI", 10),
            fg=self.colors['text'],
            bg=self.colors['surface']
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Log area with modern styling
        log_label = tk.Label(
            progress_frame,
            text="📋 Activity Log",
            font=("Segoe UI", 10, "bold"),
            anchor=tk.W,
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        log_label.pack(anchor=tk.W, pady=(0, 6))
        
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=8,
            font=("Consolas", 9),
            bg="#fafafa",
            fg=self.colors['text'],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary'],
            padx=10,
            pady=10,
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Bottom buttons already packed at the beginning of create_widgets()
        
    def browse_file(self):
        """Open file dialog to select Excel file"""
        filename = filedialog.askopenfilename(
            title='Select Excel File with Patent Numbers',
            filetypes=[
                ('Excel Files', '*.xlsx *.xls'),
                ('All Files', '*.*')
            ]
        )
        if filename:
            self.excel_file.set(filename)
            self.log(f"Selected file: {filename}")
            
    def log(self, message):
        """Add message to log area"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def update_status(self, message, status_type='info'):
        """Update status label with icon based on type"""
        icons = {
            'info': '⏳',
            'success': '✅',
            'error': '❌',
            'downloading': '⬇️',
            'complete': '🎉'
        }
        icon = icons.get(status_type, '⏳')
        self.status_label.config(text=f"{icon} {message}")
        self.root.update_idletasks()
        
    def update_progress(self, current, total):
        """Update progress bar"""
        percentage = (current / total * 100) if total > 0 else 0
        self.progress_var.set(percentage)
        self.root.update_idletasks()
        
    def open_output_folder(self):
        """Open the downloads folder in file explorer"""
        if os.path.exists(self.output_dir):
            os.startfile(os.path.abspath(self.output_dir))
        else:
            messagebox.showinfo("Info", "No downloads folder yet. Start downloading to create it!")
            
    def open_main_log(self):
        """Open the main log file"""
        log_file = "patent_download_gui.log"
        if os.path.exists(log_file):
            try:
                os.startfile(os.path.abspath(log_file))
            except Exception as e:
                messagebox.showerror("Error", f"Could not open log file:\n{e}")
        else:
            messagebox.showinfo("Info", "Main log file doesn't exist yet.\nIt will be created when you start downloading.")
            
    def open_failed_log(self):
        """Open the failed patents log file"""
        log_file = "failed_patents.log"
        if os.path.exists(log_file):
            try:
                os.startfile(os.path.abspath(log_file))
            except Exception as e:
                messagebox.showerror("Error", f"Could not open log file:\n{e}")
        else:
            messagebox.showinfo("Info", "No failed patents log yet.\nThis file is created when a patent fails to download.")
            
    def start_download(self):
        """Start the download process in a separate thread"""
        if not self.excel_file.get():
            messagebox.showwarning("No File Selected", "Please select an Excel file first!")
            return
            
        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("File Not Found", f"Excel file not found:\n{self.excel_file.get()}")
            return
            
        if self.is_downloading:
            messagebox.showinfo("Already Downloading", "Download is already in progress!")
            return
            
        # Start download in separate thread
        self.is_downloading = True
        self.failed_patents = []  # Clear failed patents list
        self.download_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        
        download_thread = threading.Thread(target=self.download_patents, daemon=True)
        download_thread.start()
        
    def stop_download(self):
        """Stop the download process"""
        self.is_downloading = False
        self.log("Stopping download...")
        
    def read_patent_numbers(self, column_name='Display Key'):
        """Read patent numbers from Excel file"""
        try:
            df = pd.read_excel(self.excel_file.get())
            self.log(f"Excel file loaded. Columns: {df.columns.tolist()}")
            
            if column_name not in df.columns:
                self.log(f"ERROR: Column '{column_name}' not found")
                self.log(f"Available columns: {df.columns.tolist()}")
                return []
            
            patent_numbers = df[column_name].dropna().tolist()
            self.log(f"Found {len(patent_numbers)} patent numbers")
            return patent_numbers
            
        except Exception as e:
            self.log(f"ERROR reading Excel: {e}")
            return []
            
    def setup_driver(self):
        """Setup Chrome WebDriver"""
        try:
            self.log("Initializing Chrome browser...")
            chrome_options = Options()
            
            # Create output directory
            Path(self.output_dir).mkdir(parents=True, exist_ok=True)
            
            # Set download directory
            prefs = {
                "download.default_directory": os.path.abspath(self.output_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "safebrowsing.enabled": True
            }
            chrome_options.add_experimental_option("prefs", prefs)
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.log("Browser initialized successfully!")
            return True
            
        except Exception as e:
            self.log(f"ERROR: Could not start Chrome browser")
            self.log(f"Details: {e}")
            self.log("Make sure Chrome and ChromeDriver are installed")
            messagebox.showerror(
                "Browser Error",
                "Could not start Chrome browser.\n\n"
                "Required:\n"
                "1. Google Chrome installed\n"
                "2. ChromeDriver installed\n\n"
                "Install: pip install webdriver-manager"
            )
            return False
            
    def clean_patent_number(self, patent_number):
        """Clean patent number"""
        cleaned = str(patent_number).strip()
        cleaned = cleaned.replace(' ', '').replace('-', '').replace(',', '').replace('/', '')
        return cleaned
        
    def construct_pdf_url(self, patent_number):
        """Construct direct PDF URL from patent number (works for many patents)"""
        clean_number = self.clean_patent_number(patent_number)
        
        # Google Patents PDF URL pattern
        # Example: US1234567A -> https://patentimages.storage.googleapis.com/.../US1234567.pdf
        # This is a common pattern but may not work for all patents
        
        # Try to construct the patent page URL first
        patent_url = f"https://patents.google.com/patent/{clean_number}"
        
        # For direct PDF, we'll try common patterns
        # Pattern 1: Standard format
        base_url = "https://patentimages.storage.googleapis.com"
        
        # Extract country code and number
        import re
        match = re.match(r'([A-Z]{2})(\d+)([A-Z]\d*)?', clean_number)
        if match:
            country = match.group(1)
            number = match.group(2)
            kind = match.group(3) if match.group(3) else ''
            
            # Try common PDF URL pattern
            # Note: The actual hash in the URL is unpredictable, so this is a fallback
            # We'll need to fetch the page to get the real PDF URL
            return None  # Return None to indicate we need to fetch the page
        
        return None
        
    def try_direct_download(self, patent_number):
        """Try to download patent directly without Chrome"""
        clean_number = self.clean_patent_number(patent_number)
        
        # Method 1: Try to fetch the patent page and extract PDF link using requests
        try:
            patent_url = f"https://patents.google.com/patent/{clean_number}"
            self.log(f"  Trying direct download (no browser)...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(patent_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            # Try to find PDF link in the HTML
            import re
            pdf_pattern = r'https://patentimages\.storage\.googleapis\.com/[^"\']+\.pdf'
            pdf_matches = re.findall(pdf_pattern, response.text)
            
            if pdf_matches:
                pdf_url = pdf_matches[0]
                self.log(f"  Found PDF URL: {pdf_url}")
                if self.download_pdf_direct(pdf_url, clean_number):
                    self.log(f"  Direct download successful!")
                    return True
            
            return False
            
        except Exception as e:
            self.log(f"  Direct download failed: {e}")
            return False
        
    def log_failed_patent(self, original_number, clean_number, reason, url):
        """Log failed patent to separate log file"""
        failed_info = {
            'original': original_number,
            'cleaned': clean_number,
            'reason': reason,
            'url': url
        }
        self.failed_patents.append(failed_info)
        
        # Log to separate failed patents file
        failed_logger.info(
            f"FAILED | Original: {original_number} | Cleaned: {clean_number} | "
            f"Reason: {reason} | URL: {url}"
        )
        
    def download_patent(self, patent_number):
        """Download a single patent - Direct download only, no browser fallback"""
        clean_number = self.clean_patent_number(patent_number)
        url = f"https://patents.google.com/patent/{clean_number}"
        
        # Try direct download (no Chrome needed)
        if self.try_direct_download(patent_number):
            return True
        else:
            # Direct download failed - just log it, don't use browser
            error_msg = "PDF URL not found or download failed"
            self.log(f"  FAILED: {error_msg}")
            self.log_failed_patent(patent_number, clean_number, error_msg, url)
            return False
            
    def download_pdf_direct(self, pdf_url, patent_number):
        """Download PDF directly"""
        try:
            response = requests.get(pdf_url, stream=True)
            response.raise_for_status()
            
            filename = os.path.join(self.output_dir, f"{patent_number}.pdf")
            with open(filename, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
                    
        except Exception as e:
            self.log(f"  ERROR downloading PDF: {e}")
            raise  # Re-raise the exception so the caller knows it failed
            
    def print_to_pdf(self, patent_number):
        """Print page to PDF"""
        try:
            import base64
            pdf_data = self.driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": False,
                "paperWidth": 8.5,
                "paperHeight": 11,
                "marginTop": 0.4,
                "marginBottom": 0.4,
                "marginLeft": 0.4,
                "marginRight": 0.4
            })
            
            filename = os.path.join(self.output_dir, f"{patent_number}.pdf")
            with open(filename, 'wb') as f:
                f.write(base64.b64decode(pdf_data['data']))
            return True
                
        except Exception as e:
            self.log(f"  ERROR printing to PDF: {e}")
            raise  # Re-raise the exception so the caller knows it failed
            
    def download_patents(self):
        """Main download process"""
        try:
            # Log session start to failed patents log
            failed_logger.info("="*80)
            failed_logger.info(f"NEW DOWNLOAD SESSION STARTED")
            failed_logger.info("="*80)
            
            # Read patent numbers
            patent_numbers = self.read_patent_numbers()
            
            if not patent_numbers:
                self.log("No patent numbers found!")
                messagebox.showerror("Error", "No patent numbers found in Excel file!")
                self.is_downloading = False
                self.download_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
            # Direct download mode - no browser needed
            self.log("Direct download mode - downloading PDFs without browser")
            self.log("Failed downloads will be logged to failed_patents.log")
                
            # Download each patent
            successful = 0
            failed = 0
            
            for i, patent_number in enumerate(patent_numbers, 1):
                if not self.is_downloading:
                    self.log("Download stopped by user")
                    break
                    
                self.update_status(f"Downloading {i}/{len(patent_numbers)}: {patent_number}", 'downloading')
                self.log(f"[{i}/{len(patent_numbers)}] Downloading: {patent_number}")
                
                if self.download_patent(patent_number):
                    successful += 1
                    self.log(f"  SUCCESS")
                else:
                    failed += 1
                    self.log(f"  FAILED")
                    
                self.update_progress(i, len(patent_numbers))
                time.sleep(2)
                
            # Summary
            self.log("\n" + "="*50)
            self.log("DOWNLOAD COMPLETE!")
            self.log("="*50)
            self.log(f"Total patents:  {len(patent_numbers)}")
            self.log(f"Successful:     {successful}")
            self.log(f"Failed:         {failed}")
            if failed > 0:
                self.log(f"Failed patents logged to: failed_patents.log")
            self.log("="*50)
            
            self.update_status(f"Complete! {successful}/{len(patent_numbers)} successful", 'complete')
            
            # Create message with failed patents info
            message = f"Downloaded {successful} out of {len(patent_numbers)} patents!\n\n"
            message += f"Files saved in: {os.path.abspath(self.output_dir)}"
            if failed > 0:
                message += f"\n\n{failed} patent(s) failed to download.\nCheck 'failed_patents.log' for details."
            
            messagebox.showinfo("Download Complete", message)
            
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            
        finally:
            if self.driver:
                self.driver.quit()
                self.log("Browser closed")
                
            self.is_downloading = False
            self.download_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)


def main():
    """Main function"""
    root = tk.Tk()
    app = PatentDownloaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

