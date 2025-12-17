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


class PatentDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Patent PDF Downloader")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
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
        
        # File Selection Section - Modern card style
        file_frame = tk.Frame(
            main_frame,
            bg=self.colors['surface'],
            highlightbackground=self.colors['border'],
            highlightthickness=1,
            padx=20,
            pady=20
        )
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Section header
        section_label = tk.Label(
            file_frame,
            text="📁 Select Excel File",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['surface'],
            fg=self.colors['text'],
            anchor=tk.W
        )
        section_label.pack(fill=tk.X, pady=(0, 12))
        
        file_path_frame = tk.Frame(file_frame, bg=self.colors['surface'])
        file_path_frame.pack(fill=tk.X)
        
        self.file_entry = tk.Entry(
            file_path_frame,
            textvariable=self.excel_file,
            font=("Segoe UI", 11),
            state="readonly",
            bg="#f1f3f4",
            fg=self.colors['text'],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary']
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12), ipady=8)
        
        browse_btn = tk.Button(
            file_path_frame,
            text="📂 Browse Files",
            command=self.browse_file,
            bg=self.colors['primary'],
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=25,
            pady=8,
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
            font=("Segoe UI", 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['surface']
        )
        info_label.pack(anchor=tk.W, pady=(10, 0))
        
        # Download Button Section
        button_frame = tk.Frame(main_frame, bg=self.colors['background'])
        button_frame.pack(fill=tk.X, pady=(5, 15))
        
        self.download_btn = tk.Button(
            button_frame,
            text="🚀 Start Download",
            command=self.start_download,
            bg=self.colors['success'],
            fg="white",
            font=("Segoe UI", 14, "bold"),
            padx=50,
            pady=15,
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
            padx=20,
            pady=15
        )
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Section header
        progress_label = tk.Label(
            progress_frame,
            text="📊 Download Progress",
            font=("Segoe UI", 13, "bold"),
            bg=self.colors['surface'],
            fg=self.colors['text'],
            anchor=tk.W
        )
        progress_label.pack(fill=tk.X, pady=(0, 15))
        
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
            font=("Segoe UI", 11),
            fg=self.colors['text'],
            bg=self.colors['surface']
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 15))
        
        # Log area with modern styling
        log_label = tk.Label(
            progress_frame,
            text="📋 Activity Log",
            font=("Segoe UI", 11, "bold"),
            anchor=tk.W,
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        log_label.pack(anchor=tk.W, pady=(0, 8))
        
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=7,
            font=("Consolas", 10),
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
        
        # Bottom Frame with modern buttons - FIXED POSITION
        bottom_frame = tk.Frame(main_frame, bg=self.colors['background'], pady=5)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, expand=False)
        
        self.stop_btn = tk.Button(
            bottom_frame,
            text="⏹️ Stop Download",
            command=self.stop_download,
            bg=self.colors['danger'],
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=25,
            pady=10,
            state=tk.DISABLED,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#c5221f",
            activeforeground="white",
            disabledforeground="white"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        open_folder_btn = tk.Button(
            bottom_frame,
            text="📂 Open Downloads Folder",
            command=self.open_output_folder,
            bg="#5f6368",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=25,
            pady=10,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            activebackground="#4a4d50",
            activeforeground="white",
            disabledforeground="white"
        )
        open_folder_btn.pack(side=tk.LEFT)
        
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
        
    def download_patent(self, patent_number):
        """Download a single patent"""
        try:
            clean_number = self.clean_patent_number(patent_number)
            url = f"https://patents.google.com/patent/{clean_number}"
            
            self.driver.get(url)
            time.sleep(3)
            
            # Check if patent exists
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "h1"))
                )
            except:
                self.log(f"  ERROR: Patent {clean_number} not found")
                return False
            
            # Try to find and download PDF
            try:
                # Method 1: Direct PDF link
                pdf_link = self.driver.find_element(By.XPATH, "//a[contains(@href, '.pdf')]")
                pdf_url = pdf_link.get_attribute('href')
                self.download_pdf_direct(pdf_url, clean_number)
                return True
            except:
                try:
                    # Method 2: Download button
                    download_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Download PDF')]"))
                    )
                    download_button.click()
                    time.sleep(2)
                    return True
                except:
                    # Method 3: Print to PDF
                    self.print_to_pdf(clean_number)
                    return True
                    
        except Exception as e:
            self.log(f"  ERROR: {e}")
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
                    
        except Exception as e:
            self.log(f"  ERROR downloading PDF: {e}")
            
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
                
        except Exception as e:
            self.log(f"  ERROR printing to PDF: {e}")
            
    def download_patents(self):
        """Main download process"""
        try:
            # Read patent numbers
            patent_numbers = self.read_patent_numbers()
            
            if not patent_numbers:
                self.log("No patent numbers found!")
                messagebox.showerror("Error", "No patent numbers found in Excel file!")
                self.is_downloading = False
                self.download_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
            # Setup browser
            if not self.setup_driver():
                self.is_downloading = False
                self.download_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                return
                
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
            self.log("="*50)
            
            self.update_status(f"Complete! {successful}/{len(patent_numbers)} successful", 'complete')
            
            messagebox.showinfo(
                "Download Complete",
                f"Downloaded {successful} out of {len(patent_numbers)} patents!\n\n"
                f"Files saved in: {os.path.abspath(self.output_dir)}"
            )
            
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

