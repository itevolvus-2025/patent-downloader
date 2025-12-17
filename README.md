# 📄 Google Patent PDF Downloader

A beautiful, modern GUI application to automatically download patents from Google Patents.

![Version](https://img.shields.io/badge/version-2.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

## ✨ Features

- 🎨 **Modern GUI** - Beautiful, elegant interface with Google Material Design colors
- 📊 **Real-time Progress** - Live progress bar and activity log
- 🚀 **Fast & Efficient** - Download hundreds of patents automatically
- ⏸️ **Stop/Resume** - Full control over downloads
- 📂 **Easy Access** - One-click to open downloads folder
- 🛡️ **Error Handling** - Automatic fallback methods for reliable downloads
- 💾 **Auto-save** - All PDFs saved with patent numbers as filenames

## 🚀 Quick Start

### Windows Users (Easiest!)

1. **Double-click:** `🚀 START HERE - GUI.bat`
2. **Click "Browse Files"** → Select your Excel file
3. **Click "Start Download"** → Done!

### Manual Method

```bash
python patent_downloader_gui.py
```

## 📋 Requirements

- **Python:** 3.8 or higher
- **Browser:** Google Chrome
- **Excel File:** Must have "Display Key" column with patent numbers
- **Internet:** Required for downloading

## 🔧 Installation

### First Time Setup

1. **Clone this repository:**
   ```bash
   git clone https://github.com/yourusername/patent-downloader.git
   cd patent-downloader
   ```

2. **Install Python packages:**
   ```bash
   pip install -r requirements.txt
   ```
   
   Or double-click: `install_requirements.bat` (Windows)

3. **Run the application:**
   - Windows: Double-click `🚀 START HERE - GUI.bat`
   - Manual: `python patent_downloader_gui.py`

## 📁 Excel File Format

Your Excel file must have a column named **"Display Key"** with patent numbers:

| Display Key |
|-------------|
| US1234567A |
| EP9876543B1 |
| WO2020123456A1 |

## 📂 Output

All downloaded PDFs are saved in: `downloaded_patents/`

Each file is named: `PatentNumber.pdf` (e.g., `US1234567A.pdf`)

## 🎨 Screenshots

### Main Interface
Modern, clean GUI with real-time progress tracking and activity log.

### Features in Action
- File browser for easy Excel selection
- Real-time progress bar
- Detailed activity log
- Action buttons for control

## 🛠️ Troubleshooting

### "Python not found"
- Install Python from [python.org](https://www.python.org/downloads/)
- ✅ Check "Add Python to PATH" during installation

### "Module not found"
- Run: `install_requirements.bat` or `pip install -r requirements.txt`

### "Chrome browser error"
- Make sure Google Chrome is installed
- ChromeDriver will be installed automatically

### "No patent numbers found"
- Check your Excel has a "Display Key" column
- Make sure the column contains patent numbers

## 📦 Dependencies

- **pandas** - Excel file reading
- **openpyxl** - Excel format support
- **selenium** - Browser automation
- **requests** - HTTP downloads
- **tkinter** - GUI (included with Python)

## 🎯 How It Works

1. Reads patent numbers from Excel file
2. Opens Chrome browser automatically
3. Navigates to each patent on Google Patents
4. Downloads PDF using multiple methods:
   - Direct PDF link (fastest)
   - Download button click
   - Print to PDF (fallback)
5. Saves with patent number as filename
6. Shows progress and summary

## 💡 Tips

- **Batch Processing:** Download hundreds of patents at once
- **Resume:** If interrupted, just run again
- **Logs:** Check `.log` files for detailed information
- **Backup:** Old downloads are preserved in backup folders

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👨‍💻 Author

Created with ❤️ for patent researchers and professionals

## 🌟 Support

If you find this tool helpful, please give it a ⭐ on GitHub!

## 📞 Contact

For issues, questions, or suggestions, please open an issue on GitHub.

---

**Version:** 2.0  
**Compatible with:** Python 3.8 - 3.14+  
**Last Updated:** December 2025

