# Excel Automation Tool 🚀

A lightweight, modern GUI application built with Python to automate batch processing of Excel files. This tool allows users to scan directories, select specific files, and perform bulk actions like cloning or converting entire workbooks to PDF without opening Excel manually.

## 🌟 Features

* **Smart Scanning:** Automatically detects `.xlsx`, `.xls`, and `.csv` files in a selected directory.
* **Interactive Data Table:** View file details (Name, Type, Size) in a clean, sortable table.
* **Batch Operations:** Select individual files or use "Select All" to process multiple files at once.
* **PDF Conversion:**
    * **Full Workbook Support:** Automatically converts *all* sheets in a workbook to a single PDF (not just the active sheet).
    * **Background Processing:** Runs on a separate thread with a progress bar, ensuring the app never freezes during large jobs.
* **Safe Execution:** Robust error handling ensures no "zombie" Excel processes are left running in the background.
* **Modern UI:** Built with `CustomTkinter` for a sleek, high-DPI aware interface that matches system themes.

## 🛠️ Technical Stack

* **Language:** Python 3.10+
* **GUI Framework:** [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
* **Automation:** `pywin32` (Win32 COM) for native Excel control.
* **Threading:** `pythoncom` & `threading` for non-blocking UI operations.

## 📋 Prerequisites

To run this application (either via source or `.exe`), the host machine must have:

1.  **Windows OS** (Windows 10/11 recommended).
2.  **Microsoft Excel** installed (Desktop version).

## 📦 Installation & Usage

### Option 1: Run the Executable (Recommended for Users)

No Python installation is required.

1.  Download the latest `main.exe` from the `dist` folder.
2.  Double-click `main.exe` to launch.
3.  Follow the on-screen instructions.

### Option 2: Run from Source (For Developers)

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/KYFoong/excel-automation-clone-pdf.git
    cd excel-automation-clone-pdf
    ```

2.  **Install dependencies:**
    ```bash
    pip install customtkinter CTkMessagebox pywin32
    ```

3.  **Run the application:**
    ```bash
    python main.py
    ```

## 📝 How to Use

1.  **Select Source:** Click "Browse Folder" to select the directory containing your Excel files.
2.  **Select Files:**
    * The table will populate with detected files.
    * Check the boxes next to the files you want to process, or click the "Select" header to toggle all.
3.  **Select Destination:** Click "Select Output" to choose where the processed files will be saved.
4.  **Choose Action:**
    * **Clone Files:** Copies the selected files to the destination folder.
    * **Convert to PDF:** Converts the selected workbooks into PDF format. Watch the progress bar for status.

## ⚠️ Troubleshooting

* **App Freezes?** The app is designed to be threaded. If it freezes, ensure no dialog boxes are open in Excel itself (e.g., "Activate License" or "Recovery" popups).
* **PDF Conversion Failed?** Ensure the Excel file is not corrupted and is not password protected.

---

**Developer Note:**
This project demonstrates the use of `pythoncom` CoInitialize/CoUninitialize to handle OLE automation safely within background threads, preventing GUI blocking during heavy COM operations.
