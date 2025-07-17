@echo off
echo Installing RVTools CSV to Excel Converter dependencies...

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: Python is not installed or not in PATH. Please install Python first.
    exit /b 1
)

:: Install required packages
echo Installing required Python packages...
pip install pandas openpyxl

echo Installation complete!
echo You can now run the converter using: python rvtools_csv2excel.py
echo For help and options, run: python rvtools_csv2excel.py --help
pause