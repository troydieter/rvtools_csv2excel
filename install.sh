#!/bin/bash
# Installation script for RVTools CSV to Excel Converter

echo "Installing RVTools CSV to Excel Converter dependencies..."

# Check if pip is installed
if ! command -v pip &> /dev/null; then
    echo "Error: pip is not installed. Please install Python and pip first."
    exit 1
fi

# Install required packages
echo "Installing required Python packages..."
pip install pandas openpyxl

# Make the script executable
chmod +x rvtools_csv2excel.py

echo "Installation complete!"
echo "You can now run the converter using: ./rvtools_csv2excel.py"
echo "For help and options, run: ./rvtools_csv2excel.py --help"