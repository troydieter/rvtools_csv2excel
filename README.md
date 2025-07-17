# RVTools CSV to Excel Converter

This tool converts RVTools CSV exports to Excel format, preserving data and applying formatting to match the official RVTools Excel export format.

## Features

- Convert single or multiple CSV files to a single Excel workbook
- Scan directories for CSV files
- Apply RVTools-style formatting (header colors, fonts, etc.)
- Preserve data types
- Reorder sheets to match standard RVTools order
- Add metadata sheet
- Support for recursive directory scanning
- Filter files by prefix

## Requirements

- Python 3.6 or higher
- Required Python packages:
  - pandas
  - openpyxl

## Installation

1. Ensure you have Python 3.6+ installed
2. Install required packages:

```bash
pip install pandas openpyxl
```

3. Download the script or clone this repository

## Usage

### Basic Usage

```bash
python rvtools_csv2excel.py
```

This will scan the current directory for CSV files with the prefix "RVTools_tab" and create an Excel file named "rvtools_export.xlsx".

### Command Line Options

```
python rvtools_csv2excel.py [options]

Options:
  -h, --help              Show this help message and exit
  -i, --input DIR         Input directory containing CSV files (default: current directory)
  -o, --output FILE       Output Excel file (default: rvtools_export.xlsx)
  -r, --recursive         Scan subdirectories for CSV files
  -p, --prefix PREFIX     Only process files with this prefix (default: RVTools_tab)
  -v, --verbose           Show detailed processing information
```

### Examples

Convert all RVTools CSV files in the current directory:
```bash
python rvtools_csv2excel.py
```

Convert files in a specific directory:
```bash
python rvtools_csv2excel.py -i /path/to/csv/files
```

Specify an output file:
```bash
python rvtools_csv2excel.py -o my_rvtools_export.xlsx
```

Scan subdirectories recursively:
```bash
python rvtools_csv2excel.py -r
```

Process all CSV files (not just those with RVTools_tab prefix):
```bash
python rvtools_csv2excel.py -p ""
```

Show detailed processing information:
```bash
python rvtools_csv2excel.py -v
```

## Output Format

The output Excel file will have:
- One sheet per CSV file
- Sheets ordered according to standard RVTools order
- Black header row with white text (Verdana 9pt bold)
- Data rows with Verdana 9pt font
- Conditional formatting for certain columns (e.g., red text for powered off VMs)
- Auto-adjusted column widths
- Frozen header row
- Metadata sheet

## Troubleshooting

If you encounter issues:

1. Try running with the `-v` flag to see detailed processing information
2. Ensure your CSV files are properly formatted
3. Check that you have the required Python packages installed
4. Verify that the CSV files have the expected prefix (default: "RVTools_tab")

## License

This tool is provided under the APACHE 2.0 License.