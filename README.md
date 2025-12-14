# Excel File Comparison Tool

A simple desktop app for comparing two Excel files and spotting differences. Built this for work, sharing in case it helps someone else.

## What It Does

- Compares two Excel files and shows missing/extra records
- Optional detailed comparison for specific columns
- Customizable column mapping
- Dark UI with persistent settings

## Quick Start

```bash
git clone https://github.com/yourusername/excel-comparison-tool.git
cd excel-comparison-tool
pip install pandas xlrd openpyxl
python excel_comparison.py
```

## How to Use

1. Select two Excel files
2. Click "Start Comparison"
3. View results

For detailed comparison, check the box before starting.

## Settings

Click "Settings" to configure column indices. (starts from 0)

Settings save automatically in `excel_compare_settings.json`.

## Build Standalone App

```bash
pip install pyinstaller
pyinstaller --onefile --windowed excel_comparison.py
```

Executable will be in `dist` folder.

## Note

Column indices start from 0. Test with small files first to verify your settings.

## License

MIT
