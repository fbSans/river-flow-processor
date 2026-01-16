# Hydrological Data Processor

This Python project processes monthly hydrological Excel sheets, extracting daily averages for river flows and heights, and compiles them into consolidated Excel files. It demonstrates my skills in Python data processing, Excel automation, and handling real-world datasets.

## Features
- Processes monthly sheets (`OUTUBRO` → `SETEMBRO`) for multiple years.
- Automatically handles leap years when counting days.
- Extracts a fixed column of daily averages from each sheet.
- Marks invalid or non-numeric cells as `'f'`.
- Merges consecutive years into a single output file per segment.
- Saves processed data in new Excel `.xlsx` files.

## Requirements
- Python 3
- `openpyxl`
- `xlrd`

Install dependencies:

```bash
pip install openpyxl xlrd