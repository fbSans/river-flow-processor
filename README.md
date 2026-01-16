# Hydrological Data Processor

This Python script processes monthly hydrological Excel sheets, extracting daily averages for heights and flows, and compiling them into consolidated Excel files.

Original data files contain hydrological measurements (river levels and flows). 

## Features
- Processes monthly sheets (`OUTUBRO` → `SETEMBRO`) for multiple years.
- Handles leap years automatically when counting days.
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
