# Python Excel Data Entry Script

This Python script is designed to read data from old Excel files and write it into new Excel files. It uses the `openpyxl` library to interact with Excel files.

## How it works

1. The script first loads a JSON file that contains the names of Excel files. 
2. It then iterates over all Excel files in a specified directory (old_excel_directory). For each old Excel file, it opens the workbook and checks each sheet for specific search values ("MODUL", "MODÜL").
3. If these values are found, it collects the values in the rows below these cells and stores them in lists. It then adjusts the order of some of these lists to meet a specific format.
4. The script then iterates over new Excel files in a different directory (new_excel_directory). If it finds a new Excel file that matches the name of the sheet it's currently processing, it opens the new workbook and writes the collected values into specific columns of the active sheet.
5. The script also handles exceptions and prints an error message if it encounters any issues when opening a workbook.

## Requirements

- Python 3.6+
- openpyxl library

## Usage

1. Update the `new_excel_directory`, `old_excel_directory`, and `json_file_path` variables with your own paths.
2. Run the script with `python copy_old_schema_to_new_schema.py`.

## Note

This script is designed to work with a specific format of Excel files. If your files have a different format, you may need to adjust the script accordingly.