# Python_Excel_Data_Entry
 
## Description
This script utilizes the `win32com.client` module in Python to interact with Excel. It reads data from `rot.json` and `mamul.json` to perform operations on an active Excel workbook. Users are prompted to select a sheet, and based on the selection, the script updates cell values and formatting.

## Dependencies
- `win32com.client`
- `json`
- `questionary`

## Usage
1. Install the required Python modules.
   ```bash
   pip install pywin32 json questionary

Run the script.
python app.py

Follow on-screen prompts to select a sheet, enter the start cell, and specify the number of cells on the left.

Important Note
Ensure that Excel is installed on your machine.
The script may need adjustments based on your JSON file structure.
Error Handling
The script has basic exception handling to capture and display errors. If an error occurs, the script prints an error message with details.

Termination
To exit the program, select "Çıkış" when prompted to choose a sheet.    
