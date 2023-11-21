# Excel Formula Script

## Overview

This Python script automates tasks in Microsoft Excel, providing a streamlined workflow for users. It leverages the `win32com.client` library for Excel interaction.

## Features

### 1. Sheet Selection
   - Choose a sheet from the active Excel workbook.
   - Option to exit the program.

### 2. Route and Code Selection
   - Read route names and codes from `rot.json`.
   - Select a route for further actions.

### 3. Cell Range Input
   - Input the starting cell and the number of cells to the left.

### 4. Data Population
   - Populate Excel cells with values based on user inputs and data from `mamul.json`.

## Getting Started

### 1. Environment Setup
   - Ensure Python is installed.
   - Install the required libraries using the following command:
     ```
     pip install pywin32 questionary
     ```
   - These libraries are essential for the proper execution of the script.


### 2. Data Files
   - Prepare `rot.json` and `mamul.json` files with necessary data.

### 3. Run the Script
   - Execute the script in a Python environment.

### 4. Follow the Prompts
   - Select a sheet, choose a route, and input cell details as prompted.

## User Prompts

- **Sheet Selection:** Choose a sheet or exit the program.
- **Route Selection:** Select a route from available options.
- **Cell Range Input:** Input starting cell and the number of cells to the left.
- **Additional Information:** Follow prompts for data input.

## Error Handling

The script includes basic error handling to catch exceptions and display error messages.

## Dependencies

- Python 3.x
- `win32com.client` library
- `questionary` library

## Notes

- Customize the script for specific Excel automation tasks.
- Adjustments may be needed for different use cases.
