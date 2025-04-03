# Consolidator: Excel File Merger

## Overview
The **Consolidator** is a Python script designed to **automate** the process of merging multiple Excel files into a single master spreadsheet. It scans a given directory for Excel files, extracts relevant data, standardizes column formats, and consolidates everything into one unified file.

## Features
- **Automatic detection** of Excel files in a directory.
- **Standardizes columns** across all files, filling missing values with `N/A`.
- **Adds source tracking** by appending the filename as a new column.
- **Saves output** as an Excel file for easy access and analysis.

## Requirements
Ensure you have the following installed:
- Python 3.x
- `pandas` library
- `openpyxl` library (for `.xlsx` support)

You can install the required dependencies using:
```bash
pip install pandas openpyxl
```

## Usage
### 1. Set Up Your File Paths
Modify the script to specify:
- The **input directory** where your Excel files are stored.
- The **output file path** where the consolidated spreadsheet will be saved.

Example:
```python
consolidator_obj = Consolidator(r'path/to/your/excel/files', output_file='output/master.xlsx')
```

### 2. Run the Script
Execute the script using:
```bash
python consolidator.py
```

### 3. Output
- The script will scan all valid Excel files in the directory.
- It will consolidate the data into a single file, filling in missing columns with `N/A`.
- The resulting **master spreadsheet** will be saved to the specified output path.

## Code Breakdown
### **Class: `Consolidator`**
#### **Methods:**
1. **`check_files_and_append()`**
   - Reads all Excel files in the directory.
   - Ensures they contain the required columns.
   - Appends them to an internal data list.
   
2. **`consolidate()`**
   - Merges all extracted data into a single Pandas DataFrame.
   - Saves the final DataFrame as an Excel file.

## Error Handling
- The script **skips** non-Excel files.
- If a file fails to process, an error message is printed, and execution continues.
- If **no valid files** are found, the script notifies the user.

## Future Enhancements
- Add logging for better debugging.
- Implement a GUI for easier user interaction.
- Allow configuration of required columns via a settings file.

## License
This project is open-source and free to use.

---
Feel free to customize this README further based on your specific project needs!

