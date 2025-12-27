# Excel Comparer

A Python GUI application that compares two Excel files with the same sheet and column names, highlighting differences and providing a summary report.

## Functionality

The Excel Comparer tool performs the following operations:

1. **File Comparison**: Compares two Excel files (.xlsx or .xls) that have the same sheet names and column structures
2. **Difference Detection**: Identifies three types of differences:
   - **Added Rows**: Rows present in File 2 but not in File 1 (highlighted in green)
   - **Removed Rows**: Rows present in File 1 but not in File 2 (highlighted in red)
   - **Modified Rows**: Rows that exist in both files but have different values (highlighted in yellow, with only changed cells colored)
3. **Smart Matching**: Uses intelligent algorithms to match rows:
   - Longest Common Subsequence (LCS) algorithm for row-by-row comparison
   - Key-based matching when identifier columns are detected (e.g., ID, Code, SKU)
4. **Data Normalization**: Normalizes data for accurate comparison:
   - Handles whitespace, Unicode characters, and formatting differences
   - Treats empty/NaN values consistently
   - Case-insensitive string comparison
5. **Summary Report**: Generates a comprehensive summary sheet showing:
   - Number of rows in each file
   - Count of added, removed, and modified rows
   - Total differences per sheet
   - Validation check to ensure data consistency

## Features

- **User-Friendly GUI**: Simple graphical interface built with tkinter
- **Color-Coded Output**: Visual highlighting of differences in the output Excel file
- **Multi-Sheet Support**: Compares all common sheets between the two files
- **Timestamped Output**: Output files are automatically named with timestamps
- **Configurable**: Option to ignore specific columns via the `IGNORE_COLUMNS` configuration

## Requirements

- Python 3.10 or higher
- pandas
- openpyxl


## Steps to Run

1. **Install Dependencies** (if not already installed):
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Application**:
   ```bash
   python excel_comparer.py
   ```

3. **Using the GUI**:
   - Click "Upload Excel File 1" to select the first Excel file (baseline/reference file)
   - Click "Upload Excel File 2" to select the second Excel file (file to compare)
   - Click "Select Output Folder" to choose where the comparison results will be saved
   - Click "Compare Excel Files" to start the comparison
   - Wait for the comparison to complete (status will be shown in blue/green)
   - The output file will be saved as `YYYYMMDD_HHMMSS_excel_differences.xlsx` in the selected output folder

4. **View Results**:
   - Open the generated Excel file
   - Check the "Difference_Summary" sheet for an overview
   - Review individual sheet tabs for detailed differences with color coding:
     - 🟢 Green: Added rows
     - 🔴 Red: Removed rows
     - 🟡 Yellow: Modified cells

## Converting to Executable (.exe) with PyInstaller

Follow these steps to convert the Python script into a standalone executable file:

### Prerequisites

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Ensure all dependencies are installed**:
   ```bash
   pip install -r requirements.txt
   ```

### Steps to Create Executable

1. **Navigate to the project directory**:
   ```bash
   cd /path/to/ExcelComparer
   ```

2. **Create the executable** (basic command):
   ```bash
   pyinstaller excel_comparer.py
   ```

3. **Recommended: Create a single-file executable with windowed mode**:
   ```bash
   pyinstaller --onefile --windowed --name "ExcelComparer" excel_comparer.py
   ```

   **Options explained**:
   - `--onefile`: Creates a single executable file instead of a folder
   - `--windowed`: Hides the console window (no terminal window appears)
   - `--name "ExcelComparer"`: Sets the name of the executable

4. **Advanced: Create executable with icon** (optional):
   ```bash
   pyinstaller --onefile --windowed --name "ExcelComparer" --icon=icon.ico excel_comparer.py
   ```
   (Note: You'll need to provide an `.ico` file for the icon)

5. **Find the executable**:
   - The executable will be created in the `dist` folder
   - For single-file mode: `dist/ExcelComparer.exe` (Windows) or `dist/ExcelComparer` (Linux/Mac)
   - For folder mode: `dist/excel_comparer/` directory containing the executable

### Additional PyInstaller Options

- **Include additional files** (if needed):
  ```bash
  pyinstaller --onefile --windowed --add-data "config.ini;." excel_comparer.py
  ```

- **Exclude unnecessary modules** (to reduce file size):
  ```bash
  pyinstaller --onefile --windowed --exclude-module matplotlib --exclude-module numpy excel_comparer.py
  ```

- **Create a spec file for customization**:
  ```bash
  pyinstaller --onefile --windowed excel_comparer.py
  ```
  Then edit `excel_comparer.spec` and rebuild:
  ```bash
  pyinstaller excel_comparer.spec
  ```

### Troubleshooting

- **If the executable is large**: Use `--exclude-module` to remove unused modules
- **If it fails to run**: Try without `--windowed` first to see error messages
- **If dependencies are missing**: Ensure all packages in `requirements.txt` are installed before building
- **For Windows**: Make sure you're running PyInstaller on Windows to create a `.exe` file
- **For cross-platform**: You need to build on each target platform separately

### Distribution

Once created, you can distribute the executable file from the `dist` folder. Users don't need Python installed to run it.

## Configuration

You can customize the comparison by modifying the `IGNORE_COLUMNS` variable in `excel_comparer.py`:

```python
IGNORE_COLUMNS = set(['ColumnName1', 'ColumnName2'])  # Add columns to ignore
```

## Notes

- The tool compares only sheets that exist in both files
- Columns must have the same names in both files for accurate comparison
- The comparison is case-insensitive and handles whitespace normalization
- Large files may take some time to process

## License

See LICENSE file for details.

