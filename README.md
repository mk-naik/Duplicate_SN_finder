# Multi-File Barcode Duplicate Finder

This application is a Python-based tool that allows you to detect duplicate barcodes across multiple Excel files. It's built with `Tkinter` for the graphical user interface and uses `pandas` for data processing.

## Features

- Select multiple Excel files and sheets to analyze for duplicates.
- Automatically detects columns containing barcodes.
- Supports Excel files with `.xlsx` and `.xls` formats.
- Displays progress during the processing of files.
- Outputs a detailed Excel report listing duplicate barcodes and their file locations.

## Requirements

To run this application, you need:

- **Python 3.8 or above**
- Required Python packages:
  - `tkinter` (comes with Python standard library)
  - `pandas` (working with tabular data like Excel files or CSVs.)
  - `openpyxl` (for reading and writing `.xlsx` files)
  - `xlrd` (for reading older `.xls` files)
  - `psutil` (Provides system and process utilities to get information about CPU, memory, disk, and network usage, and can be used for performance monitoring)

## Installation

1. Clone or download this repository:
   ```bash
   git clone https://github.com/your-username/duplicate-finder.git
   cd duplicate-finder
   ```

2. Install the required Python packages:
   ```bash
   pip install pandas openpyxl xlrd psutil
   ```

3. Run the application:
   ```bash
   python duplicate_finder.py
   ```

## Usage

1. Launch the application. The main interface will appear.
2. Click **Select Files** to choose one or more Excel files.
3. For each file, select the desired sheet to analyze using the dropdown menu.
4. Click **Start Duplicate Check** to find duplicates across the selected files and sheets.
5. If duplicates are found, they will be saved to an Excel file named `All_Duplicates_<timestamp>.xlsx` in the application directory.

## Output

The output file will contain:
- A column for the duplicate barcodes.
- Columns listing the file(s) where the duplicate appears.

The number of columns dynamically adjusts to fit the data.

## Screenshots

_Add screenshots of the application UI here._

## Error Handling

- If no files or sheets are selected, appropriate warnings are displayed.
- If no barcode columns are found in a file, a warning will notify you.

## Limitations

- The application assumes barcode data is in columns with "barcode" in their name (case insensitive).
- It skips rows with missing barcode data.
- Only supports Excel files.

## Future Enhancements

- Support for additional file formats like CSV.
- Configurable column detection rules.
- Export duplicate results to other formats (CSV, JSON).

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Contributing

Contributions are welcome! If you encounter bugs or have feature requests, feel free to submit an issue or create a pull request.

