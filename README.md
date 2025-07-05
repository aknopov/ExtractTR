# Extractor of Soil Test Results

## User Notes
- The app copies soil test information from source file(s) to the destination.
- It merges certain cells in columns based on source samples and sorts the resulting data by the soil group in column "CG".
- The app can process a single Excel file or all files in a directory.
- Supported file formats: XLSX, XLSM, and XLSB.
- To clear any path in the main window, open the corresponding file dialog and press the "Cancel" button.
- The app searches for source and destination files starting the last used directories.
This directories are stored in `%USERPROFILE%\.extractrc` file
- It also creates a log file at `%USERPROFILE%\extractr.log`, which records processed files and any encountered errors.
To view the log, copy this path into Explorer. You may be prompted to choose a program to open the file.
- The original destination file is preserved as `*.org.xlsx`.

## Developer Notes
- The program is written in Python 3 using the `tkinter` and `openpyxl` packages.
- To create package, install "pyinstaller" and run `cd install; pyinstaller --clean ExtractTR.spec`.
You might need to exclude project directory from virus protection.
