# Excel File Processor

## Description
This script is a simple GUI application to process Excel files with specific formatting. The application reads an input Excel file and generates a new Excel file with the processed data. The processing includes:
- Filtering specific columns from the input file.
- Creating a new sheet with only the rows that contain a "WARNING" status.
- Applying conditional formatting to specific cells based on the "QC status" column.
- Auto-adjusting column widths and row heights based on the content.

## Requirements
- Python 3.6 or higher
- pandas
- openpyxl
- tkinter


## Usage
1. Run the script using the command: `Excel_QC_processor.py`.
2. The application window will open. Click on the "Examinar" button to select the input Excel file.
3. Once the input file is selected, click on the "Run" button to start the processing.
4. The progress bar will show the progress of the processing, and a message will be displayed upon completion.
5. If any errors are encountered during processing, an error message will be displayed in the application window.

## Notes
- The input Excel file must contain specific column names for the script to work correctly. If the input file does not match the expected format, an error message will be displayed.
- The output file will be saved as "new_filename.xlsx" in the same directory as the script. Make sure to backup or rename any existing files with the same name before running the script, as they may be overwritten.

## Troubleshooting
If you encounter any issues or errors, please make sure that:
- The input Excel file is in the correct format and contains the required columns.
- The required packages (pandas, openpyxl, and tkinter) are installed correctly.
- You are using a compatible version of Python (3.6 or higher).
