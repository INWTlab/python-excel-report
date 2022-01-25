# python-excel-report
_Create pretty Excel reports with Python_

## Basic usage

First make sure the following Python packages are installed:

- openpyxl
- pandas
- XlsxWriter

It's sufficient to adapt the `input_file_path` in the `CONFIG` object and then run the whole script.

## Configuration

All configuration takes place in the `CONFIG` object at the top of the script.
- `input_file_path`: path to a csv file with your data. 
  The data will be read from this path after the definition of the `CONFIG` object.
  If you want to load the data from any other file than a csv file, also adapt the code that loads the file.
- `output_file_name`: path where the Excel file will be saved
- `sheet_name`: name of the spreadsheet in the Excel file
- `column_header_format`: formatting options for the column names, e.g., font, font colour, and font size
- `text_format`: formatting options for the columns' content
- `column_widths`: widths of the columns in the Excel file.
  It's a dictionary where the keys are the column numbers (starting from 0)
  and the values are the respective width.
