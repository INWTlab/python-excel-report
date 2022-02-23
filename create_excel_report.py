import pandas as pd

# Configure everything here:
CONFIG = {
    "input_file_path": "input_data.csv",
    "output_file_name": "excel_report.xlsx",
    "sheet_name": "Report",
    "column_header_format": {"font": "Open Sans", "size": "11", "font_color": "#2b2b2b", "bold": True},
    "text_format": {
        "font": "Open Sans",
        "size": "11",
        "font_color": "#2b2b2b",
        "valign": "top",
    },
    "column_widths": {0: 6, 1: 12, 2: 12, 3: 16},
    "columns_2_decimal_places": [3],
    "column_heights": {0: 15, 1: 15, 2: 30, 3: 15},
    "columns_linebreaks": [1],
}

# -------------------------------------------------------------------------------------------------------------------- #

# Load data from csv file
data = pd.read_csv(CONFIG["input_file_path"])

# Set up basic Excel report
writer = pd.ExcelWriter(path=CONFIG["output_file_name"])
data.to_excel(writer, index=False, sheet_name=CONFIG["sheet_name"])
sheet_report = writer.sheets[CONFIG["sheet_name"]]
last_col = data.shape[1] - 1

# Format as table
sheet_report.add_table(
    first_row=0,
    first_col=0,
    last_row=data.shape[0] - 1,
    last_col=last_col,
    options={"columns": [{"header": col} for col in data.columns]},
)

# Set custom format for columns and header
workbook = writer.book
default_text_format = workbook.add_format(CONFIG["text_format"])
sheet_report.set_column(first_col=0, last_col=last_col, cell_format=default_text_format)
column_header_format = workbook.add_format(CONFIG["column_header_format"])
for col_num, col_name in enumerate(data.columns.values):
    sheet_report.write(0, col_num, col_name, column_header_format)


# Format number columns with 2 decimal places
num_format_2_decimals = CONFIG["text_format"].copy()
num_format_2_decimals["num_format"] = 0x02
num_format = workbook.add_format(num_format_2_decimals)
for col_number in CONFIG["columns_2_decimal_places"]:
    sheet_report.set_column(first_col=col_number, last_col=col_number, cell_format=num_format)

# Allow linebreaks in predefined columns
text_wrap_format = CONFIG["text_format"].copy()
text_wrap_format["text_wrap"] = "True"
text_wrap_format = workbook.add_format(text_wrap_format)
for col_number in CONFIG["columns_linebreaks"]:
    sheet_report.set_column(first_col=col_number, last_col=col_number, cell_format=text_wrap_format)

# Freeze header row
sheet_report.freeze_panes(row=1, col=0)

# Set column widths and heights
for col_index, width in CONFIG["column_widths"].items():
    sheet_report.set_column(col_index, col_index, width)
for row_index, height in CONFIG["column_heights"].items():
    sheet_report.set_row(row_index, height)

# Save file
writer.save()
