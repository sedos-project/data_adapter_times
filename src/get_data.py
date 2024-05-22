import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


def format_and_save_excel(file_path, processed_df):
    """
    Format and save the processed data into an Excel file.

    Parameters:
    processed_df (pandas.DataFrame): The DataFrame containing the processed data.
    file_path (str): The path where the Excel file will be saved.

    Returns:
    str: The path where the Excel file is saved.
    """
    wb = load_workbook(file_path)
    ws = wb.active

    # Existing setup for fills, fonts, borders
    # Define two color fills for alternating rows
    fill1 = PatternFill(start_color="DDD9C4", end_color="DDD9C4", fill_type="solid")
    fill2 = PatternFill(start_color="C5d9F1", end_color="C5d9F1", fill_type="solid")
    header_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    subheader_fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )  # Assuming light green color for subheaders
    header_font = Font(name="Arial", size=10, bold=True, color="000000")
    sub_header_font = Font(name="Arial", size=10, bold=False, color="000000")
    table_name_font = Font(name="Arial", size=12, bold=True, color="0000FF")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    no_border = Border()

    def style_cell(cell, fill=None, font=None, border=None, alignment=None):
        if fill:
            cell.fill = fill
        if font:
            cell.font = font
        else:
            cell.font = sub_header_font
        if border is not None:  # Only apply the border if it is explicitly given
            cell.border = border
        if alignment:
            cell.alignment = alignment

    # Headers and subheaders for the Excel sheet
    headers = [
        "",
        "TechName",
        "*TechDesc",
        "Attribute",
        "Comm-IN",
        "Comm-OUT",
        "CommGrp",
        "TimeSlice",
        "LimType",
        "2021",
        "2024",
        "2027",
        "2030",
        "2035",
        "2040",
        "2045",
        "2050",
        "2060",
        "2070",
    ]

    subheaders = [
        "",
        "*Technology Name",
        "Technology Description",
        "Attribute Declaration\nColumn",
        "Input\nCommodity",
        "Output\nCommodity",
        "Commodity\nGroup",
        "Time Slices\ndefinition",
        "Bound\ndefinition",
        "Base\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
        "Data\nYear",
    ]

    # Write the table name with blue font
    table_name_cell = ws.cell(row=1, column=2, value="~FI_T")
    style_cell(
        table_name_cell, font=table_name_font, alignment=Alignment(horizontal="center")
    )

    # Write the headers with style
    for col, header_title in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header_title)
        style_cell(
            cell,
            fill=header_fill,
            font=header_font,
            border=thin_border,
            alignment=Alignment(horizontal="center", wrap_text=True),
        )

    # Write the subheaders with style
    for col, sub_header_title in enumerate(subheaders, start=1):
        cell = ws.cell(row=3, column=col, value=sub_header_title)
        style_cell(
            cell,
            fill=subheader_fill,
            font=sub_header_font,
            border=thin_border,
            alignment=Alignment(horizontal="center", wrap_text=True),
        )

    # Calculate column widths based on headers and subheaders
    column_widths = [
        max(len(header), max(len(part) for part in subheader.split("\n")))
        for header, subheader in zip(headers, subheaders)
    ]

    # Initialize the variable to keep track of the current process and the fill to apply
    current_process = None
    current_fill = fill1

    # Write the data and format the cells with alternating colors
    for row_index, (idx, row) in enumerate(
        processed_df.iterrows(), start=4
    ):  # Data starts from row 4
        process = row["TechName"]
        if process != current_process:
            # Switch the fill when the process changes
            current_fill = fill2 if current_fill == fill1 else fill1
            current_process = process

        for col_index, (col, value) in enumerate(row.items(), start=2):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            style_cell(cell, fill=current_fill, border=no_border)
            # Update the max length if the current value is longer
            column_widths[col_index - 1] = max(
                column_widths[col_index - 1], len(str(value))
            )

    # Set column widths with a little extra padding
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width + 1

    # Save the workbook
    wb.save(file_path)
    return file_path


def fetch_data(url):
    response = requests.get(url)
    return pd.DataFrame(response.json())


def data_mapping(df, process_name):
    """
    Fetches data from the API for a given process name and updates the times_df DataFrame.

    Parameters:
    times_df (pandas.DataFrame): The DataFrame containing the initial data.
    process_name (str): The name of the process to fetch and process data for.

    Returns:
    pandas.DataFrame: The updated DataFrame with the new data merged.
    """
    # Filter for the specific process
    times_df_filtered = df[times_df["TechName"] == process_name]

    # Fetch data from the API for the specific process
    API_URL = f"https://openenergy-platform.org/api/v0/schema/model_draft/tables/{process_name}/rows"
    data = fetch_data(API_URL)

    return times_df_filtered, data


# Paths and URLs
TIMES_FILE_PATH = "test_output.xlsx"

# Read the pickle file and print the DataFrame
PICKLE_FILE_PATH = "times_df.pkl"
times_df = pd.read_pickle(PICKLE_FILE_PATH)
print(times_df)

# Fetch and process data for a specific process
process = "ind_steel_blafu_0"
updated_df, api_data = data_mapping(times_df, process)

# format_and_save_excel(TIMES_FILE_PATH, times_df)
print(f"Excel file saved")
