import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


def process_data(original_df):
    technology_names = []
    comms_in = []
    comms_out = []
    attributes = []

    for _, row in original_df.iterrows():
        process = row["process"] if "process" in row else None
        inputs = str(row["input"]).split(",") if "input" in row else []
        outputs = str(row["output"]).split(",") if "output" in row else []

        for inp in inputs:
            technology_names.append(process)
            comms_in.append(inp.strip())
            comms_out.append(None)
            attributes.append("Input")

        for out in outputs:
            technology_names.append(process)
            comms_in.append(None)
            comms_out.append(out.strip())
            attributes.append("Output")

    data = {
        "Technology Name": technology_names,
        "TechDesc": ["" for _ in technology_names],
        "Attribute": attributes,
        "Comm-IN": [
            inp if attr == "Input" else "" for inp, attr in zip(comms_in, attributes)
        ],
        "Comm-OUT": [
            out if attr == "Output" else "" for out, attr in zip(comms_out, attributes)
        ],
    }

    return pd.DataFrame(data)


def format_and_save_excel(processed_df, file_path):
    wb = Workbook()
    ws = wb.active

    header_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    header_font = Font(bold=True, color="000000")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def style_cell(cell, fill=None, font=None, border=None, alignment=None):
        if fill:
            cell.fill = fill
        if font:
            cell.font = font
        if border:
            cell.border = border
        if alignment:
            cell.alignment = alignment

    headers = ["Technology Name", "*TechDesc", "Attribute", "Comm-IN", "Comm-OUT"]
    for col, header_title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header_title)
        style_cell(
            cell,
            fill=header_fill,
            font=header_font,
            border=thin_border,
            alignment=Alignment(horizontal="center"),
        )

    column_widths = [
        len(header) for header in headers
    ]  # Initialize with header lengths
    print(column_widths)

    for row_index, (idx, row) in enumerate(processed_df.iterrows(), start=2):
        for col_index, (col, value) in enumerate(row.items(), start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            style_cell(cell, border=thin_border)
            if value:  # Update the max length if the current value is longer
                column_widths[col_index - 1] = max(
                    column_widths[col_index - 1], len(str(value))
                )

    # Set column widths
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = (
            width + 2
        )  # +2 for a little extra padding

    wb.save(file_path)
    return file_path


def fetch_data_as_dataframe(url):
    """
    Fetches data from the specified URL and returns it as a pandas DataFrame.

    Parameters:
    - url (str): The URL from which to fetch the data.

    Returns:
    - DataFrame: A pandas DataFrame containing the fetched data.
    """
    # Fetch the data
    response = requests.get(url)
    data = response.json()

    # Convert the JSON data into a DataFrame
    df = pd.DataFrame(data)

    return df


# Load the original DataFrame
original_df = pd.read_excel("test_data.xlsx")

# Process the data
processed_df = process_data(original_df)

# Format and save the Excel file
file_path = "test_output.xlsx"
formatted_file_path = format_and_save_excel(processed_df, file_path)

print(f"Excel file saved as: {formatted_file_path}")

# url = "https://openenergy-platform.org/api/v0/schema/model_draft/tables/ind_steel_blafu_0/rows"
# df = fetch_data_as_dataframe(url)
# print(df.head())
