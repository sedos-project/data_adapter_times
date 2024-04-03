import pandas as pd
import requests
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


def process_data(original_df):
    technology_names = []
    comms_in = []
    comms_out = []
    attributes = []

    # Regex pattern to detect elements within square brackets
    bracket_pattern = re.compile(r"\[([^]]+)\]")

    for _, row in original_df.iterrows():
        process = row.get("process", None)

        # Check for NaN and convert to empty string if necessary
        input_str = str(row.get("input", "")) if pd.notna(row.get("input")) else ""
        output_str = str(row.get("output", "")) if pd.notna(row.get("output")) else ""

        # Find all bracketed items and process them separately
        bracketed_items = bracket_pattern.findall(input_str)
        for bracketed_item in bracketed_items:
            # Split each bracketed item by comma and strip whitespace
            items = [item.strip() for item in bracketed_item.split(",")]
            for item in items:
                technology_names.append(process)
                comms_in.append(item)
                comms_out.append(None)
                attributes.append("FLO_SHAR")

        # Remove bracketed items from the input string before splitting
        input_str = bracket_pattern.sub("", input_str)
        inputs = [i.strip() for i in input_str.split(",") if i.strip()]

        outputs = [o.strip() for o in output_str.split(",") if o.strip()]

        for inp in inputs:
            technology_names.append(process)
            comms_in.append(inp)
            comms_out.append(None)
            attributes.append("INPUT")

        for out in outputs:
            technology_names.append(process)
            comms_in.append(None)
            comms_out.append(out)
            attributes.append("OUTPUT")

    # Building the final DataFrame with specified columns
    data = {
        "TechName": technology_names,
        "TechDesc": ["" for _ in technology_names],
        "Attribute": [attr.upper() for attr in attributes],
        "Comm-IN": comms_in,
        "Comm-OUT": comms_out,
        "CommGrp": ["" for _ in technology_names],
        "TimeSlice": ["" for _ in technology_names],
        "LimType": ["" for _ in technology_names],
        "2021": ["" for _ in technology_names],
        "2024": ["" for _ in technology_names],
        "2027": ["" for _ in technology_names],
        "2030": ["" for _ in technology_names],
        "2035": ["" for _ in technology_names],
        "2040": ["" for _ in technology_names],
        "2045": ["" for _ in technology_names],
        "2050": ["" for _ in technology_names],
        "2060": ["" for _ in technology_names],
        "2070": ["" for _ in technology_names],
    }

    df = pd.DataFrame(data)

    # Sort the DataFrame by "TechName" and "Attribute"
    # Assigning a custom order for "Attribute"
    attribute_order = {"INPUT": 1, "OUTPUT": 2, "FLO_SHAR": 3}
    df["AttributeRank"] = df["Attribute"].map(attribute_order)
    df.sort_values(by=["TechName", "AttributeRank"], inplace=True)
    df.drop("AttributeRank", axis=1, inplace=True)

    return df


def format_and_save_excel(processed_df, file_path):
    """
    The format_and_save_excel function takes a dataframe and saves it to an excel file.
    The function also formats the excel file with headers, borders, and column widths.


    :param processed_df: Pass the dataframe to be saved
    :param file_path: Specify the path to save the file
    :return: The file_path
    :doc-author: Trelent
    """
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

    headers = [
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
print(processed_df)

# Format and save the Excel file
file_path = "test_output.xlsx"
formatted_file_path = format_and_save_excel(processed_df, file_path)

print(f"Excel file saved as: {formatted_file_path}")

# url = "https://openenergy-platform.org/api/v0/schema/model_draft/tables/ind_steel_blafu_0/rows"
# df = fetch_data_as_dataframe(url)
# print(df.head())
