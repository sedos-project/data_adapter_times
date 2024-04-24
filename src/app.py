import pandas as pd
import re
from openpyxl import Workbook, load_workbook
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


def add_comm_sheet_to_workbook(file_path, processed_df):
    # Load the existing workbook
    wb = load_workbook(file_path)

    # Create a new sheet
    ws_comm = wb.create_sheet("Commodities")

    # Define fills, fonts, borders, and alignment
    header_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    subheader_fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    blue_font = Font(color="0000FF", size=11, bold=True)
    header_font = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Set ~FI_Comm title
    ws_comm["B1"] = "~FI_Comm"
    ws_comm["B1"].font = blue_font
    ws_comm["B1"].alignment = align_center

    subheader_font = Font(color="000000")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Define headers and subheaders
    headers = [
        "Csets",
        "CommName",
        "CommDesc",
        "Unit",
        "LimType",
        "CTSLvl",
        "PeakTS",
        "Ctype",
    ]

    subheaders = [
        "I: Commodity Set Membership",
        "Commodity Name",
        "Commodity Description",
        "Unit",
        "Balance Equ Type Override",
        "Timeslice Tracking Level",
        "Peak Monitoring",
        "Electricity Indicator",
    ]

    # Write headers and subheaders
    for col, header in enumerate(headers, start=2):  # Start from the second column
        cell = ws_comm.cell(row=2, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = align_center

    for col, subheader in enumerate(
        subheaders, start=2
    ):  # Start from the second column
        cell = ws_comm.cell(row=3, column=col)
        cell.value = subheader
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.border = thin_border
        cell.alignment = align_center

    # Load the commodity_set data from mapping.xlsx
    wb_mapping = load_workbook("mapping.xlsx", data_only=True)
    ws_mapping = wb_mapping["commodity_set"]

    # Extract the commodities into a set for faster membership testing
    mat_commodities = set(
        row[0].lower().strip()
        for row in ws_mapping.iter_rows(min_row=2, max_col=1, values_only=True)
        if row[0]
    )

    # Determine the commodity set membership
    commodity_sets = {}
    for commodity in set(
        processed_df["Comm-IN"].dropna().unique().tolist()
        + processed_df["Comm-OUT"].dropna().unique().tolist()
    ):
        commodity_lower = commodity.lower()
        if "exo" in commodity_lower:
            commodity_sets[commodity] = "DEM"
        elif "emi" in commodity_lower:
            commodity_sets[commodity] = "ENV"
        elif commodity_lower in mat_commodities:
            commodity_sets[commodity] = "MAT"
        else:
            commodity_sets[commodity] = "NRG"

    # Populate the CommName column with unique commodities and set membership
    for row_idx, comm in enumerate(
        commodity_sets.keys(), start=4
    ):  # Start from the fourth row
        ws_comm.cell(row=row_idx, column=2, value=commodity_sets[comm])  # Csets
        ws_comm.cell(row=row_idx, column=3, value=comm)  # CommName

    # Adjust column widths
    for col in ws_comm.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length + 1  # Add 1 padding and adjust for a better fit
        ws_comm.column_dimensions[column].width = adjusted_width

    # Save the workbook
    wb.save(file_path)


def add_process_sheet_to_workbook(file_path, processed_df):
    # Load the existing workbook
    wb = load_workbook(file_path)

    # Create a new sheet
    ws_process = wb.create_sheet("Processes")

    # Define fills, fonts, borders, and alignment for headers and subheaders
    header_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    subheader_fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    blue_font = Font(color="0000FF", size=11, bold=True)
    header_font = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # Set ~FI_Process title
    ws_process["B1"] = "~FI_Process"
    ws_process["B1"].font = blue_font
    ws_process["B1"].alignment = align_center

    subheader_font = Font(color="000000")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Define headers and subheaders
    headers = [
        "Sets",
        "TechName",
        "TechDesc",
        "Tact",
        "Tcap",
        "Tslvl",
        "PrimaryCG",
        "Vintage",
    ]

    subheaders = [
        "I: Process Set Membership",
        "Technology Name",
        "Technology Description",
        "Activity Unit",
        "Capacity Unit",
        "Timeslice Operational Level",
        "Operational Commodity Group",
        "Vintage Tracking",
    ]

    # Write headers and subheaders
    for col, header in enumerate(headers, start=2):  # Start from the second column
        cell = ws_process.cell(row=2, column=col)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = align_center

    for col, subheader in enumerate(
        subheaders, start=2
    ):  # Start from the second column
        cell = ws_process.cell(row=3, column=col)
        cell.value = subheader
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.border = thin_border
        cell.alignment = align_center

    # Determine the process set membership
    process_sets = {}
    for _, row in processed_df.iterrows():
        tech_name = row["TechName"]
        output_commodities = row["Comm-OUT"]
        if "chp" in tech_name.lower():
            process_sets[tech_name] = "CHP"
        elif pd.notna(output_commodities) and "exo" in output_commodities.lower():
            process_sets[tech_name] = "DEM"
        else:
            process_sets.setdefault(tech_name, "PRE")

    # Populate the process sheet
    for row_idx, tech_name in enumerate(process_sets.keys(), start=4):
        ws_process.cell(row=row_idx, column=2, value=process_sets[tech_name])  # Sets
        ws_process.cell(row=row_idx, column=3, value=tech_name)  # TechName

    # Adjust column widths
    for col in ws_process.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length + 1  # Add padding and adjust for a better fit
        ws_process.column_dimensions[column].width = adjusted_width

    # Save the workbook
    wb.save(file_path)


def format_and_save_excel(processed_df, file_path):
    """
    The format_and_save_excel function takes a dataframe and saves it as an Excel file.
    The function also formats the Excel file with headers, subheaders, column widths, etc.

    :param processed_df: Pass the processed dataframe to the function
    :param file_path: Specify the location where the excel file will be saved
    :return: The file_path
    :doc-author: Trelent
    """
    wb = Workbook()
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


# Load the original DataFrame
original_df = pd.read_excel("test_data.xlsx")

# Process the data
processed_df = process_data(original_df)
print(processed_df)

# Format and save the Excel file
file_path = "test_output.xlsx"
formatted_file_path = format_and_save_excel(processed_df, file_path)

add_comm_sheet_to_workbook("test_output.xlsx", processed_df)
add_process_sheet_to_workbook("test_output.xlsx", processed_df)

print(f"Excel file saved as: {formatted_file_path}")
