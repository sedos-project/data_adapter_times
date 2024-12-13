import pandas as pd
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from unit_conversion import convert_unit

# Initialize the counter
fetch_data_counter = 0
units_mapping = {}
updated_units_mapping = {}
desired_units_mapping = {}


def load_desired_units_mapping(file_path="config_data/units_mapping.xlsx"):
    """
    Loads the source-desired unit pairs from the 'Unique Units' sheet into a dictionary.

    Parameters:
    file_path (str): The path to the units_mapping Excel file.

    Returns:
    dict: A dictionary where keys are source units, and values are desired units.
    """
    wb = load_workbook(file_path, data_only=True)
    if "Unique Units" not in wb.sheetnames:
        print("Unique Units sheet not found in the Excel file.")
        return {}

    ws = wb["Unique Units"]
    units_dict = {}

    for row in ws.iter_rows(min_row=2, values_only=True):  # Skip header row
        unit, desired_unit = row
        if unit and desired_unit:
            units_dict[unit] = desired_unit

    return units_dict


def format_and_save_excel(file_path, processed_df):
    """
    Format and save the processed data into an Excel file.

    Parameters:
    processed_df (pandas.DataFrame): The DataFrame containing the processed data.
    file_path (str): The path where the Excel file will be saved.

    Returns:
    str: The path where the Excel file is saved.
    """

    # Replace pd.NA with empty strings
    processed_df = processed_df.fillna("")

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
            # Convert empty lists to empty strings
            if isinstance(value, list) and not value:
                value = ""
            elif isinstance(value, list):
                value = ", ".join(map(str, value))
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


def data_units(metadata):
    """
    Converts the units of the fields in the API data according to the metadata.

    Parameters:
    metadata (dict): The metadata containing information about the units.

    Returns:
    dict: A dictionary where the keys are resource names and the values are lists of field names and units.
    """
    units_dict = {}

    if not metadata:
        return units_dict

    for resource in metadata.get("resources", []):
        resource_name = resource.get("name", "").replace(
            "model_draft.", ""
        )  # Remove "model_draft." prefix
        fields = resource.get("schema", {}).get("fields", [])

        # Initialize the list for each resource if it doesn't exist
        if resource_name not in units_dict:
            units_dict[resource_name] = []

        # Append each field and its unit as a dictionary to the resource's list
        for field in fields:
            field_name = field["name"]
            field_unit = field.get("unit")
            if field_name and field_unit:
                units_dict[resource_name].append(
                    {"field_name": field_name, "unit": field_unit}
                )

    return units_dict


def update_commodity_list_units(excel_file_path, units_mapping):
    """
    Update the 'Commodity List' sheet in the Excel file with the units from units_mapping.
    """
    from openpyxl import load_workbook

    # Load the Excel file
    wb = load_workbook(excel_file_path)
    if "Commodity List" in wb.sheetnames:
        ws = wb["Commodity List"]
    else:
        print("Commodity List sheet not found in the Excel file.")
        return

    header_row = find_header_row(ws, "CommName")
    if header_row is None:
        print("CommName header not found in Commodity List sheet.")
        return

    # Get column indices for CommName and Unit
    headers = {cell.value: cell.column for cell in ws[header_row]}
    if "CommName" in headers and "Unit" in headers:
        commname_col = headers["CommName"]
        unit_col = headers["Unit"]
    else:
        print("CommName or Unit column not found in Commodity List sheet.")
        return

    # Loop over each resource in units_mapping
    for resource_name, fields in units_mapping.items():
        for field in fields:
            field_name = field["field_name"]
            field_unit = field["unit"]
            if field_name.startswith("conversion_factor_"):
                comm_name = field_name[len("conversion_factor_") :]
                found = False
                for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
                    commname_cell = row[commname_col - 1]
                    if commname_cell.value and isinstance(commname_cell.value, str):
                        if (
                            commname_cell.value.strip().lower()
                            == comm_name.strip().lower()
                        ):
                            unit_cell = row[unit_col - 1]
                            unit_cell.value = field_unit
                            found = True
                            break
                if not found:
                    print(
                        f"CommName '{comm_name}' with unit '{field_unit}' not found in Commodity List sheet."
                    )

    # Save the workbook
    wb.save(excel_file_path)
    print("Commodity List sheet updated with units.")


def update_process_list_sheet(excel_file_path, units_mapping):
    """
    Update the 'Process List' sheet in the Excel file with specific values:
    - Set 'Vintage' to 'NO' for all rows.
    - Set 'PrimaryCG' to the first output commodity for each process.
    - Set 'Tact' to the unit of the first output commodity, as found in 'Commodity List' sheet.
    """
    # Load the workbook and sheets
    wb = load_workbook(excel_file_path)
    if "Process List" not in wb.sheetnames or "Commodity List" not in wb.sheetnames:
        print("Required sheets not found in the Excel file.")
        return

    ws_process_list = wb["Process List"]
    ws_commodity_list = wb["Commodity List"]

    # Identify columns in 'Process List' sheet
    header_row = find_header_row(ws_process_list, "TechName")
    headers = {cell.value: cell.column for cell in ws_process_list[header_row]}

    techname_col = headers.get("TechName")
    vintage_col = headers.get("Vintage")
    primary_cg_col = headers.get("PrimaryCG")
    tact_col = headers.get("Tact")

    if not (techname_col and vintage_col and primary_cg_col and tact_col):
        print("Some required columns are missing in 'Process List' sheet.")
        return

    # Set Vintage column to 'NO' and populate PrimaryCG and Tact
    for row in ws_process_list.iter_rows(min_row=header_row + 1, values_only=False):
        techname_cell = row[techname_col - 1]
        if techname_cell.value:
            process_name = techname_cell.value.strip()

            # Set Vintage to 'NO'
            row[vintage_col - 1].value = "NO"

            # Set PrimaryCG to first output commodity and Tact to its unit from Commodity List
            output_commodity = (
                updated_df.loc[
                    (updated_df["TechName"] == process_name)
                    & (updated_df["Attribute"] == "OUTPUT"),
                    "Comm-OUT",
                ]
                .dropna()
                .values
            )

            if output_commodity.size > 0:  # Check if array is not empty
                if "exo_" in output_commodity[0]:
                    primary_cg = "DEMO"
                else:
                    primary_cg = output_commodity[0]
                row[primary_cg_col - 1].value = primary_cg

                # Find the Tact unit by searching the resource in units_mapping
                for resource_name, fields in units_mapping.items():
                    for field in fields:
                        if field["field_name"] == f"conversion_factor_{primary_cg}":
                            row[tact_col - 1].value = field["unit"]
                            break

    wb.save(excel_file_path)
    print("Process List sheet updated with Vintage, PrimaryCG, and Tact columns.")


def get_desired_unit_from_excel(
    source_unit, file_path="output_data/units_mapping.xlsx"
):
    """
    Reads the Unique Units sheet in the units_mapping.xlsx file to find the desired unit
    corresponding to a given source unit.

    Parameters:
    source_unit (str): The unit to convert from.
    file_path (str): The path to the units_mapping Excel file.

    Returns:
    str: The desired unit, if found; otherwise, returns None.
    """
    wb = load_workbook(file_path, data_only=True)
    if "Unique Units" not in wb.sheetnames:
        print("Unique Units sheet not found in the Excel file.")
        return None

    ws = wb["Unique Units"]

    for row in ws.iter_rows(min_row=2, values_only=True):  # Skip header row
        unit, desired_unit = row
        if unit == source_unit:
            return desired_unit
    return None


def fetch_data(url, process_name):
    global fetch_data_counter
    fetch_data_counter += 1  # Increment the counter
    try:
        response = requests.get(url)
        if response.status_code == 200:
            print(
                f"Data fetched successfully for process {fetch_data_counter}: {process_name}"
            )
            return pd.DataFrame(response.json())
        else:
            print(
                f"Failed to fetch data for process {fetch_data_counter}: {process_name}, status code: {response.status_code}"
            )
            return pd.DataFrame()  # Return an empty DataFrame if status code is not 200
    except requests.RequestException as e:
        print(
            f"No data found for process {fetch_data_counter}: {process_name}, error: {e}"
        )
        return pd.DataFrame()  # Return an empty DataFrame in case of error


def fetch_process_metadata(process):
    """
    Fetches the metadata of a process and extracts the units for all the resources.

    Parameters:
    process_name (str): The name of the process to fetch metadata for.

    Returns:
    dict: A dictionary where the keys are the resource names and the values are their corresponding units.
    """
    try:
        url = f"https://openenergy-platform.org/api/v0/schema/model_draft/tables/{process}/meta/"
        response = requests.get(url)
        if response.status_code == 200:
            metadata = response.json()
            # print(f"Units fetched successfully for process: {process}")
            return metadata
        else:
            # print(
            #     f"Failed to fetch metadata for process: {process}, status code: {response.status_code}"
            # )
            return {}
    except requests.RequestException as e:
        print(f"Error fetching metadata for process: {process}, error: {e}")
        return {}


def find_header_row(sheet, header_name):
    """
    This function finds the row number of the first occurrence of a specific header in a given sheet.

    Parameters:
    sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet where the header is to be found.
    header_name (str): The name of the header to be found.

    Returns:
    int: The row number of the first occurrence of the header. If the header is not found, a ValueError is raised.

    Raises:
    ValueError: If the header is not found within the first 20 rows.
    """
    for row in range(1, 20):  # Assume headers are within the first 10 rows
        for col in range(1, sheet.max_column + 1):
            cell_value = str(sheet.cell(row=row, column=col).value)
            if header_name.lower() in cell_value.lower():
                return row
    raise ValueError("Header row not found within the first 10 rows.")


def data_mapping(times_df, process_name, is_group=False):
    """
    Fetches data from the API for a given process name or group and updates the times_df DataFrame.

    Parameters:
    times_df (pandas.DataFrame): The DataFrame containing the initial data.
    process_name (str): The name of the process or process group to fetch and process data for.
    is_group (bool): Flag indicating whether the process_name is a group.

    Returns:
    pandas.DataFrame: The updated DataFrame with the new data merged.
    """
    api_process_data = fetch_data(
        f"https://openenergy-platform.org/api/v0/schema/model_draft/tables/{process_name}/rows",
        process_name,
    )

    if api_process_data.empty:
        return times_df  # Return the original DataFrame if no data is fetched

    if is_group:
        # Divide the data based on the 'type' column
        process_groups = api_process_data.groupby("type")

        process_count = 0  # Initialize a counter for the processes handled

        for process, group_data in process_groups:
            if process.endswith("_ag"):  # Skip processes ending with _ag
                continue
            handled_processes.append(process)
            times_df = data_mapping_internal(
                times_df, process, group_data
            )  # Call internal function for each process
            process_count += 1  # Increment the counter for each handled process

        print(
            f"{process_count} processes were handled inside the process group: {process_name}"
        )
        return times_df
    else:
        return data_mapping_internal(times_df, process_name, api_process_data)


def data_mapping_internal(times_df, process_name, api_process_data):

    # Fetch metadata
    metadata = fetch_process_metadata(process_name)

    # Update units_mapping
    global units_mapping, updated_units_mapping
    units_mapping.update(data_units(metadata))
    updated_units_mapping.update(data_units(metadata))

    # Filter for the specific process and keep track of the index range
    times_df_filtered = times_df[times_df["TechName"] == process_name]
    if times_df_filtered.empty:
        print(f"{process_name} was not found in the SEDOS input and hence was skipped")
        return times_df  # Skip if there is no matching process

    start_idx = times_df.index.get_loc(times_df_filtered.index[0])
    end_idx = times_df.index.get_loc(times_df_filtered.index[-1])

    # Load the mapping file
    mapping_file_path = "config_data/mapping_v4.xlsx"
    wb = load_workbook(mapping_file_path, data_only=True)
    sheet = wb["SEDOS_parameters"]

    # Find the header row for 'SEDOS'
    header_row = find_header_row(sheet, "SEDOS")

    # Extract the SEDOS, TIMES, and Constraints columns
    sedos_list = []
    times_list = []
    constraints_list = []

    for row in sheet.iter_rows(min_row=header_row + 1, max_row=sheet.max_row):
        sedos_value = row[0].value  # Assuming SEDOS is in the first column
        times_value = row[1].value  # Assuming TIMES is in the second column
        constraints_value = row[8].value  # Assuming Constraints is in the ninth column
        if sedos_value and times_value:
            sedos_list.append(sedos_value)
            times_list.append(times_value)
            constraints_list.append(
                constraints_value if constraints_value is not None else ""
            )

    # Modify the SEDOS list items
    sedos_list = [item.split("<")[0].lower().strip() for item in sedos_list]

    # Create a mapping dictionary
    mapping_dict = dict(zip(sedos_list, times_list))
    constraints_dict = dict(zip(sedos_list, constraints_list))

    # Create a dictionary for matched SEDOS items and API column names
    matched_columns = {}

    for sedos_item in sedos_list:
        for api_col in api_process_data.columns:
            if sedos_item in api_col.lower():
                if sedos_item not in matched_columns:
                    matched_columns[sedos_item] = []
                matched_columns[sedos_item].append(api_col)

    # Add the TIMES list items and constraints corresponding to the matched SEDOS items
    extended_matched_columns = {
        sedos_item: (api_cols, mapping_dict[sedos_item], constraints_dict[sedos_item])
        for sedos_item, api_cols in matched_columns.items()
    }

    # Update the times_df_filtered with the api_process_data based on the matched columns
    for sedos_item, (
        api_cols,
        times_col,
        constraint,
    ) in extended_matched_columns.items():
        for api_col in api_cols:
            if api_col in api_process_data.columns:
                # Extract the values and year from the API data
                api_values = api_process_data[api_col]
                years = api_process_data["year"]
                comm_col_value = api_col.replace("conversion_factor_", "")
                for api_value, year in zip(api_values, years):
                    # Find the column in times_df_filtered that matches the year
                    if str(year) in times_df_filtered.columns:
                        # Check if sedos_item contains 'conversion_factor_'
                        if "conversion_factor_" in sedos_item:
                            # Check if both the Attribute and Comm-IN/Comm-OUT match
                            matching_row = times_df_filtered[
                                (
                                    (times_df_filtered["Attribute"] == times_col)
                                    | (times_df_filtered["Attribute"] == "OUTPUT")
                                    | (times_df_filtered["Attribute"] == "INPUT")
                                )
                                & (
                                    (times_df_filtered["Comm-IN"] == comm_col_value)
                                    | (times_df_filtered["Comm-OUT"] == comm_col_value)
                                )
                            ]
                            if not matching_row.empty:
                                # Fetch unit from units_mapping by matching resource_name to process_name
                                source_unit = None
                                for resource_name, fields in units_mapping.items():
                                    if resource_name == process_name:
                                        for field in fields:
                                            if (
                                                field["field_name"]
                                                == f"conversion_factor_{comm_col_value}"
                                            ):
                                                source_unit = field["unit"]
                                                break
                                        if source_unit:
                                            break

                                # If source unit is found, fetch the desired unit
                                if source_unit:
                                    desired_unit = desired_units_mapping.get(
                                        source_unit, None
                                    )
                                    # Only proceed with conversion if the desired unit is found
                                    if desired_unit:
                                        converted_value, conversion_flag = convert_unit(
                                            api_value, source_unit, to_unit=desired_unit
                                        )

                                        if conversion_flag == 1:
                                            for (
                                                resource
                                            ) in updated_units_mapping.values():
                                                for field in resource:
                                                    if (
                                                        field["field_name"]
                                                        == f"conversion_factor_{comm_col_value}"
                                                    ):
                                                        field["unit"] = desired_unit
                                    else:
                                        print(
                                            f"Desired unit for {source_unit} for {api_col} not found."
                                        )
                                        converted_value = api_value
                                else:
                                    converted_value = api_value  # If no unit is found, use the value as is

                                for idx in matching_row.index:
                                    if api_value is not None:
                                        times_df_filtered.at[idx, str(year)] = (
                                            converted_value
                                        )
                            else:
                                matching_row = times_df_filtered[
                                    (times_df_filtered["Attribute"] == "ACT_EFF")
                                ]
                                if not matching_row.empty:
                                    # Fetch unit from units_mapping by matching resource_name to process_name
                                    source_unit = None
                                    for resource_name, fields in units_mapping.items():
                                        if resource_name == process_name:
                                            for field in fields:
                                                if field["field_name"] == api_col:
                                                    source_unit = field["unit"]
                                                    break
                                            if source_unit:
                                                break

                                    # If source unit is found, fetch the desired unit
                                    if source_unit:
                                        desired_unit = desired_units_mapping.get(
                                            source_unit, None
                                        )
                                        # Only proceed with conversion if the desired unit is found
                                        if desired_unit:
                                            converted_value, conversion_flag = (
                                                convert_unit(
                                                    api_value,
                                                    source_unit,
                                                    to_unit=desired_unit,
                                                )
                                            )

                                            if conversion_flag == 1:
                                                for (
                                                    resource
                                                ) in updated_units_mapping.values():
                                                    for field in resource:
                                                        if (
                                                            field["field_name"]
                                                            == api_col
                                                        ):
                                                            field["unit"] = desired_unit
                                        else:
                                            print(
                                                f"Desired unit for {source_unit} for {api_col} not found."
                                            )
                                            converted_value = api_value
                                    else:
                                        converted_value = api_value  # If no unit is found, use the value as is

                                    for idx in matching_row.index:
                                        if api_value is not None:
                                            times_df_filtered.at[idx, str(year)] = (
                                                converted_value
                                            )
                        elif "flow_share" in sedos_item:
                            # Add flow share values
                            matching_row = times_df_filtered[
                                times_df_filtered["Attribute"] == times_col
                            ]
                            if not matching_row.empty:
                                comm_in_out_values = []
                                for idx in matching_row.index:
                                    comm_in_out_values.extend(
                                        [
                                            times_df_filtered.at[idx, "Comm-IN"],
                                            times_df_filtered.at[idx, "Comm-OUT"],
                                        ]
                                    )
                                comm_in_out_values = list(
                                    filter(pd.notna, comm_in_out_values)
                                )
                                # Parse the current api_col to get flow share commodity
                                flow_share_commodity = api_col.replace(
                                    sedos_item, ""
                                ).strip("_")

                                # Add values to the matching rows
                                sum_of_matched_values = 0
                                for idx in matching_row.index:
                                    comm_in = times_df_filtered.at[idx, "Comm-IN"]
                                    comm_out = times_df_filtered.at[idx, "Comm-OUT"]
                                    if flow_share_commodity in (comm_in, comm_out):
                                        if api_value is not None:
                                            times_df_filtered.at[idx, str(year)] = (
                                                api_value / 100
                                            )
                                            times_df_filtered.at[idx, "LimType"] = (
                                                constraint
                                            )
                                            sum_of_matched_values += api_value / 100

                                # Handle the rows that do not match the flow share commodity
                                for idx in matching_row.index:
                                    if flow_share_commodity not in (
                                        times_df_filtered.at[idx, "Comm-IN"],
                                        times_df_filtered.at[idx, "Comm-OUT"],
                                    ):
                                        times_df_filtered.at[idx, str(year)] = (
                                            1 - sum_of_matched_values
                                        )
                                        times_df_filtered.at[idx, "LimType"] = (
                                            constraint
                                        )
                        elif (
                            "availability_constant" in sedos_item
                            or "efficiency_sto_in" in sedos_item
                        ):
                            # Handle availability constants or time series fixed
                            matching_row = times_df_filtered[
                                times_df_filtered["Attribute"] == times_col
                            ]
                            if matching_row.empty:
                                # Add a new row if the Attribute does not exist
                                new_row = pd.Series(
                                    {col: pd.NA for col in times_df_filtered.columns}
                                )
                                new_row["TechName"] = process_name
                                new_row["Attribute"] = times_col
                                new_row["LimType"] = constraint
                                times_df_filtered = pd.concat(
                                    [times_df_filtered, new_row.to_frame().T],
                                    ignore_index=True,
                                )
                                new_row_idx = times_df_filtered[
                                    times_df_filtered["Attribute"] == times_col
                                ].index[-1]
                                if api_value is not None:
                                    times_df_filtered.at[new_row_idx, str(year)] = (
                                        api_value / 100
                                    )
                            else:
                                for idx in matching_row.index:
                                    if api_value is not None:
                                        times_df_filtered.at[idx, str(year)] = (
                                            api_value / 100
                                        )
                                        times_df_filtered.at[idx, "LimType"] = (
                                            constraint
                                        )
                        elif (
                            "availability_timeseries_fixed" in sedos_item
                            or "availability_timeseries_max" in sedos_item
                        ):
                            # temporary fix
                            continue
                        else:
                            # Check if only the Attribute matches
                            matching_row = times_df_filtered[
                                times_df_filtered["Attribute"] == times_col
                            ]
                            if matching_row.empty:
                                # Add a new row if the Attribute does not exist
                                new_row = pd.Series(
                                    {col: pd.NA for col in times_df_filtered.columns}
                                )
                                new_row["TechName"] = process_name
                                new_row["Attribute"] = times_col
                                new_row["LimType"] = constraint
                                times_df_filtered = pd.concat(
                                    [times_df_filtered, new_row.to_frame().T],
                                    ignore_index=True,
                                )
                                new_row_idx = times_df_filtered[
                                    times_df_filtered["Attribute"] == times_col
                                ].index[-1]
                                if api_value is not None:
                                    # If the sedos_item contains 'cb_coefficient', apply 1/api_value
                                    if "cb_coefficient" in sedos_item:
                                        times_df_filtered.at[new_row_idx, str(year)] = (
                                            1 / api_value
                                        )
                                    else:
                                        # For all other cases, just use the api_value directly
                                        times_df_filtered.at[new_row_idx, str(year)] = (
                                            api_value
                                        )
                            else:
                                for idx in matching_row.index:
                                    if api_value is not None:
                                        # Fetch unit from units_mapping by matching resource_name to process_name
                                        source_unit = None
                                        for (
                                            resource_name,
                                            fields,
                                        ) in units_mapping.items():
                                            if resource_name == process_name:
                                                for field in fields:
                                                    if field["field_name"] == api_col:
                                                        source_unit = field["unit"]
                                                        break
                                                if source_unit:
                                                    break

                                        # If source unit is found, fetch the desired unit
                                        if source_unit:
                                            desired_unit = desired_units_mapping.get(
                                                source_unit, None
                                            )
                                            # Only proceed with conversion if the desired unit is found
                                            if desired_unit:
                                                converted_value, conversion_flag = (
                                                    convert_unit(
                                                        api_value,
                                                        source_unit,
                                                        to_unit=desired_unit,
                                                    )
                                                )

                                                if conversion_flag == 1:
                                                    for (
                                                        resource
                                                    ) in updated_units_mapping.values():
                                                        for field in resource:
                                                            if (
                                                                field["field_name"]
                                                                == api_col
                                                            ):
                                                                field["unit"] = (
                                                                    desired_unit
                                                                )
                                            else:
                                                print(
                                                    f"Desired unit for {source_unit} for {api_col} not found."
                                                )
                                                converted_value = api_value
                                        else:
                                            converted_value = api_value  # If no unit is found, use the value as is
                                        times_df_filtered.at[idx, str(year)] = (
                                            converted_value
                                        )
                                        times_df_filtered.at[idx, "LimType"] = (
                                            constraint
                                        )

    # Implement CAP2ACT logic
    # Check if any output commodities contain "exo_"
    if any(
        isinstance(comm_out, str) and "exo_" in comm_out.lower()
        for comm_out in times_df_filtered["Comm-OUT"]
    ):
        if (
            process_name == "tra_road_const_ice_diesel_0"
            or process_name == "tra_road_agri_ice_diesel_1"
            or process_name == "tra_road_agri_ice_diesel_0"
            or process_name == "tra_road_const_ice_diesel_1"
        ):
            cap2act_value = 1
        else:
            cap2act_value = (
                0.000000001  # Set CAP2ACT to 0.001 if "exo_" is in any output commodity
            )

    # Check if the process name contains "battery"
    elif "battery" in process_name.lower():
        cap2act_value = (
            0.0036  # Set CAP2ACT to 0.0036 if process name contains "battery"
        )

    else:
        cap2act_value = 1  # Default CAP2ACT value

    # Add CAP2ACT as a new row in times_df_filtered
    cap2act_row = pd.Series(
        {
            "TechName": process_name,
            "Attribute": "CAP2ACT",
            "LimType": pd.NA,
            "2021": cap2act_value,
            "2024": cap2act_value,
            "2027": cap2act_value,
            "2030": cap2act_value,
            "2035": cap2act_value,
            "2040": cap2act_value,
            "2045": cap2act_value,
            "2050": cap2act_value,
            "2060": cap2act_value,
            "2070": cap2act_value,
        }
    )
    times_df_filtered = pd.concat(
        [times_df_filtered, cap2act_row.to_frame().T], ignore_index=True
    )

    # Replace <NA> with empty strings before updating the original times_df
    with pd.option_context("future.no_silent_downcasting", True):
        times_df_filtered = times_df_filtered.fillna("")

    # Ensure the updated times_df_filtered has the same or larger index range
    if len(times_df_filtered) > (end_idx - start_idx + 1):
        # Split the original times_df into three parts
        before = times_df.iloc[:start_idx]
        after = times_df.iloc[end_idx + 1 :]

        # Concatenate the before part, updated times_df_filtered, and the after part
        times_df = pd.concat([before, times_df_filtered, after], ignore_index=True)
    else:
        times_df.iloc[start_idx : end_idx + 1] = times_df_filtered.values

    return times_df


def calculate_act_eff(times_df, process_list_file_path):
    """
    Calculates the ACT_EFF attribute for processes with 'DEM' in their Sets column.
    For 'tra_road' processes, calculates ACT_EFF directly as input_value / 1000000000.
    For other processes, performs additional checks and calculations if 'INPUT' or 'ACTFLO~DEMO' are present.

    Parameters:
    times_df (pandas.DataFrame): The DataFrame containing the TIMES data.
    process_list_file_path (str): The path to the Excel file containing the 'Process List' sheet.

    Returns:
    pandas.DataFrame: The updated DataFrame with the 'ACT_EFF' attributes calculated.
    """
    from openpyxl import load_workbook

    # Read the Excel file and 'Process List' sheet
    wb = load_workbook(process_list_file_path)
    if "Process List" in wb.sheetnames:
        ws = wb["Process List"]
    else:
        print("Process List sheet not found in the Excel file.")
        return times_df  # Return original times_df if sheet not found

    # Find the header row containing 'TechName' and 'Sets'
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value == "TechName":
                header_row = cell.row
                break
        if header_row is not None:
            break

    if header_row is None:
        print("TechName header not found in Process List sheet.")
        return times_df

    # Get column indices for TechName and Sets
    headers = {cell.value: idx for idx, cell in enumerate(ws[header_row], start=1)}
    if "TechName" in headers and "Sets" in headers:
        techname_col = headers["TechName"]
        sets_col = headers["Sets"]
    else:
        print("TechName or Sets column not found in Process List sheet.")
        return times_df

    # For every TechName where Sets column has 'DEM' in it
    dem_technames = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        techname_cell = row[techname_col - 1]
        sets_cell = row[sets_col - 1]
        if techname_cell.value and sets_cell.value:
            if "DEM" in str(sets_cell.value):
                dem_technames.append(techname_cell.value)

    # Collect process positions
    process_positions = []
    for process_name in dem_technames:
        # Get the subset of times_df for this process_name
        times_df_filtered = times_df[times_df["TechName"] == process_name]

        if times_df_filtered.empty:
            print(f"{process_name} was not found in times_df and hence was skipped")
            continue  # Skip if no matching process

        # Find the index positions in times_df where this process's data is located
        process_indices = times_df_filtered.index
        start_index = process_indices[0]
        end_index = process_indices[-1]

        # Add to list
        process_positions.append((start_index, end_index, process_name))

    # Sort the process_positions list in descending order of start_index
    process_positions.sort(reverse=True)

    # Now process each process in reverse order
    for start_index, end_index, process_name in process_positions:
        times_df_filtered = times_df.loc[start_index:end_index]

        if (
            process_name == "tra_road_const_ice_diesel_0"
            or process_name == "tra_road_agri_ice_diesel_1"
            or process_name == "tra_road_agri_ice_diesel_0"
            or process_name == "tra_road_const_ice_diesel_1"
        ):
            continue
        else:
            # Check if the process is 'tra_road' type
            if process_name.startswith("tra_road"):
                # For 'tra_road' processes, check for 'INPUT' or 'ACT_EFF' rows
                input_rows = times_df_filtered[
                    times_df_filtered["Attribute"] == "INPUT"
                ]
                act_eff_rows = times_df_filtered[
                    times_df_filtered["Attribute"] == "ACT_EFF"
                ]
                flo_shar_rows = times_df_filtered[
                    times_df_filtered["Attribute"] == "FLO_SHAR"
                ]

                # New condition for processes containing '_hyb_'
                if "_hyb_" in process_name and not input_rows.empty:
                    act_eff_values_list = (
                        []
                    )  # List to store act_eff values for each input row
                    for _, input_row in input_rows.iterrows():
                        act_eff_values = {}
                        for year in [
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
                        ]:
                            try:
                                input_value = float(input_row[year])
                                act_eff_values[year] = input_value / 1000000000
                            except (ValueError, ZeroDivisionError, KeyError, TypeError):
                                act_eff_values[year] = ""
                        act_eff_values_list.append(act_eff_values)

                    # Replace the input rows' values with CEFF and calculated act_eff values
                    for idx, input_row in enumerate(input_rows.index):
                        for year, value in act_eff_values_list[idx].items():
                            times_df.at[input_row, year] = value
                        times_df.at[input_row, "Attribute"] = (
                            "CEFF"  # Change Attribute to CEFF
                        )

                    continue  # Skip further processing for '_hyb_' processes

                # Existing 'tra_road' condition for processes without '_hyb_'
                if (
                    input_rows.empty
                    and not act_eff_rows.empty
                    and not flo_shar_rows.empty
                ):
                    act_eff_row = act_eff_rows.iloc[0]
                    input_row = act_eff_row
                elif not input_rows.empty:
                    input_row = input_rows.iloc[0]
                else:
                    print(
                        f"No 'INPUT' or 'ACT_EFF' found for tra_road process {process_name}, skipping."
                    )
                    continue  # Skip if no input or act_eff data

                act_eff_values = {}
                for year in [
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
                ]:
                    try:
                        input_value = float(input_row[year])
                        act_eff_values[year] = input_value / 1000000000
                    except (ValueError, ZeroDivisionError, KeyError, TypeError):
                        act_eff_values[year] = ""

                if act_eff_rows.empty:
                    # Create a new EFF row if ACT_EFF was not originally present
                    new_row = {col: "" for col in times_df.columns}
                    new_row["TechName"] = process_name
                    new_row["Attribute"] = "EFF"
                    for year, value in act_eff_values.items():
                        new_row[year] = value
                    new_row_df = pd.DataFrame([new_row])
                    times_df = pd.concat(
                        [
                            times_df.iloc[: end_index + 1],
                            new_row_df,
                            times_df.iloc[end_index + 1 :],
                        ]
                    ).reset_index(drop=True)
                else:
                    # Update the existing ACT_EFF (now 'EFF') row with calculated values
                    for year, value in act_eff_values.items():
                        times_df.at[input_row.name, year] = value

                # Clear yearly data in rows with 'INPUT' and 'OUTPUT' attributes for tra_road
                for attr in ["INPUT", "OUTPUT"]:
                    attr_rows = times_df_filtered[
                        times_df_filtered["Attribute"] == attr
                    ]
                    for idx in attr_rows.index:
                        times_df.loc[
                            idx,
                            [
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
                            ],
                        ] = ""

                continue  # Skip further processing for tra_road

            # For other processes, proceed with the existing checks
            # Find the row which has 'Comm-OUT' starting with 'exo_'
            exo_rows = times_df_filtered[
                times_df_filtered["Comm-OUT"].astype(str).str.startswith("exo_")
            ]
            if exo_rows.empty:
                print(f"No 'exo_' in 'Comm-OUT' for process {process_name}")
                continue

            exo_row = exo_rows.iloc[0]  # Take the first one

            # Find the row which has 'ACTFLO~DEMO' in 'Attribute' column
            actflo_demo_rows = times_df_filtered[
                times_df_filtered["Attribute"] == "ACTFLO~DEMO"
            ]
            if actflo_demo_rows.empty:
                print(f"No 'ACTFLO~DEMO' in 'Attribute' for process {process_name}")
                continue

            actflo_demo_row = actflo_demo_rows.iloc[0]

            # Initialize a flag to check whether we need to create a new ACT_EFF row or update existing one
            create_new_act_eff = False

            # Find the first row which has 'INPUT' in 'Attribute' column
            input_rows = times_df_filtered[times_df_filtered["Attribute"] == "INPUT"]
            if not input_rows.empty:
                input_row = input_rows.iloc[0]
                create_new_act_eff = True  # We will create a new ACT_EFF row
            else:
                # If 'INPUT' not found, check for 'ACT_EFF' attribute
                act_eff_rows = times_df_filtered[
                    times_df_filtered["Attribute"] == "ACT_EFF"
                ]
                if act_eff_rows.empty:
                    print(
                        f"No 'INPUT' or 'ACT_EFF' in 'Attribute' for process {process_name}"
                    )
                    continue
                input_row = act_eff_rows.iloc[0]
                create_new_act_eff = False

            # Now compute the ACT_EFF values for each year
            act_eff_values = {}
            years_columns = [
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
            for year in years_columns:
                try:
                    input_value = float(input_row[year])
                    exo_value = float(exo_row[year])
                    actflo_demo_value = float(actflo_demo_row[year])
                    act_eff = exo_value / actflo_demo_value / input_value
                    act_eff_values[year] = act_eff
                except (ValueError, ZeroDivisionError, KeyError, TypeError):
                    # Handle any errors, set value to empty string
                    act_eff_values[year] = ""

            if create_new_act_eff:
                # Create a new ACT_EFF row and insert after end_index
                new_row = {
                    col: "" for col in times_df.columns
                }  # Initialize with empty strings
                new_row["TechName"] = process_name
                new_row["Attribute"] = "EFF"

                # Set the values for the years
                for year, value in act_eff_values.items():
                    new_row[year] = value

                # Convert new_row to DataFrame
                new_row_df = pd.DataFrame([new_row])

                # Insert new_row_df into times_df after end_index
                # Split times_df into before, new_row_df, after
                times_df = pd.concat(
                    [
                        times_df.iloc[: end_index + 1],
                        new_row_df,
                        times_df.iloc[end_index + 1 :],
                    ]
                ).reset_index(drop=True)
            else:
                # Update the existing ACT_EFF row with the new values
                act_eff_index = input_row.name  # The index of the existing ACT_EFF row
                for year, value in act_eff_values.items():
                    times_df.at[act_eff_index, year] = value

            # Clear yearly data in rows with 'INPUT' and 'OUTPUT' attributes
            for attr in ["INPUT", "OUTPUT"]:
                attr_rows = times_df_filtered[times_df_filtered["Attribute"] == attr]
                for idx in attr_rows.index:
                    times_df.loc[idx, years_columns] = ""

    times_df = times_df.fillna("")

    # Return the updated times_df
    return times_df


def extract_and_save_actflo_demo_data(
    times_df, output_file_path="output_data/Scen_tra_actflo.xlsx", sheet_name="INS"
):
    """
    Extracts rows from times_df where Attribute == 'ACTFLO~DEMO', transforms them, and writes them
    into another Excel file in the specified sheet in the desired format.

    The final format in the sheet is expected to have the following columns:
    TimeSlice | LimType | Attribute | Year | Other_Indexes | DE | Pset_PN | Pset_Set | Pset_PD | Pset_CI | Pset_CO | Cset_Set | Cset_CN | Cset_CD | Attrib_Cond | Val_Cond

    Example of desired format (based on the user-provided snippet):

    Trans - Insert
        ~TFM_INS
        TimeSlice   LimType Attribute Year Other_Indexes DE       Pset_PN                Pset_Set Pset_PD Pset_CI Pset_CO Cset_Set Cset_CN Cset_CD Attrib_Cond Val_Cond
                    ACTFLO  2021     DEMO 8,2          tra_road_bus_bev_pass_short_0
                    ACTFLO  2024     DEMO 9            tra_road_bus_bev_pass_short_0
                    ACTFLO  2027     DEMO              tra_road_bus_bev_pass_short_0

    In this example:
    - Attribute is always 'ACTFLO'
    - Other_Indexes is always 'DEMO'
    - Pset_PN is the process (TechName)
    - DE and Year values come from the times_df row where Attribute was 'ACTFLO~DEMO'
    - The rest columns can be left blank as per the provided snippet.
    """

    from openpyxl import load_workbook

    # Identify all processes that have at least one 'ACTFLO~DEMO' row
    processes_with_demo = times_df.loc[
        times_df["Attribute"] == "ACTFLO~DEMO", "TechName"
    ].unique()

    # If no processes found, just return
    if len(processes_with_demo) == 0:
        return

    # Prepare a DataFrame to hold transformed data
    transformed_columns = [
        "TimeSlice",
        "LimType",
        "Attribute",
        "Year",
        "Other_Indexes",
        "DE",
        "Pset_PN",
        "Pset_Set",
        "Pset_PD",
        "Pset_CI",
        "Pset_CO",
        "Cset_Set",
        "Cset_CN",
        "Cset_CD",
        "Attrib_Cond",
        "Val_Cond",
    ]
    transformed_df = pd.DataFrame(columns=transformed_columns)

    # Define the years to consider
    years = [
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

    # For each process, extract ACTFLO~DEMO rows and transform
    for process in processes_with_demo:
        subset = times_df[
            (times_df["TechName"] == process) & (times_df["Attribute"] == "ACTFLO~DEMO")
        ]
        # There could be multiple rows for the same process with ACTFLO~DEMO
        for _, row in subset.iterrows():
            # For each year that has a value
            for yr in years:
                if yr in row and row[yr] != "" and pd.notna(row[yr]):
                    # Create a new transformed row
                    new_row = {
                        "TimeSlice": "",  # Blank as per the snippet
                        "LimType": "",  # Blank as per the snippet
                        "Attribute": "ACTFLO",
                        "Year": yr,
                        "Other_Indexes": "DEMO",
                        "DE": row[yr],  # Yearly data from the ACTFLO~DEMO row
                        "Pset_PN": process,  # Process name
                        "Pset_Set": "",
                        "Pset_PD": "",
                        "Pset_CI": "",
                        "Pset_CO": "",
                        "Cset_Set": "",
                        "Cset_CN": "",
                        "Cset_CD": "",
                        "Attrib_Cond": "",
                        "Val_Cond": "",
                    }
                    transformed_df = pd.concat(
                        [transformed_df, pd.DataFrame([new_row])], ignore_index=True
                    )
                else:
                    # Even if empty, we might still want to add a row if the snippet suggests it?
                    # The snippet shows empty rows for 2027 for example.
                    # If we want to include empty rows as well:
                    # Uncomment the following block if needed:

                    # new_row = {
                    #     "TimeSlice": "",
                    #     "LimType": "",
                    #     "Attribute": "ACTFLO",
                    #     "Year": yr,
                    #     "Other_Indexes": "DEMO",
                    #     "DE": "",
                    #     "Pset_PN": process,
                    #     "Pset_Set": "", "Pset_PD": "", "Pset_CI": "", "Pset_CO": "",
                    #     "Cset_Set": "", "Cset_CN": "", "Cset_CD": "",
                    #     "Attrib_Cond": "", "Val_Cond": ""
                    # }
                    # transformed_df = pd.concat([transformed_df, pd.DataFrame([new_row])], ignore_index=True)
                    pass

    if transformed_df.empty:
        return

    # Load the output workbook and find the header in the INS sheet
    wb = load_workbook(output_file_path)
    if sheet_name not in wb.sheetnames:
        print(f"{sheet_name} sheet not found in {output_file_path}.")
        return
    ws = wb[sheet_name]

    # Find the header row containing 'Attribute'
    header_row = find_header_row(ws, "Attribute")

    # --- Clear existing data below the header row ---
    # Delete all rows after the header row, if any
    max_row = ws.max_row
    if max_row > header_row:
        ws.delete_rows(header_row + 1, max_row - header_row)

    # Get column headers and indices as before
    headers = {cell.value: cell.column for cell in ws[header_row]}
    transformed_columns = [
        "TimeSlice",
        "LimType",
        "Attribute",
        "Year",
        "Other_Indexes",
        "DE",
        "Pset_PN",
        "Pset_Set",
        "Pset_PD",
        "Pset_CI",
        "Pset_CO",
        "Cset_Set",
        "Cset_CN",
        "Cset_CD",
        "Attrib_Cond",
        "Val_Cond",
    ]

    col_indices = {}
    for col_name in transformed_columns:
        if col_name in headers:
            col_indices[col_name] = headers[col_name]
        else:
            last_col = ws.max_column
            ws.cell(row=header_row, column=last_col + 1, value=col_name)
            col_indices[col_name] = last_col + 1

    # Write the transformed data starting right after the header row
    start_row = header_row + 1
    for i, row_data in transformed_df.iterrows():
        for col_name in transformed_columns:
            ws.cell(
                row=start_row + i,
                column=col_indices[col_name],
                value=row_data[col_name],
            )

    wb.save(output_file_path)
    print(
        f"ACTFLO~DEMO data extracted and replaced in {output_file_path}, sheet {sheet_name}."
    )


# Paths and URLs
TIMES_FILE_PATH = "output_data/vt_DE_tra.xlsx"

# Read the pickle file and print the DataFrame
PICKLE_FILE_PATH = "output_data/times_df_tra.pkl"
times_df = pd.read_pickle(PICKLE_FILE_PATH)
# format_and_save_excel("test_output_cmp.xlsx", times_df)

# Create a copy of times_df to work with
updated_df = times_df.copy()

# Pre-defined process groups to handle
process_groups = [
    "tra_air_pass_0",
    "tra_air_pass_1",
    "tra_rail_pass_0",
    "tra_water_frei_0",
    "tra_water_frei_1",
    "tra_rail_frei_0",
    "tra_rail_frei_1",
    "tra_rail_pass_1",
]

# Load the desired units mapping once at the start
desired_units_mapping = load_desired_units_mapping()

# Define a global list to keep track of processes that have been handled
handled_processes = []

# Handle pre-defined process groups first
for process_group in process_groups:
    updated_df = data_mapping(updated_df, process_group, is_group=True)

# Fetch and process data for each unique process in the TechName column that starts with 'tra'
unique_processes = times_df["TechName"].unique()
tra_processes = [process for process in unique_processes if process.startswith("tra")]

# Skip processes that end with '_ag'
tra_processes = [process for process in tra_processes if not process.endswith("_ag")]

for process in tra_processes:
    if process not in handled_processes:
        updated_df = data_mapping(
            updated_df, process
        )  # Perform data mapping and update updated_df

# Calculate ACT_EFF attributes
updated_df = calculate_act_eff(updated_df, TIMES_FILE_PATH)
extract_and_save_actflo_demo_data(
    updated_df, output_file_path="output_data/Scen_tra_actflo.xlsx", sheet_name="INS"
)
format_and_save_excel(TIMES_FILE_PATH, updated_df)
update_commodity_list_units(TIMES_FILE_PATH, updated_units_mapping)
update_process_list_sheet(TIMES_FILE_PATH, updated_units_mapping)
print("Excel file saved")
