import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from unit_conversion import convert_unit

# Initialize the counter
fetch_data_counter = 0
units_mapping = {}
updated_units_mapping = {}
desired_units_mapping = {}


def load_desired_units_mapping(file_path="config_data/units_mapping_x2x.xlsx"):
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
    if "CommName" in headers and "Unit" in headers and "Ctype" in headers:
        commname_col = headers["CommName"]
        unit_col = headers["Unit"]
        ctype_col = headers["Ctype"]
    else:
        print("CommName or Unit or Ctype column not found in Commodity List sheet.")
        return

    # Loop over each resource in units_mapping
    for resource_name, fields in units_mapping.items():
        for field in fields:
            field_name = field["field_name"]
            field_unit = field["unit"]
            # Only process field names starting with 'conversion_factor_'
            if field_name.startswith("conversion_factor_"):
                # Remove the 'conversion_factor_' prefix directly
                comm_name = field_name[len("conversion_factor_") :]
                found = False
                for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
                    commname_cell = row[
                        commname_col - 1
                    ]  # openpyxl columns are 1-based
                    if commname_cell.value and isinstance(commname_cell.value, str):
                        commname = commname_cell.value.strip().lower()
                        # print(commname)
                        # Check if "_elec_" or other boundary conditions exist
                        if (
                            "_elec_" in commname
                            or commname.startswith("elec_")
                            or commname.endswith("_elec")
                            or commname == "elec"
                        ):
                            ctype_cell = row[ctype_col - 1]
                            ctype_cell.value = "ELC"
                        if (
                            "_heat_" in commname
                            or commname.startswith("heat_")
                            or commname.endswith("_heat")
                            or commname == "heat"
                        ):
                            ctype_cell = row[ctype_col - 1]
                            ctype_cell.value = "HTHEAT"
                        if (
                            commname_cell.value.strip().lower()
                            == comm_name.strip().lower()
                        ):
                            unit_cell = row[unit_col - 1]
                            unit_cell.value = field_unit
                            found = True
                            break  # Assuming CommName is unique
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
    tcap_col = headers.get("Tcap")

    if not (techname_col and vintage_col and primary_cg_col and tact_col and tcap_col):
        print("Some required columns are missing in 'Process List' sheet.")
        return

    # Set Vintage column to 'NO' and populate PrimaryCG, Tact, and TCap
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
                if process_name in ["x2x_g2p_pemfc_ls_1", "x2x_g2p_sofc_ls_1"]:
                    primary_cg = "NRGO"
                elif (
                    "_oref_" in process_name
                    or "_x2liquid_ft_" in process_name
                    or "_biogas_treatment" in process_name
                    or "_aec_" in process_name
                    or "_pemec_" in process_name
                    or "_soec_" in process_name
                    or "_coel_" in process_name
                ):
                    primary_cg = "NRGI"
                else:
                    primary_cg = output_commodity[0]

                row[primary_cg_col - 1].value = primary_cg
                primary_cg_temp = output_commodity[0]

                # Find the Tact unit by searching the resource in units_mapping
                for resource_name, fields in units_mapping.items():
                    for field in fields:
                        if (
                            field["field_name"] == f"conversion_factor_{primary_cg}"
                            or field["field_name"]
                            == f"conversion_factor_{primary_cg_temp}"
                        ):
                            if field["unit"]:
                                row[tact_col - 1].value = field["unit"]
            else:
                if process_name in ["x2x_g2p_pemfc_ls_1", "x2x_g2p_sofc_ls_1"]:
                    primary_cg = "NRGO"
                elif (
                    "_oref_" in process_name
                    or "_x2liquid_ft_" in process_name
                    or "_biogas_treatment" in process_name
                    or "_aec_" in process_name
                    or "_pemec_" in process_name
                    or "_soec_" in process_name
                    or "_coel_" in process_name
                ):
                    primary_cg = "NRGI"
                else:
                    primary_cg = "notFound"
                row[primary_cg_col - 1].value = primary_cg

                # Set Tact to the unit from Commodity List
                row[tact_col - 1].value = "notFound"

    # Save the workbook
    wb.save(excel_file_path)
    print("Process List sheet updated with Vintage, PrimaryCG, and Tact columns.")


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


def extract_single_value(value):
    """
    Extracts the single integer or float from a list if applicable.
    If the list is empty, returns None. Otherwise, returns the value as is.
    """
    if isinstance(value, list):
        if len(value) == 1:
            return value[0]
        elif len(value) == 0:
            return None
    return value


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

    if not api_process_data.empty:
        # Check if 'version' column exists before filtering
        if "version" in api_process_data.columns:
            api_process_data = api_process_data[
                api_process_data["version"] == "srd_point_draft"
            ]
        else:
            print(f"'version' column not found in the {process_name} process data.")

    if api_process_data.empty:
        return times_df  # Return the original DataFrame if no data is fetched

    # Fetch metadata
    metadata = fetch_process_metadata(process_name)

    if is_group:
        grp_name = process_name
        # Divide the data based on the 'type' column
        process_groups = api_process_data.groupby("type")

        process_count = 0  # Initialize a counter for the processes handled

        for process, group_data in process_groups:
            if process.endswith("_ag"):  # Skip processes ending with _ag
                continue

            # Remove columns where all values are NaN (i.e., columns without any data)
            group_data = group_data.dropna(axis=1, how="all")

            # Only process if not already handled
            if process not in handled_processes:
                handled_processes.add(process)
                times_df = data_mapping_internal(
                    times_df, process, group_data, metadata, grp_name
                )  # Call internal function for each process
                process_count += 1  # Increment the counter for each handled process

        print(
            f"{process_count} processes were handled inside the process group: {process_name}"
        )
        return times_df
    else:
        return data_mapping_internal(
            times_df, process_name, api_process_data, metadata, grp_name="default"
        )


def data_mapping_internal(times_df, process_name, api_process_data, metadata, grp_name):

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

    pasted_combinations = set()

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
                    api_value = extract_single_value(api_value)
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
                            # Fetch unit from units_mapping by matching resource_name to process_name
                            source_unit = None
                            for resource_name, fields in units_mapping.items():
                                if (
                                    resource_name == process_name
                                    or resource_name == grp_name
                                ):
                                    for field in fields:
                                        if field["field_name"] == api_col:
                                            source_unit = field["unit"]
                                            break
                                    if source_unit:
                                        break
                            if not matching_row.empty:
                                # Fetch unit from units_mapping by matching resource_name to process_name
                                # If source unit is found, fetch the desired unit
                                if source_unit and source_unit not in {
                                    "kWh",
                                    "MWh",
                                    "GWh",
                                    "PJ",
                                    "MWh/MWh",
                                }:
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
                                    if source_unit not in {
                                        "kWh",
                                        "MWh",
                                        "GWh",
                                        "PJ",
                                        "MWh/MWh",
                                    }:
                                        print(
                                            f"Source unit {source_unit} for {api_col} not found."
                                        )

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
                                    # If source unit is found, fetch the desired unit
                                    if source_unit and source_unit not in {
                                        "kWh",
                                        "MWh",
                                        "GWh",
                                        "PJ",
                                        "MWh/MWh",
                                    }:
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
                                        if source_unit not in {
                                            "kWh",
                                            "MWh",
                                            "GWh",
                                            "PJ",
                                            "MWh/MWh",
                                        }:
                                            print(
                                                f"Source unit {source_unit} for {api_col} not found."
                                            )
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
                                current_combo = (flow_share_commodity, year)

                                for idx in matching_row.index:
                                    comm_in = times_df_filtered.at[idx, "Comm-IN"]
                                    comm_out = times_df_filtered.at[idx, "Comm-OUT"]
                                    comm_grp = times_df_filtered.at[idx, "CommGrp"]
                                    if flow_share_commodity in (comm_in, comm_out):
                                        if current_combo not in pasted_combinations:
                                            pasted_combinations.add(current_combo)
                                            if api_value is not None:
                                                times_df_filtered.at[idx, str(year)] = (
                                                    api_value / 100
                                                )
                                                times_df_filtered.at[idx, "LimType"] = (
                                                    constraint
                                                )
                                        else:
                                            matching_row_new = times_df_filtered[
                                                (
                                                    times_df_filtered["LimType"]
                                                    == constraint
                                                )
                                                & (
                                                    times_df_filtered["Attribute"]
                                                    == times_col
                                                )
                                                & (
                                                    (
                                                        times_df_filtered["Comm-IN"]
                                                        == flow_share_commodity
                                                    )
                                                    | (
                                                        times_df_filtered["Comm-OUT"]
                                                        == flow_share_commodity
                                                    )
                                                )
                                            ]
                                            if matching_row_new.empty:

                                                # Add a new row if the Attribute does not exist
                                                new_row = pd.Series(
                                                    {
                                                        col: pd.NA
                                                        for col in times_df_filtered.columns
                                                    }
                                                )
                                                new_row["TechName"] = process_name
                                                new_row["Comm-IN"] = comm_in
                                                new_row["Comm-OUT"] = comm_out
                                                new_row["CommGrp"] = comm_grp
                                                new_row["Attribute"] = times_col
                                                new_row["LimType"] = constraint
                                                if api_value is not None:
                                                    new_row[str(year)] = api_value / 100
                                                times_df_filtered = pd.concat(
                                                    [
                                                        times_df_filtered,
                                                        new_row.to_frame().T,
                                                    ],
                                                    ignore_index=True,
                                                )
                                            else:
                                                for idx in matching_row_new.index:
                                                    times_df_filtered.at[
                                                        idx, "CommGrp"
                                                    ] = comm_grp
                                                    if api_value is not None:
                                                        times_df_filtered.at[
                                                            idx, str(year)
                                                        ] = (api_value / 100)

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
                            # Fetch unit from units_mapping by matching resource_name to process_name
                            source_unit = None
                            for (
                                resource_name,
                                fields,
                            ) in units_mapping.items():
                                if (
                                    resource_name == process_name
                                    or resource_name == grp_name
                                ):
                                    for field in fields:
                                        if field["field_name"] == api_col:
                                            source_unit = field["unit"]
                                            break
                                    if source_unit:
                                        break
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
                                            print(
                                                f"Source unit {source_unit} for {api_col} not found."
                                            )
                                            converted_value = api_value  # If no unit is found, use the value as is
                                        times_df_filtered.at[new_row_idx, str(year)] = (
                                            1 / converted_value
                                        )
                                        times_df_filtered.at[new_row_idx, "LimType"] = (
                                            constraint
                                        )
                                    else:
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
                                            print(
                                                f"Source unit {source_unit} for {api_col} not found."
                                            )
                                            converted_value = api_value  # If no unit is found, use the value as is
                                        times_df_filtered.at[new_row_idx, str(year)] = (
                                            converted_value
                                        )
                                        times_df_filtered.at[new_row_idx, "LimType"] = (
                                            constraint
                                        )
                            else:
                                for idx in matching_row.index:
                                    if api_value is not None:
                                        if "cb_coefficient" in sedos_item:
                                            # times_df_filtered.at[new_row_idx, str(year)] = (
                                            #     1 / api_value
                                            # )
                                            # If source unit is found, fetch the desired unit
                                            if source_unit:
                                                desired_unit = (
                                                    desired_units_mapping.get(
                                                        source_unit, None
                                                    )
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
                                                        ) in (
                                                            updated_units_mapping.values()
                                                        ):
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
                                                print(
                                                    f"Source unit {source_unit} for {api_col} not found."
                                                )
                                                converted_value = api_value  # If no unit is found, use the value as is
                                            times_df_filtered.at[idx, str(year)] = (
                                                1 / converted_value
                                            )
                                            times_df_filtered.at[idx, "LimType"] = (
                                                constraint
                                            )
                                        else:
                                            # If source unit is found, fetch the desired unit
                                            if source_unit:
                                                desired_unit = (
                                                    desired_units_mapping.get(
                                                        source_unit, None
                                                    )
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
                                                        ) in (
                                                            updated_units_mapping.values()
                                                        ):
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
                                                print(
                                                    f"Source unit {source_unit} for {api_col} not found."
                                                )
                                                converted_value = api_value  # If no unit is found, use the value as is
                                            times_df_filtered.at[idx, str(year)] = (
                                                converted_value
                                            )
                                            times_df_filtered.at[idx, "LimType"] = (
                                                constraint
                                            )

    # Implement CAP2ACT logic
    cap2act_value = 1  # Default to empty if no match is found

    if "storage" in process_name.lower():
        cap2act_value = (
            0.0036  # Set CAP2ACT to 0.0036 if process name contains "battery"
        )
    elif process_name.endswith("_1"):
        # Check if 'cost_inv_p' exists in the API process data columns
        if "cost_inv_p" in api_process_data.columns:
            cap2act_value = 31.536
    elif process_name.endswith("_0"):
        # Check if 'capacity_p_inst_0' exists in the API process data columns
        if "capacity_p_inst_0" in api_process_data.columns:
            cap2act_value = 31.536

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


def calculate_act_eff(times_df):
    """
    Calculates the ACT_EFF attribute for processes.

    For processes that have a FLO_SHAR attribute, the new ACT_EFF values are calculated
    as the first ACT_EFF row's value divided by the first INPUT row's value for each year.
    Then, for two specific processes ("x2x_g2p_pemfc_ls_1" and "x2x_g2p_sofc_ls_1"),
    the values are calculated as the first OUTPUT row (with Comm-OUT equal to "sec_elec")
    divided by the first INPUT row's value.

    In each case, a new row is inserted and the year column values of the INPUT and the used
    rows (ACT_EFF or OUTPUT) are cleared.

    Parameters:
    times_df (pandas.DataFrame): The DataFrame containing the TIMES data.

    Returns:
    pandas.DataFrame: The updated DataFrame with the new ACT_EFF (EFF) rows calculated.
    """

    # Create a copy of the DataFrame to work with
    updated_times_df = times_df.copy()

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

    # ----------------------------------------------------------
    # Process all processes that have a FLO_SHAR attribute.
    # Use FLO_SHAR as a flag to determine which processes to handle.
    flo_shar_process_positions = []
    for process in times_df["TechName"].unique():
        process_df = times_df[times_df["TechName"] == process]
        if not process_df[process_df["Attribute"] == "FLO_SHAR"].empty:
            start_index = process_df.index[0]
            end_index = process_df.index[-1]
            flo_shar_process_positions.append((start_index, end_index, process))
    # Sorting in descending order (if index shifting were a concern)
    flo_shar_process_positions.sort(reverse=True)

    for start_index, end_index, process in flo_shar_process_positions:
        process_subset = updated_times_df.iloc[start_index : end_index + 1]
        input_rows = process_subset[process_subset["Attribute"] == "INPUT"]
        act_eff_rows = process_subset[process_subset["Attribute"] == "ACT_EFF"]

        if input_rows.empty or act_eff_rows.empty:
            print(
                f"Missing INPUT or ACT_EFF for process {process}, skipping FLO_SHAR-based calculation."
            )
            continue

        # Use the first available rows
        input_row = input_rows.iloc[0]
        act_eff_row = act_eff_rows.iloc[0]
        act_eff_index = act_eff_rows.index[0]

        # Calculate new ACT_EFF values (ACT_EFF cell divided by INPUT cell for each year)
        for year in years_columns:
            try:
                input_value = float(input_row[year])
                act_eff_value = float(act_eff_row[year])
                new_val = act_eff_value / input_value if input_value else ""
            except (ValueError, ZeroDivisionError, KeyError, TypeError):
                new_val = ""
            updated_times_df.loc[act_eff_index, year] = new_val

        # Clear the year column values for the INPUT and OUTPUT rows.
        for index in input_rows.index:
            for year in years_columns:
                updated_times_df.loc[index, year] = ""

    # --------------------------------------------------------------------------
    # Now, process the two specific processes using the original logic.
    # Here, we calculate ACT_EFF as the first OUTPUT row's value (with Comm-OUT "sec_elec")
    # divided by the first INPUT row's value.
    specific_processes = ["x2x_g2p_pemfc_ls_1", "x2x_g2p_sofc_ls_1"]
    specific_process_positions = []
    for process_name in specific_processes:
        process_df = times_df[times_df["TechName"] == process_name]
        if process_df.empty:
            print(f"{process_name} was not found in times_df and hence was skipped")
            continue
        start_index = process_df.index[0]
        end_index = process_df.index[-1]
        specific_process_positions.append((start_index, end_index, process_name))
    specific_process_positions.sort(reverse=True)

    for start_index, end_index, process_name in specific_process_positions:
        process_subset = updated_times_df.iloc[start_index : end_index + 1]

        # Filter to get INPUT rows and the first OUTPUT row with Comm-OUT equal to "sec_elec".
        input_rows = process_subset[process_subset["Attribute"] == "INPUT"]
        output_rows = process_subset[
            (process_subset["Attribute"] == "OUTPUT")
            & (process_subset["Comm-OUT"] == "sec_elec")
        ]

        if input_rows.empty or output_rows.empty:
            print(
                f"Missing INPUT or OUTPUT for process {process_name}, skipping specific process calculation."
            )
            continue

        input_row = input_rows.iloc[0]
        output_row = output_rows.iloc[0]

        # Calculate new ACT_EFF values as (output value)/(input value) for each year.
        act_eff_values = {}
        for year in years_columns:
            try:
                input_value = float(input_row[year])
                output_value = float(output_row[year])
                act_eff = output_value / input_value if input_value else ""
                act_eff_values[year] = act_eff
            except (ValueError, ZeroDivisionError, KeyError, TypeError):
                act_eff_values[year] = ""

        # Create a new row (with Attribute "EFF") for the calculated values.
        new_row = {col: "" for col in times_df.columns}
        new_row["TechName"] = process_name
        new_row["Attribute"] = "EFF"
        for year, value in act_eff_values.items():
            new_row[year] = value

        # Insert the new row immediately after the process's rows.
        updated_times_df = pd.concat(
            [
                updated_times_df.iloc[: end_index + 1],
                pd.DataFrame([new_row]),
                updated_times_df.iloc[end_index + 1 :],
            ],
            ignore_index=True,
        )

        # Clear the year column values for the INPUT and OUTPUT rows.
        for index in input_rows.index:
            for year in years_columns:
                updated_times_df.loc[index, year] = ""
        for index in output_rows.index:
            for year in years_columns:
                updated_times_df.loc[index, year] = ""

    return updated_times_df


# Paths and URLs
TIMES_FILE_PATH = "output_data/vt_DE_x2x.xlsx"

# Read the pickle file and print the DataFrame
PICKLE_FILE_PATH = "output_data/times_df_x2x.pkl"
times_df = pd.read_pickle(PICKLE_FILE_PATH)

# Create a copy of times_df to work with
updated_df = times_df.copy()

# Pre-defined process groups to handle
process_groups = [
    # Add other process groups here if needed
]
# Load the desired units mapping once at the start
desired_units_mapping = load_desired_units_mapping()

# Define a global list to keep track of processes that have been handled
handled_processes = set()

# Handle pre-defined process groups first
for process_group in process_groups:
    updated_df = data_mapping(updated_df, process_group, is_group=True)

# Fetch and process data for each unique process in the TechName column that starts with 'x2x'
unique_processes = times_df["TechName"].unique()
x2x_processes = [process for process in unique_processes if process.startswith("x2x")]

# Skip processes that end with '_ag'
x2x_processes = [process for process in x2x_processes if not process.endswith("_ag")]

for process in x2x_processes:
    if process in handled_processes:
        print(f"Process {process} already handled in process group, skipping.")
        continue
    updated_df = data_mapping(
        updated_df, process
    )  # Perform data mapping and update updated_df

# Apply ACT_EFF calculation
updated_df = calculate_act_eff(updated_df)
# Save the updated DataFrame
format_and_save_excel(TIMES_FILE_PATH, updated_df)
update_commodity_list_units(TIMES_FILE_PATH, updated_units_mapping)
update_process_list_sheet(TIMES_FILE_PATH, updated_units_mapping)
print("Excel file saved")
