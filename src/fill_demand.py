import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# API URLs for the pow_scalars data
API_URLS = [
    "https://openenergyplatform.org/api/v0/schema/model_draft/tables/pow_scalars_b/rows",
    "https://openenergyplatform.org/api/v0/schema/model_draft/tables/pow_scalars_e/rows",
    "https://openenergyplatform.org/api/v0/schema/model_draft/tables/pow_scalars_de/rows",
]

# Path to the Excel file where the demand data should be pasted
EXCEL_FILE_PATH = "output_data/vt_DE_Demand_pow.xlsx"


def fetch_data(url):
    """
    Fetch data from the given URL and return a DataFrame.
    """
    try:
        response = requests.get(url)
        if response.status_code == 200:
            print(f"Data fetched successfully from API: {url}")
            return pd.DataFrame(response.json())
        else:
            print(
                f"Failed to fetch data from {url}, status code: {response.status_code}"
            )
            return pd.DataFrame()
    except requests.RequestException as e:
        print(f"Error fetching data from {url}: {e}")
        return pd.DataFrame()


def find_header_row(sheet, header_name):
    """
    Finds the row number of the first occurrence of header_name in the sheet.
    """
    for row in range(1, 21):  # check first 20 rows
        for col in range(1, sheet.max_column + 1):
            cell_value = str(sheet.cell(row=row, column=col).value)
            if header_name.lower() in cell_value.lower():
                return row
    raise ValueError(f"Header row with '{header_name}' not found.")


def clear_existing_data(ws, header_row):
    """
    Clears all rows after header_row in the worksheet.
    """
    print(f"Clearing data after row {header_row+1}")
    for row in ws.iter_rows(min_row=header_row + 2, max_row=ws.max_row):
        for cell in row:
            cell.value = None


def get_column_indices(ws, header_row, required_headers):
    """
    Returns a dictionary mapping header names to their column indices
    in the row specified by header_row.
    """
    headers = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            for req in required_headers:
                if req.lower() in str(cell.value).lower():
                    headers[req] = cell.column
    if len(headers) != len(required_headers):
        raise ValueError("Not all required headers found in the sheet.")
    return headers


def main():
    # Collect all data from the three API URLs
    all_data = pd.DataFrame()
    for url in API_URLS:
        df = fetch_data(url)
        if not df.empty:
            all_data = pd.concat([all_data, df], ignore_index=True)
    if all_data.empty:
        print("No data fetched from any API. Exiting.")
        return

    # We want only the columns that start with "exo" and also the "year" column
    exo_columns = [col for col in all_data.columns if col.startswith("exo")]
    if "year" not in all_data.columns:
        print("Year column not found in API data. Exiting.")
        return

    # Filter data to keep only the exo columns plus the year column
    data_filtered = all_data[exo_columns + ["year"]]
    print(f"Filtered API data to columns: {exo_columns + ['year']}")

    # Load the Excel workbook and open the Demand sheet.
    wb = load_workbook(EXCEL_FILE_PATH)
    sheet_name = None
    for s in wb.sheetnames:
        if s.lower() == "demand":
            sheet_name = s
            break
    if sheet_name is None:
        print("Demand sheet not found; using active sheet.")
        ws = wb.active
    else:
        ws = wb[sheet_name]

    # Find header row by searching for "CommName"
    try:
        header_row = find_header_row(ws, "CommName")
    except ValueError as e:
        print(e)
        return

    # Get column indices for the required headers ("CommName", "Demand", "Year")
    required_headers = ["CommName", "Demand", "Year"]
    try:
        col_indices = get_column_indices(ws, header_row, required_headers)
    except ValueError as e:
        print(e)
        return

    # Clear existing data rows after header_row+1
    clear_existing_data(ws, header_row)
    current_row = header_row + 2

    # For each exo column, iterate over each API data row and paste a new row for each non-null record.
    for exo_col in exo_columns:
        for _, api_row in data_filtered.iterrows():
            year_value = api_row["year"]
            demand_value = api_row[exo_col]
            if pd.notna(demand_value):
                ws.cell(row=current_row, column=col_indices["CommName"], value=exo_col)
                ws.cell(
                    row=current_row, column=col_indices["Demand"], value=demand_value
                )
                ws.cell(row=current_row, column=col_indices["Year"], value=year_value)
                current_row += 1

    # Optionally adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2

    wb.save(EXCEL_FILE_PATH)
    print(f"Demand data processing completed and saved to {EXCEL_FILE_PATH}.")


if __name__ == "__main__":
    main()
