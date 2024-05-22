import pandas as pd
import requests
from openpyxl import load_workbook


def read_and_modify_excel(file_path, sheet_name="Sheet"):
    # Load the Excel file and sheet
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]

    # Convert the sheet data to a DataFrame
    data = [row for row in ws.iter_rows(values_only=True)]
    df = pd.DataFrame(data)

    # Find the row containing '~FI_T' in the first 5 rows and modify the DataFrame
    for idx, row in df.head(5).iterrows():
        if "~FI_T" in row.values:
            header_row_idx = idx + 1
            df.columns = df.iloc[header_row_idx]

            # Find the index of the "TechName" column
            techname_col_idx = df.columns.get_loc("TechName")

            # Select columns from "TechName" to the right
            df = df.iloc[header_row_idx + 1 :, techname_col_idx:].reset_index(drop=True)
            return df

    raise ValueError("Row with '~FI_T' not found in the first 5 rows")


def fetch_data(url):
    response = requests.get(url)
    return pd.DataFrame(response.json())


# Paths and URLs
EXCEL_FILE_PATH = "test_output.xlsx"
API_URL = "https://openenergy-platform.org/api/v0/schema/model_draft/tables/ind_steel_blafu_0/rows"

# Process the Excel file and print the modified DataFrame
modified_times_df = read_and_modify_excel(EXCEL_FILE_PATH)
print(modified_times_df)

# Fetch and print data from the URL
fetched_data = fetch_data(API_URL)
print(fetched_data)
