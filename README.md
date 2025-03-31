# Data Adapter Times

## Overview

This project is designed to convert SEDOS data into TIMES format for further use in energy system model. The workflow involves configuring settings, preparing input data, and generating output data by running a sequence of scripts. The output includes multiple Excel and pickle files that represent emissions factors and other relevant energy sector data in TIMES format.

## Requirements

- **Python 3.11**
- **Required Libraries**: Ensure you have all necessary Python packages installed by running:

  ```bash
  pip install -r requirements.txt
  ```

## Folder Structure

The project folder is organized as follows:

```plaintext
/config_data
    ├── mapping_v4.xlsx
    ├── SysSettings.xlsx
    ├── units_mapping_{sector}.xlsx
/input_data
    ├── Modellstruktur.xlsx
    ├── Process_Groups.csv (only in Power sector)
/output_data
    ├── Scen_emission_all_{sector}.xlsx
    ├── vt_DE_Demand_{sector}.xlsx
    ├── vt_DE_{sector}.xlsx
    ├── times_df_{sector}.pkl
    ├── ef_units_mapping.pkl
/src
    ├── fill_emission_factors.py
    ├── fill_demand.py
    ├── get_data.py
    ├── process_input.py
    ├── unit_conversion.py
.gitignore
.pylintrc
README.md
```

## Data Flow Process

The data conversion process follows these steps:

1. **Data Extraction**: Input SEDOS data from Modellstruktur.xlsx is read and validated
2. **Data Transformation**: The process data is mapped according to configuration parameters such as Input, Output, Attribute etc in mapping_v4.xlsx
3. **Data Enhancement**: Process data is retrieved from OEP (Open Energy Platform) and pasted to the appropriate process 
4. **Emission Factor Extraction**: Emission factors are extracted and applied accordingly
5. **Demand Extraction**: Demand factors are extracted and applied
6. **Output Generation**: Final data is exported in vt_DE_{sector} in both Excel (.xlsx) and pickle (.pkl) formats


## Configuration Files

The ``/config_data/`` folder contains the configuration files required for running the conversion process:

``mapping_v4.xlsx``: This file contains mappings used during the conversion of SEDOS data to TIMES format.
``SysSettings.xlsx``: This file includes system-wide settings and configurations necessary for processing.
``units_mapping_{sector}.xlsx``: This file contains mappings used during the conversion of unit of SEDOS data parameters to TIMES unit format.

## Src folder overview

### 1. `process_input.py`
- **Purpose**:
  - Processes the input file `Modellstruktur.xlsx` to extract and structure process-related data.
  - Generates intermediate data for further processing.
  - Creates and updates Excel sheets for commodities and processes.

- **Key Functions**:
  - `process_data`: Processes the input DataFrame to extract relevant data, including technology names, attributes, and commodity groups.
  - `add_comm_sheet_to_workbook`: Adds a "Commodity List" sheet to the Excel workbook, defining commodity set memberships.
  - `add_process_sheet_to_workbook`: Adds a "Process List" sheet to the Excel workbook, defining process-related attributes.
  - `update_commodity_groups`: Updates the `commodity_group` sheet in the configuration file with new commodity groups.

- **Output**:
  - A filtered DataFrame saved as a pickle file (`times_df_{sector}.pkl`).
  - An Excel file (`vt_DE_{sector}.xlsx`) with structured process and commodity data.

- **Dependencies**:
  - Requires `Modellstruktur.xlsx` as input.
  - Outputs data for use in `get_data.py`.

---

### 2. `get_data.py`
- **Purpose**:
  - Fetches additional data from the Open Energy Platform (OEP) API.
  - Maps and structures the fetched data into a format compatible with TIMES.
  - Updates the Excel file with processed data.

- **Key Functions**:
  - `fetch_data`: Fetches data from the OEP API for a given process or group.
  - `data_mapping`: Maps API data to TIMES-compatible formats and updates the DataFrame.
  - `format_and_save_excel`: Formats and saves the processed data into an Excel file.
  - `update_commodity_list_units`: Updates the "Commodity List" sheet with units from the API data.
  - `update_process_list_sheet`: Updates the "Process List" sheet with additional attributes like `Vintage` and `PrimaryCG`.

- **Output**:
  - Updates the Excel file (`vt_DE_{sector}.xlsx`) with TIMES-compatible data.
  - Saves emission factor units mapping as a pickle file (`ef_units_mapping.pkl`).

- **Dependencies**:
  - Uses the output from `process_input.py` for initial data structure.
  - Outputs data for use in `fill_emission_factors.py` and `fill_demand.py`.

---

### 3. `fill_emission_factors.py`
- **Purpose**:
  - Extracts and fills emission factors for processes.
  - Converts units as needed and integrates emission data into the final output.

- **Key Functions**:
  - `fetch_data`: Fetches emission factor data from the OEP API.
  - `process_emission_factors`: Processes emission factor data and applies unit conversions.
  - `process_group_or_individual`: Handles emission factors for both groups and individual processes.

- **Output**:
  - Updates the Excel file (`Scen_emission_all_{sector}.xlsx`) with emission factor data.

- **Dependencies**:
  - Uses data from `get_data.py` and unit conversion logic from `unit_conversion.py`.

---

### 4. `fill_demand.py`
- **Purpose**:
  - Fetches and fills demand data for processes.
  - Filters and integrates demand data into the final output.

- **Key Functions**:
  - `fetch_data`: Retrieves demand data from the OEP API.
  - `find_header_row`: Identifies the header row in the Excel sheet.
  - `clear_existing_data`: Clears old data from the Excel sheet before inserting new data.

- **Output**:
  - Updates the Excel file (`vt_DE_Demand_{sector}.xlsx`) with demand data.

- **Dependencies**:
  - Uses data from `get_data.py` and unit conversion logic from `unit_conversion.py`.

---

### 5. `unit_conversion.py`
- **Purpose**:
  - Provides utility functions for unit conversions.
  - Defines standard and composed units for energy modeling.

- **Key Functions**:
  - `define_energy_model_units`: Defines units and their relationships (e.g., `kWh`, `MWh`, `EUR/kWh`).
  - `get_conversion_factor`: Calculates the conversion factor between two units.
  - `convert_unit`: Converts values between different units.

- **Output**:
  - No direct output; used as a utility module by other scripts.

- **Dependencies**:
  - Used by `get_data.py`, `fill_emission_factors.py`, and `fill_demand.py` for unit conversions.

---

## Input Data

The ``/input_data/`` folder should contain the following file:

``Modellstruktur.xlsx``: This is the input data file, which must be updated with relevant SEDOS data before running the scripts. The scripts will read from this file and process the data.

## Output Data

The ``/output_data/`` folder will contain the following output files after running the scripts:

``Scen_emission_all_{sector}.xlsx``: Contains emission data for the scenarios.
``vt_DE_Demand_{sector}.xlsx``: Contains demand data for the sector.
``vt_DE_{sector}.xlsx``: This is the main output file for the process data. These are sector-specific files for heating, industry, transportation, and cross-sector processes.
``times_df_{sector}.pkl``: Intermediate Pickle files that hold data frames generated during the processing.
``ef_units_mapping.pkl``: Intermediate Pickle files that hold units mapping with source and destination unit for the exo outputs required in emission factors calculation.


## Usage
### Step 1: Run the Conversion Process

Run the scripts in the following order to process the data and generate output:

``process_input.py``:
This script processes the input file (Modellstruktur.xlsx) and prepares the data for conversion.

``get_data.py``:
This script extracts the data from OEP and structures the required data into a format compatible with TIMES.

``fill_emission_factors.py``:
This script fills the emission factors and generates the final output files in the /output_data/ folder.

``fill_demand.py``:
This script fills the demand values and generates the final output files in the /output_data/ folder.

### Step 2: Review the Output

After running the scripts, the output files will be generated in the /output_data/ folder. These files will be ready for further analysis or integration with the TIMES model.

## Troubleshooting

Common issues and their solutions:

- **Missing dependencies**: Ensure all requirements are installed
- **File not found errors**: Verify that the input files are in the correct location
- **OEP connection issues**: Check your internet connection and OEP API access
- **Excel file access errors**: Ensure Excel files are not open in another application
