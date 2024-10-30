# Data Adapter Times

## Overview

This project is designed to convert SEDOS data into TIMES format for further use in emissions and energy modeling. The workflow involves configuring settings, preparing input data, and generating output data by running a sequence of scripts. The output includes multiple Excel and pickle files that represent emissions factors and other relevant energy sector data in TIMES format.

## Folder Structure

The project folder is organized as follows:

```plaintext
/config_data
    ├── mapping_v3.xlsx
    ├── SysSettings.xlsx
/input_data
    ├── Modellstruktur.xlsx
/output_data
    ├── Scen_emission_all.xlsx
    ├── vt_DE_CO2_Emission.xlsx
    ├── vt_DE_hea.xlsx
    ├── vt_DE_ind.xlsx
    ├── vt_DE_tra.xlsx
    ├── vt_DE_x2x.xlsx
    ├── times_df_hea.pkl
    ├── times_df_ind.pkl
    ├── times_df_tra.pkl
    ├── times_df_x2x.pkl
/src
    ├── fill_emission_factors.py
    ├── get_data.py
    ├── process_input.py
.gitignore
.pylintrc
README.md
```

## Requirements

- **Python 3.11**
- **Required Libraries**: Ensure you have all necessary Python packages installed by running:

  ```bash
  pip install -r requirements.txt
  ```

## Configuration Files

The ``/config_data/`` folder contains the configuration files required for running the conversion process:

``mapping_v3.xlsx``: This file contains mappings used during the conversion of SEDOS data to TIMES format.
``SysSettings.xlsx``: This file includes system-wide settings and configurations necessary for processing.

## Input Data

The ``/input_data/`` folder should contain the following file:

``Modellstruktur.xlsx``: This is the input data file, which must be updated with relevant SEDOS data before running the scripts. The scripts will read from this file and process the data.

## Output Data

The ``/output_data/`` folder will contain the following output files after running the scripts:

``Scen_emission_all.xlsx``: Contains aggregated emission data for the scenarios.
``vt_DE_CO2_Emission.xlsx``: Contains CO2 emission data for Germany.
``vt_DE_hea.xlsx, vt_DE_ind.xlsx, vt_DE_tra.xlsx, vt_DE_x2x.xlsx``: These are sector-specific files for heating, industry, transportation, and cross-sector processes.
``times_df_hea.pkl, times_df_ind.pkl, times_df_tra.pkl, times_df_x2x.pkl``: Intermediate Pickle files that hold data frames generated during the processing.


## Usage
### Step 1: Run the Conversion Process

Run the scripts in the following order to process the data and generate output:

``process_input.py``:
This script processes the input file (Modellstruktur.xlsx) and prepares the data for conversion.

``get_data.py``:
This script extracts the data from OEP and structures the required data into a format compatible with TIMES.

``fill_emission_factors.py``:
This script fills the emission factors and generates the final output files in the /output_data/ folder.

### Step 2: Review the Output

After running the scripts, the output files will be generated in the /output_data/ folder. These files will be ready for further analysis or integration with the TIMES model.