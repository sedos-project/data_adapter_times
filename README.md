# Data Adapter Times

## Overview

This branch is designed to convert TIMES optimized results output into a format that can be uploaded in OEP for SEDOS.  

## Folder Structure

The project folder is organized as follows:

```plaintext
/input_excels
    ├── Excels of input data with commodity and process information
    ├── 
/result_csv
    ├── final result outout as csv for each scnerios
/result_vd
    ├── .vd files holding results out of TIMES
    ├── 
    ├── 

/src
    ├── sedos_output_format.py
    ├── unit_mapping.py
    ├── 
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

## input_excels

The ``/input_excels/`` folder will contain all the input excels with commodity and processes names and units that are part of the results

## result_vd

The ``/result_vd/`` folder will contain the following .vd files

## result_csv

The ``/result_csv/`` folder contains .csv outputs



## Usage
### Step 1: Run the Conversion Process

Run the scripts in the following order to process the data and generate output:

``sedos_output_format.py``:
This script processes the .vd files and creates .csv for each scenario
