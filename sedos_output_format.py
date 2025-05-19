import os
import numpy as np
import pandas as pd
import re
import json


# input file location, location that contains .vd files
input_folder_path = "data_adapter_times/results_vd" 
output_folder_path = "data_adapter_times/results_csv"

result_files = ["tokio_w_ghg-cap_2304.vd", "riga_w_gas-ghg-cap_2304.vd", "sina_w_min_h2_ghg-cap_2304.vd",
                "sinb_w_min_wind_pv-elec_ghg-cap_2304.vd"] # list of result file names


# Prepare .vd file into dataframe

all_results_data = []  # empty list to store cleaned data


for file_name in result_files:
    if file_name.endswith('.vd'):
        input_file_path = os.path.join(input_folder_path, file_name)
        # extract scenario name from .vd file
        scenario_name, _ = os.path.splitext(os.path.basename(input_file_path))
        
        # open .vd input file to read data
        with open(input_file_path, 'r') as file:
            data = file.readlines()  # all lines are stored in data
            
        # process and clean data
        for line in data:
            if not line.strip() or line.startswith('*'):
                continue
            if scenario_name not in line:
                line = f'{scenario_name}, {line}'  # add scenario_name if not already present in line
            # replace "", split line by comma and add it to list
            all_results_data.append(line.replace('"', '').strip().split(','))  
            
# convert list to dataframe
cleaned_df = pd.DataFrame(all_results_data)


# column names for dataframe
col_names = ['Scenario', 'Attribute', 'Commodity', 'Process', 'Year', 'Region', 'Vintage',
             'TimeSlice', 'UserConstraint', 'Value']

cleaned_df.columns = col_names # assign column name in df
 
# remove white space in Atribute column 
cleaned_df['Attribute'] = cleaned_df['Attribute'].str.strip()

# change the 'Value' type from object to float
cleaned_df['Value'] = cleaned_df['Value'].astype(float)


# remove dummy processes that are not part of SEDOS

# delete IMPDEMZ, IMPNRGZ, IMPMATZ, -, becasue they are not relavent 
cleaned_df = cleaned_df[~cleaned_df['Process'].str.contains('IMPDEMZ|IMPNRGZ|IMPMATZ|-')]
#cleaned_df['Process'].unique()


# rewrite Scenario name
# SEDOS scenario name:  framework_model-structure_scenario

# chnage scenario name
scenario_name_dict = {
    "tokio_w_ghg-cap_2304": "t_all_tokio_v03",
    'riga_w_gas-ghg-cap_2304': 't_all_riga_v02',
    'sina_w_min_h2_ghg-cap_2304': 't_all_sina_v01',
    'sinb_w_min_wind_pv-elec_ghg-cap_2304': 't_all_sinb_v01'

}
cleaned_df['Scenario'] = cleaned_df['Scenario'].map(scenario_name_dict)
cleaned_df.head()


# columns to delete
# delete columns that are not necessary for SEDOS output format

# columns to drop
columns_2_dlt = ['Region', 'Vintage', 'TimeSlice', 'UserConstraint']
sedos_result_df = cleaned_df.drop(columns=columns_2_dlt)


# rename column header 
new_column_header = ['scenario', 'parameter', 'commodity', 'process', 'year', 'value']

sedos_result_df.columns = new_column_header
sedos_result_df.head()


#sedos_result_df['commodity'].unique()


# create new emission rows with emi_co2_eq 

emi_co2_eq_rows = []

for index, row in sedos_result_df.iterrows():
    if 'emi_ch4_f_' in row['commodity'] or 'emi_ch4_p_' in row['commodity']:
        emi_co2_eq_row = row.copy()
        emi_co2_eq_row['commodity'] = 'emi_co2_eq'
        # convert CH4 emission into CO2 equivalent
        emi_co2_eq_row['value'] = row['value'] * 28
        emi_co2_eq_rows.append(emi_co2_eq_row)
    elif 'emi_n2o_f_' in row['commodity'] or 'emi_n2o_p_' in row['commodity']:
        emi_co2_eq_row = row.copy()
        emi_co2_eq_row['commodity'] = 'emi_co2_eq'
        # convert N2O emission into CO2 equivalent
        emi_co2_eq_row['value'] = row['value'] * 265
        emi_co2_eq_rows.append(emi_co2_eq_row)
    elif 'emi_co2_f_' in row['commodity'] or 'emi_co2_p_' in row['commodity']:
        emi_co2_eq_row = row.copy()
        emi_co2_eq_row['commodity'] = 'emi_co2_eq'
        emi_co2_eq_rows.append(emi_co2_eq_row)
        
emi_co2_eq_df = pd.DataFrame(emi_co2_eq_rows)
emi_co2_eq_df.head()


# concate emi_co2_eq_df with sedos_result_df

sedos_result_df = pd.concat([sedos_result_df, emi_co2_eq_df], ignore_index=True)
#sedos_result_df.head()

#sedos_result_df['commodity'].unique()


# create specfic columns for SEDOS results output

sedos_result_df[['sector', 'category', 'specification', 'groups', 'new', 'exo',
                 'input_groups', 'output_groups', 'unit']] = ''
sedos_result_df.head()


# Existing and New technologies
# power sector: 01, 02 in the power sector always denote 0 "existing" processes but different ones. I think it is mainly used for different renewable potentials which have the same characteristics but different potentials.
# 11, 12 instead denote different 1 "new" processes.

for index, row in sedos_result_df.iterrows():
    # First, handle process names ending with _ag
    if row['process'].endswith('_ag'):
        # Look for the number before _ag
        ag_match = re.search(r'_(\d+)_ag$', row['process'])
        if ag_match:
            number = ag_match.group(1)
            # Use the same logic as before
            if len(number) == 1:
                if number == '0':
                    sedos_result_df.at[index, 'new'] = '0'
                else:
                    sedos_result_df.at[index, 'new'] = '1'
            else:
                # For multi-digit numbers
                if number.startswith('0'):
                    sedos_result_df.at[index, 'new'] = '0'
                else:
                    sedos_result_df.at[index, 'new'] = '1'
        else:
            # No number before _ag
            sedos_result_df.at[index, 'new'] = '0'
    
    # Then handle the original case for process names ending with a number
    else:
        match = re.search(r'_(\d+)$', row['process'])
        if match:
            # Get the matched number
            number = match.group(1)
            
            # Check if it's a single digit (like _0 or _1)
            if len(number) == 1:
                if number == '0':
                    sedos_result_df.at[index, 'new'] = '0'
                else:
                    sedos_result_df.at[index, 'new'] = '1'
            # Check if it's a multi-digit number
            else:
                # Check if it starts with '0' (like _01, _02)
                if number.startswith('0'):
                    sedos_result_df.at[index, 'new'] = '0'
                # For other multi-digit numbers (like _11, _15)
                else:
                    sedos_result_df.at[index, 'new'] = '1'
        else:
            # No number at the end
            sedos_result_df.at[index, 'new'] = '0'
        
        
    # value for input_groups and output_groups based on parameter VAR_Fin and VAR_FOut
    # input commodity
    if 'VAR_FIn' in row['parameter']:
        sedos_result_df.at[index, 'input_groups'] = json.dumps([row['commodity']])
        #sedos_result_df.at[index, 'output_groups'] = [] # check later
        # output commodity
    elif 'VAR_FOut' in row['parameter']:
        sedos_result_df.at[index, 'output_groups'] = json.dumps([row['commodity']])
        #sedos_result_df.at[index, 'input_groups'] = [] # check later
    #else:
        #sedos_result_df.at[index, 'output_groups'] = [] # check later
        #sedos_result_df.at[index, 'input_groups'] = [] # check later
       
sedos_result_df.head()

# Unit of Cost and Capacity 
# unit for cost (investment, FOM and Variable cost) should be just 'MEUR'
# unit for capacity (VAR_Cap, VAR_NCap) can be either 'Mt', 'GW', 'PJ' or 'GWh' (how can be these nfomation efficienty implemented?) (Maybe can be read from Capacity unit of a process from VEDA-input??)
# VAR_CAP- total capacity of that milestone year, VAR_NCAP- total new installed capacity of that milestone year

# mapping, required parameters

sedos_result_df['parameter'].unique()


# assign units to commodity
# can create it from unit from 'Commodity_items' of VEDA input infrastructure (curently similar mapping is used)

results_unit_path = "data_adapter_times/result_unit/result_unit_mapping.xlsx" 

unit_mapping_commodity = pd.read_excel(results_unit_path, engine='openpyxl', sheet_name='SEDOS_commodity')

# remove duplicate
unit_mapping_commodity = unit_mapping_commodity.drop_duplicates(subset='commodity') #unit_mapping_commodity

unit_mapping_process = pd.read_excel(results_unit_path, engine='openpyxl', sheet_name='SEDOS_process')
unit_mapping_process = unit_mapping_process.drop_duplicates(subset='process') #unit_mapping_process


unit_mapping_commodity_dict = unit_mapping_commodity.set_index('commodity').to_dict()['unit'] #unit_mapping_commodity_dict


unit_mapping_process_cap_dict = unit_mapping_process.set_index('process').to_dict()['unit'] #unit_mapping_process_cap_dict



# SEDOS unit for other parameters that are not related with VAR_FOut and VAR_FIn
sedos_cost_unit_dict = {
    'costs_investment': 'MEUR',
    'costs_fixed': 'MEUR',
    'costs_variable': 'MEUR',
}


# SEDOS output parameters
sedos_output_parameters_dict = {
    'VAR_FOut': 'flow_volume',
    'VAR_FIn': 'flow_volume',
    'Cost_Inv': 'costs_investment',
    'Cost_Fom': 'costs_fixed',
    'Cost_Act': 'costs_variable',
    'VAR_Cap': 'capacity_inst',
    'VAR_Ncap': 'capacity_new'
}


# keep sedos output data with required parameters
sedos_result_df = sedos_result_df[sedos_result_df['parameter'].isin(set(sedos_output_parameters_dict.keys()))]

# replace parameteres name into SEDOS
sedos_result_df.loc[:, 'parameter'] = sedos_result_df['parameter'].replace(sedos_output_parameters_dict)
sedos_result_df


# sedos output data with required parameters
# create temporary columns of unit_commodity and unit_parameter
# sedos_result_df['unit_commodity'] = sedos_result_df['commodity'].map(unit_mapping_commodity_dict) # mapped with commodities for VAR_FOut and VAR_FIn 
# sedos_result_df['unit_cost'] = sedos_result_df['parameter'].map(sedos_cost_unit_dict) # mapped with paprematers
# sedos_result_df['unit_cap'] = sedos_result_df['process'].map(unit_mapping_process_cap_dict) # mapped with paprematers


sedos_result_df['unit_commodity'] = np.where(sedos_result_df['parameter'].str.contains('flow_volume'),
                                             sedos_result_df['commodity'].map(unit_mapping_commodity_dict), 
                                             np.nan)

sedos_result_df['unit_cost'] = np.where(sedos_result_df['parameter'].str.contains('costs'),
                                             sedos_result_df['parameter'].map(sedos_cost_unit_dict), 
                                             np.nan)


sedos_result_df['unit_cap'] = np.where(sedos_result_df['parameter'].str.contains('capacity'),
                                             sedos_result_df['process'].map(unit_mapping_process_cap_dict), 
                                             np.nan)




sedos_result_df[sedos_result_df['unit_commodity']=='PJ'] 

sedos_result_df[sedos_result_df['unit_cap']=='GW'] 


# combine both into final 'unit' column 
#sedos_result_df['unit'] = sedos_result_df['unit_commodity'].fillna(sedos_result_df['unit_cost']).fillna(sedos_result_df
sedos_result_df['unit'] = np.where(sedos_result_df['unit_commodity'].notna(), sedos_result_df['unit_commodity'], 
                                   np.where(sedos_result_df['unit_cost'].notna(), sedos_result_df['unit_cost'],
                                           sedos_result_df['unit_cap'])
                                  )

# Finally, remove the temporary columns
sedos_result_df = sedos_result_df.drop(columns=['unit_commodity', 'unit_cost', 'unit_cap'])
print(sedos_result_df)

sedos_result_df['parameter'].unique()
# sector, catergory and specification columns
# create those column by spliting the process name
# value for sector, category and specification columns
sedos_result_df['sector'] = sedos_result_df['process'].apply(lambda x: x.split('_')[0])
sedos_result_df['category'] = sedos_result_df['process'].apply(lambda x: x.split('_')[1])
sedos_result_df['specification'] = sedos_result_df['process'].apply(lambda x: x.split('_')[2])

# drop commodity column, as values of commodity are in input and output groups columns
sedos_result_df = sedos_result_df.drop(columns='commodity')

unique_scenarios = sedos_result_df['scenario'].unique()

for scenario in unique_scenarios:
    # scenario specfic df
    scenario_df = sedos_result_df[sedos_result_df['scenario'] == scenario]
    # create id column
    scenario_df['id'] = range(1, len(scenario_df) + 1)
    # specific file name for each scenario
    # rearrange column header 
    re_column_header = ['id', 'scenario', 'process', 'parameter', 'sector', 'category', 'specification', 'new',   
                    'groups',  'input_groups', 'output_groups', 'year',  'unit', 'value']
    scenario_df = scenario_df[re_column_header]
    
    filename = f"{scenario}.csv"

    output_file_path = os.path.join(output_folder_path, filename) # output file path
    # save as csv 
    scenario_df.to_csv(output_file_path, index=False, sep=';', decimal='.')
