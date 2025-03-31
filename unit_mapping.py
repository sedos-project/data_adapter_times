import pandas as pd
import os

def process_multiple_excels(input_folder, output_file):
    # Initialize empty lists to store combined results
    all_commodity_data = []
    all_process_data = []
    
    # Get all Excel files from the input folder
    input_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith('.xlsx')]
    
    for input_file in input_files:
        # Read the sheets
        commodity_df = pd.read_excel(input_file, sheet_name='Commodity List', skiprows=1)
        process_df = pd.read_excel(input_file, sheet_name='Process List', skiprows=1)
        
        # Extract relevant columns
        sedos_commodity_df = commodity_df[['CommName', 'Unit']].rename(columns={'CommName': 'commodity', 'Unit': 'unit'})
        sedos_process_df = process_df[['TechName', 'Tcap']].rename(columns={'TechName': 'process', 'Tcap': 'unit'})
        
        # Append to list
        all_commodity_data.append(sedos_commodity_df)
        all_process_data.append(sedos_process_df)
    
    # Concatenate all data
    final_commodity_df = pd.concat(all_commodity_data, ignore_index=True)
    final_process_df = pd.concat(all_process_data, ignore_index=True)
    
    # Write to new Excel file
    with pd.ExcelWriter(output_file) as writer:
        final_commodity_df.to_excel(writer, sheet_name='SEDOS_commodity', index=False)
        final_process_df.to_excel(writer, sheet_name='SEDOS_process', index=False)

# Example usage
input_folder = "input_excels"  # Replace with the actual folder path
output_excel_file = "result_unit_mapping.xlsx"  # Output file name
process_multiple_excels(input_folder, output_excel_file)
print(f"Processed data saved to {output_excel_file}")
