#This is for getting nly specific sheets as output

import pandas as pd

input_files = [
    {
        'file_path': r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\pre_process\20062023_posneg_1.xlsx",
        'sheets': ['PA1_P','PAF1_P','SA3_P','SA4_P','SA5_P','SE4c_P','SE5c_P','SE6c_P','SE6d_P','SE6e_P','SE15a_P','SE15b_P','SE17a_P']
    },
    {
        'file_path':r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\pre_process\20062023_posneg_1.xlsx",
        'sheets': ['SE18a_P']
    },
    # Add more input files and their corresponding sheets as needed
]

output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\T2D.xlsx"

# Initialize an empty dictionary to store the selected sheets from all files
dfs = {}

# Loop through each input file
for input_file in input_files:
    file_path = input_file['file_path']
    sheets = input_file['sheets']

    # Read the selected sheets from the input file
    file_dfs = pd.read_excel(file_path, sheet_name=sheets)

    # Merge the selected sheets into the main dictionary
    dfs.update(file_dfs)

# Write the selected sheets to the output Excel file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
