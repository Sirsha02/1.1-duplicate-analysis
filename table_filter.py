#This is basically filter of table


import pandas as pd

input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\akegg\allsheetscommon\common.xlsx"
column_name = "Area_8b"  # Replace with the actual column name to check values
#output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\15days_common.xlsx"

# Read the Excel file into a dictionary of DataFrames
dfs = pd.read_excel(input_file, sheet_name='Sheet2')

with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
# Process each sheet
    for sheet_name, df in dfs.items():
        if column_name in df.columns:
            # Filter the DataFrame to keep rows with values <= 1.5 in the specified column
            #####df_filtered = df[df[column_name] <= 1.5]
            #####df_filtered = df[df[column_name].isnull()]
            df_filtered = df[df[column_name].notnull()]
            
            df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

        else:
            print(f"Column '{column_name}' not found in sheet '{sheet_name}'.")
