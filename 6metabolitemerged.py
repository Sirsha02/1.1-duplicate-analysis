#This is for merging all metabolites of a group e.g. control, obesity,t2d

import pandas as pd

# Define the Excel file path
excel_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\T2D.xlsx"

# Define the columns to output
columns_to_output = ['Unnamed: 0','Metabolite Name', 'Area_inrange','CV'] # Replace with your desired column names
 # Replace with your desired column names

# Read all sheets from the Excel file
data_frames = pd.read_excel(excel_file, sheet_name=None)

#data_frames =data_frames.groupby['Metabolite name'].reset_index()

# Merge all sheets into a single DataFrame
merged_df = pd.concat(data_frames.values(), ignore_index=True)

merged_df = merged_df.drop_duplicates()

# Output the specified columns
output_df = merged_df[columns_to_output]

# Write the output DataFrame to a new sheet in the same Excel file
with pd.ExcelWriter(excel_file, mode='a', engine='openpyxl') as writer:
    output_df.to_excel(writer, sheet_name='Merged', index=False)

print("Output successfully written to the Excel file.")

