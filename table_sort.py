#This is sort function of table


import pandas as pd

input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\time_intervals\times.xlsx"
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\time_intervals\times_sorted_c.xlsx"
sort_column = "Area_b"  # Replace with the actual column name to sort by

# Read the Excel file into a dictionary of DataFrames
dfs = pd.read_excel(input_file, sheet_name=None)

# Sort the data in each sheet based on the specified column
for sheet_name, df in dfs.items():
    if sort_column in df.columns:
        df_sorted = df.sort_values(by=sort_column, ascending=True)
        dfs[sheet_name] = df_sorted
    else:
        print(f"Column '{sort_column}' not found in sheet '{sheet_name}'.")

# Write the sorted DataFrames to a new Excel file with the same sheet names
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in dfs.items():
        df.to_excel (writer, sheet_name=sheet_name, index=False)

