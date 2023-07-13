import pandas as pd

# Load the Excel file
file_path = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\15days_1.xlsx"

# Read the Excel file into a dictionary of DataFrames
dfs = pd.read_excel(file_path, sheet_name=None)

# Define the word to erase from sheet names
word_to_erase = '_P'  # Replace with the word you want to erase

# Create a separate list of sheet names
sheet_names = list(dfs.keys())

# Rename the sheets by erasing the specified word
for sheet_name in sheet_names:
    new_sheet_name = sheet_name.replace(word_to_erase, '')
    dfs[new_sheet_name] = dfs.pop(sheet_name)

# Save the modified Excel file
with pd.ExcelWriter(file_path) as writer:
    for sheet_name, df in dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)