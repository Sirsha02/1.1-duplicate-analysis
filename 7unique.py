import pandas as pd

# Function to get unique metabolite names for each sheet
def get_unique_metabolite_names(file_path):
    # Read Excel file into a dictionary of DataFrames
    sheets = pd.read_excel(file_path, sheet_name=None)
    unique_names = {}

    # Iterate over each sheet
    for sheet_name, sheet_data in sheets.items():
        # Get the column 'Metabolite name' and find unique names
        metabolite_names = sheet_data['Metabolite Name'].unique()
        
        # Remove metabolite names that appear in other sheets
        for other_sheet_name, other_sheet_data in sheets.items():
            if other_sheet_name != sheet_name:
                other_metabolite_names = other_sheet_data['Metabolite Name'].unique()
                metabolite_names = [name for name in metabolite_names if name not in other_metabolite_names]
        
        # Store unique names for the current sheet
        unique_names[sheet_name] = metabolite_names

    return unique_names

# Main code
input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\bkegg\allsheetscommon\all.xlsx"
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\bkegg\allsheetscommon\uniques.xlsx"

# Get unique metabolite names for each sheet
unique_metabolite_names = get_unique_metabolite_names(input_file)

# Write the unique names to a new Excel file
with pd.ExcelWriter(output_file) as writer:
    for sheet_name, metabolite_names in unique_metabolite_names.items():
        df = pd.DataFrame({'Metabolite Name': metabolite_names})
        df.to_excel(writer, sheet_name=sheet_name, index=False)
