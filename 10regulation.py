import pandas as pd
from openpyxl import load_workbook

# Specify the input Excel file path
input_file =r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\time_intervals\common__all_anova.xlsx"

# Specify the output Excel file path
output_file =r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\time_intervals\common_reg.xlsx"

# Specify the column names you want to consider
column_names = ['A/B','B/C','A/C','A/D','A/E','B/D','B/E','C/D','C/E','D/E']  # Replace with your column names


# Load the input Excel file
excel_file = pd.ExcelFile(input_file)

# Create a writer object to save the output Excel file
writer = pd.ExcelWriter(output_file, engine='openpyxl')

# Iterate over each sheet in the input Excel file
for sheet_name in excel_file.sheet_names:
    # Check if the sheet name starts with 'regulation'
    if sheet_name.lower().startswith('regulation'):
        # Read the sheet into a pandas DataFrame
        df = excel_file.parse(sheet_name)
        
        # Iterate over each column name
        for column_name in column_names:
            # Check if the current column name is present in the DataFrame
            if column_name in df.columns:
                # Get the column index
                column_index = df.columns.get_loc(column_name)
                
                # Append a new column to the DataFrame
                new_column_name = column_name
                df[new_column_name] = ''
                
                # Iterate over each row in the column
                for index, value in df[column_name].items():
                    # Check if the value is less than 1
                    if value < 1:
                        # Write 'UP' in the new column of the same row
                        df.at[index, new_column_name] = 'UP'
                    # Check if the value is greater than 1
                    elif value > 1:
                        # Write 'DOWN' in the new column of the same row
                        df.at[index, new_column_name] = 'DOWN'
        
        # Save the modified DataFrame to the output Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the changes and close the writer
writer._save()
writer.close()

















