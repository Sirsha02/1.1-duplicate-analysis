#This code takes different files and specific columns from them and outputs all that in one file, the sheet names need to be exact same for this to work.



import pandas as pd

def merge_excel_files(file1, file2, output_file, columns1, columns2):
    # Read both Excel files
    df1 = pd.read_excel(file1, sheet_name=None)
    df2 = pd.read_excel(file2, sheet_name=None)

    merged_dfs = {}

    # Iterate over sheets in the first file
    for sheet_name, df_file1 in df1.items():
        # Check if the sheet exists in the second file
        if sheet_name in df2:
            # Create a new DataFrame to merge the data
            merged_df = pd.DataFrame()

            # Add the specified columns from the first file
            merged_df = pd.concat([merged_df, df_file1[columns1]], axis=1)

            # Get the specified columns from the second file
            df_file2 = df2[sheet_name]
            df_file2_selected = df_file2[columns2]

            # Add a blank column
            merged_df[''] = ''
            

            # Calculate the number of columns in the merged DataFrame
            num_columns = merged_df.shape[1]

            # Rename the columns from the second file and add them to the merged DataFrame
            df_file2_selected.columns = [f'{col} (File 2)' for col in df_file2_selected.columns]
            merged_df = pd.concat([merged_df, df_file2_selected], axis=1)

            # Append the merged DataFrame to the output dictionary
            merged_dfs[sheet_name] = merged_df

    # Write the merged DataFrames to the specified output file
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, merged_df in merged_dfs.items():
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)

# Provide the file paths for both input files, the output file, and the column names
file1 = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\30052023_posneg_diff.xlsx"
file2 = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\07062026 posneg_diff.xlsx"
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\merged.xlsx"
columns1 = ['Metabolite name', 'Area1']  # Specify the columns from file1 to include
columns2 = ['Metabolite name', 'Area1']  # Specify the columns from file2 to include

# Call the function to merge the files
merge_excel_files(file1, file2, output_file, columns1, columns2)
