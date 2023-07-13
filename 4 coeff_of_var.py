import pandas as pd
import numpy as np

def coeff_of_var(input_file):
    df = pd.read_excel(input_file, sheet_name=None)

    for sheet_name, sheet_data in df.items():
        # Assuming the column with comma-separated numbers is named 'Numbers'
        if 'Mean' in sheet_data.columns and 'Stdev' in sheet_data.columns:
            mean = sheet_data['Mean1'].fillna(1) #mean or mean1
            stdev = sheet_data['Stdev1'].fillna(1) #stdev or stdev1

            # Create a new column for the matched numbers
            sheet_data['CV'] = '' #output column
            #df['CV'] = np.where((df['Mean1'].isna()) | (df['Stdev1'].isna()), 1, df['CV'])
            for index in mean.index:
                if index in stdev.index:
                    numbers = stdev.loc[index] / mean.loc[index]

                    sheet_data.at[index, 'CV'] = numbers

    with pd.ExcelWriter(input_file) as writer:
        for sheet_name, sheet_data in df.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

# Usage example
input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\P14D_CV.xlsx"


coeff_of_var(input_file)



