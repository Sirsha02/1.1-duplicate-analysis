import pandas as pd
import scipy.stats as stats

# Read the Excel file
excel_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\akegg\allsheetscommon\common - Copy.xlsx"
sheet_name = 'anova'  # Replace with the name of your specific sheet
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Define the column ranges for each group
group1_cols = df.columns[1:14]  # Adjust the column indices as per your data
group2_cols = df.columns[15:23]
group3_cols = df.columns[24:29]

# Create a new column for p-values
df['P-value'] = None

# Perform ANOVA row-wise and calculate p-values
for index, row in df.iterrows():
    group1_data = row[group1_cols].dropna()
    group2_data = row[group2_cols].dropna()
    group3_data = row[group3_cols].dropna()
    
    # Perform one-way ANOVA
    f_value, p_value = stats.f_oneway(group1_data, group2_data, group3_data)
    
    # Assign the p-value to the corresponding row in the new column
    df.at[index, 'P-value'] = p_value

# Save the updated DataFrame to a new Excel file
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\20062023\groups\akegg\allsheetscommon\anova.xlsx"
df.to_excel(output_file, index=False)
