import pandas as pd
from scipy.stats import ttest_ind

# Input file path
input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\15days_bkegg.xlsx"
# Sheet name to process
sheet_name = 'bkegg_610'

# Define the columns for each category
category1_columns = ['Area_6a','Area_10a']
category2_columns = ['Area_6b','Area_10b']

# Read input file
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Perform t-test on each row
t_statistic_values = []
p_value_values = []

for _, row in df.iterrows():
    values_a = []
    values_b = []
    for col in category1_columns:
        value_a = row[col]
        if isinstance(value_a, str) and ',' in value_a:
            values_a.extend([float(x.strip()) for x in value_a.split(',')])
        else:
            values_a.append(float(value_a))

    for col in category2_columns:
        value_b = row[col]
        if isinstance(value_b, str) and ',' in value_b:
            values_b.extend([float(x.strip()) for x in value_b.split(',')])
        else:
            values_b.append(float(value_b))

    t_statistic, p_value = ttest_ind(values_a, values_b)
    t_statistic_values.append(t_statistic)
    p_value_values.append(p_value)

# Add t-statistic and p-value columns to the DataFrame
df['TStat'] = t_statistic_values
df['PVal'] = p_value_values

# Output file path
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\15days_bkegg.xlsx"

# Write DataFrame to the output file with a new sheet
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='ttest_b4_610', index=False)


