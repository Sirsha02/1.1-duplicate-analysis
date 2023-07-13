
import pandas as pd
import numpy as np

def calculate_standard_deviation(data):
    output = []
    for item in data:
        item = str(item)#.strip('[]')# Remove leading and trailing square brackets
        if ',' in item:
            numbers = [float(num.strip()) for num in item.split(',')]
            output.append(np.std(numbers, ddof=1))
        else:
            output.append(item)
    return output

def mean(data):
    output = []
    for item in data:
        item = str(item)#.strip('[]') # Remove leading and trailing square brackets
        if ',' in item:
            numbers = [float(num.strip()) for num in item.split(',')]
            output.append(np.mean(numbers))
        else:
            output.append(item)
    return output

input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\P14D.xlsx"
output_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\P14D_CV.xlsx"

# Load the input file
xls = pd.ExcelFile(input_file)

# Create a writer for the output file
writer = pd.ExcelWriter(output_file, engine ='xlsxwriter')

# Iterate over each sheet in the input file
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
   # Get the column of numbers from the DataFrame
    input_data = df['Area_inrange'].tolist()

    output = mean(input_data)
    df['Mean1'] = output
    
    # Calculate standard deviation and return single values as is
    output_data = calculate_standard_deviation(input_data)
    
    # Add the output column to the DataFrame
    df['Stdev1'] = output_data

    
    # Save the modified sheet to the output file with the same sheet name
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the output file
writer.close() 
