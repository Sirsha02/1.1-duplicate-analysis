import pandas as pd

def check_numbers_between(input_file):
    df = pd.read_excel(input_file, sheet_name=None)

    for sheet_name, sheet_data in df.items():
        # Assuming the columns with numbers are named 'Mean' and 'Stdev'
        if 'Mean' in sheet_data.columns and 'Stdev' in sheet_data.columns:
            numbers_column = sheet_data[['Mean', 'Stdev']]
            
            # Create a new column for the matched numbers
            sheet_data['Area_inrange'] = ''

            for index, row in numbers_column.iterrows():
                number1 = row['Mean']
                number2 = row['Stdev']

                # Assuming the column with comma-separated numbers is named 'Area'
                numbers = str(sheet_data.at[index, 'AREA']).strip('[]') #excel heading of metabolite column & Area
                matched_numbers = []

                if ',' in numbers:
                    numbers = numbers.split(',')
                    numbers = [float(number.strip()) for number in numbers]  # Convert to float if necessary
                    
                    for number in numbers:
                        if ((number >= number1) and (number <= (number1 + number2))) or ((number <= number1) and (number >= (number1 - number2))):
                            matched_numbers.append(str(number))

                else:
                    matched_numbers.append(str(numbers))

                if matched_numbers:
                    matched_numbers_str = ', '.join(matched_numbers)
                    sheet_data.at[index, 'Area_inrange'] = matched_numbers_str

    with pd.ExcelWriter(input_file) as writer:
        for sheet_name, sheet_data in df.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

# Usage example
input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\P14D.xlsx"


check_numbers_between(input_file)
