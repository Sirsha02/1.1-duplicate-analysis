import pandas as pd

def check_numbers_between(input_file):
    df = pd.read_excel(input_file, sheet_name=None)

    for sheet_name, sheet_data in df.items():
        # Assuming the column with comma-separated numbers is named 'Area1'
        if 'Area' in sheet_data.columns:
            numbers_column = sheet_data['Area']
            
            # Create a new column for the matched numbers
            sheet_data['Area_inrange'] = ''
            
            for index in numbers_column.index:
                numbers = str(numbers_column.loc[index]).strip('[]')
                
                # Assuming the cells with the two numbers are named 'Number1' and 'Number2'
                number1 = sheet_data.loc[index, 'Mean'] #Mean is the Mean column
                number2 = sheet_data.loc[index, 'Stdev'] #Area is the standard deviation column
                highest_number = float('-inf')  # Initialize with negative infinity
                matched_numbers = []
                
                if ',' in numbers:
                    numbers = numbers.split(',')
                    
                    for number in numbers:
                        number = float(number.strip())  # Convert to float if necessary

                        if ((number >= number1) and (number <= (number1+number2)) or ((number <= number1) and (number >= (number1-number2)))):
                            if number > highest_number:
                                highest_number = number

                else:
                    matched_numbers.append(str(numbers))   

                if highest_number != float('-inf'):
                    matched_numbers.append(str(highest_number))

                if matched_numbers:
                    matched_numbers_str = ', '.join(matched_numbers)
                    sheet_data.at[index, 'Area_inrange'] = matched_numbers_str

    with pd.ExcelWriter(input_file) as writer:
        for sheet_name, sheet_data in df.items():
            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

# Usage example
input_file = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\IKPvsNISER\New folder\New folderIKP_1_.xlsx"
check_numbers_between(input_file)
