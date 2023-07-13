
import pandas as pd

# Define the column name containing the names to compare
#column_name = 'Metabolite name'  # Replace with your desired column name

# Read the three DataFrames from different sources
df1 = pd.read_excel(r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\bkegg\15days_comm_unique.xlsx", sheet_name='Sheet1')
#df2 = pd.read_excel(r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\groups\Obesity.xlsx", sheet_name='Merged (2)')
#df3 = pd.read_excel(r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\groups\T2D.xlsx", sheet_name='Merged (2)')
df_out =pd.read_excel(r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sohini\present\15days\bkegg\functions\functioons.xlsx", engine="openpyxl")
# Extract the names from each DataFrame
names1 = set(df1['Metabolite name_5'])
names2 = set(df1['Metabolite name_6'])
names3 = set(df1["Metabolite name_10"])

# Find unique names in df1 that do not exist in df2 and df3
unique_names_df1 = names1 - names2 - names3

# Find unique names in df2 that do not exist in df1 and df3
unique_names_df2 = names2 - names1 - names3

# Find unique names in df3 that do not exist in df1 and df2
unique_names_df3 = names3 - names1 - names2

df_out['U_5']=unique_names_df1
df_out['U_6']=unique_names_df2
df_out['U_10']=unique_names_df3

'''
print('Control:', len(names1))
print('Obesity:', len(names2))
print('T2D:', len(names3))
# Print the results
print("Unique names in Control:", len(unique_names_df1))
print("Unique names in Obesity:", len(unique_names_df2))
print("Unique names in T2D:", len(unique_names_df3))

int12 = names1.intersection(names2)
int23 = names2.intersection(names3)
int31 = names3.intersection(names1)
int123 = names1.intersection(names2,names3)

# Print the results
print("Common in Control and Obesity:", len(int12))
print("Common in Obesity and T2D:", len(int23))
print("Common in T2D and Control:", len(int31))
print("Common in all :", len(int123))

'''

print('5a:', len(names1))
print('6a:', len(names2))
print('10a:', len(names3))
# Print the results
print("Unique names in 5a:", len(unique_names_df1))
print("Unique names in 6a:", len(unique_names_df2))
print("Unique names in 10a:", len(unique_names_df3))

int12 = names1.intersection(names2)
int23 = names2.intersection(names3)
int31 = names3.intersection(names1)
int123 = names1.intersection(names2,names3)

# Print the results
print("Common in 5a and 6a:", len(int12))
print("Common in 6a and 10a:", len(int23))
print("Common in 10a and 5a:", len(int31))
print("Common in all :", len(int123))

