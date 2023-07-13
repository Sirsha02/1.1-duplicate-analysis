import pandas as pd
import numpy as np
import glob
import itertools
import os

def data_cleaning(dir):
    file_name = glob.glob(dir + "\*.xlsx")
    
    for file in file_name:
        print(file)
        output_dictionary = {}
        
        df = pd.read_excel(file, sheet_name=None)
        #merged_df = merge_pos_neg(df)
        merged_df = df
        merged_name = (
           r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sunaina"
            + os.path.splitext(os.path.basename(file))[0]
            + "_merged.xlsx"
        )

        #with pd.ExcelWriter(merged_name, engine="xlsxwriter") as writer:
         #   for sheet, df in merged_df.items():
          #      df.to_excel(writer, sheet_name=sheet)

        sheets = list(merged_df.keys())

        for sheet in sheets:
            per_sheet = merged_df[sheet]
            output = duplicate_analysis(per_sheet)
            output_dictionary[sheet] = output

        new_file_name = (
            r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sunaina"
              + os.path.splitext(os.path.basename(file))[0]
              + "_diff.xlsx")
        

        with pd.ExcelWriter(new_file_name, engine="xlsxwriter") as writer:
            for sheet, merged_df in output_dictionary.items():
                merged_df.to_excel(writer, sheet_name=sheet)

def merge_pos_neg(dictionary):
    keys = list(dictionary.keys())
    merged_diction = {}

    for a, b in itertools.combinations(keys, 2):
        '''
        if a[0:1] == b[0:1] and a[-1] == "P" and b[-1] == "N":
            merged_df = pd.concat([dictionary[a], dictionary[b]], ignore_index=True)
            new_key = a
            merged_diction[new_key] = merged_df
        if a[0:2] == b[0:2] and a[-1] == "P" and b[-1] == "N":
            merged_df = pd.concat([dictionary[a], dictionary[b]], ignore_index=True)
            new_key = a
            merged_diction[new_key] = merged_df
        if a[0:3] == b[0:3] and a[-1] == "P" and b[-1] == "N":
            merged_df = pd.concat([dictionary[a], dictionary[b]], ignore_index=True)
            new_key = a
            merged_diction[new_key] = merged_df
            '''
        if a[:-2] == b[:-2] and a[-1] == "P" and b[-1] == "N":
            merged_df = pd.concat([dictionary[a], dictionary[b]], ignore_index=True)
            new_key = a
            merged_diction[new_key] = merged_df
    return merged_diction

def duplicate_analysis(per_sheet):
    df = per_sheet.dropna(how="all", axis=1)
    df = df.sort_values("Name")
    
    df.loc[df['Area'].apply(lambda x: isinstance(x, str)), 'Area'] = np.nan
    #numeric_df = df[pd.to_numeric(df['Area'], errors='coerce').notnull()]
    std_df = df.groupby("Name")["Area"].std().reset_index()
    
    df["z_score"] = df.groupby("Name")["Area"].transform(calculate_z_scores)
    df["z_score"] = df["z_score"].fillna(0)
    df = df[abs(df["z_score"]) < 1.5]

    mean_df = (
        df.groupby("Name")
        .agg(
            Area=("Area", list),
            n=("Name", "size"),
            z_score=("z_score", list),
        )
        .reset_index()
    )

    final_df = pd.concat([mean_df, std_df['Area'].rename('Std_dev')], axis=1)

    return final_df

def calculate_z_scores(group):
    std_dev = group.std()
    if std_dev != 0:
        return (group - group.mean()) / group.std()
    else:
        return np.nan

if __name__ == "__main__":
    dir = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sunaina"
    data_cleaning(dir)
