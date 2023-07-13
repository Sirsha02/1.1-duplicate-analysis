import pandas as pd
import glob
import itertools
import os
import numpy as np
import traceback
from tqdm import tqdm

def data_cleaning(dir1, dir2):  # input as a folder
    """
    This method runs a loop through two directories, selecting their files and then processing them through
    merged() and duplicate_analysis().

    Parameters :
    - dir1 : directory of the folder named 'POSITIVE' containing all the positive modes
    - dir2 : directory of the folder named 'NEGATIVE' containing all the negative modes

    Returns :
    - two excel files saved in a new folder named output and merged

    Author : Sunaina
    """

    file_1 = glob.glob(dir1 + "\*.xlsx")  # pos
    file_2 = glob.glob(dir2 + "\*.xlsx")  # neg

    filenames = []

    # filenames = filenames.append(file_1) # ["POS.xlsx", "NEG.xlsx"]
    for file1, file2 in zip(file_1, file_2):
        try:
            positive = file1
            negative = file2

            output_dictionary = {}

            df1 = pd.read_excel(
                positive, sheet_name=None
            )  # making a dictionary where each key is related to the dataframe of each sheet

            df2 = pd.read_excel(negative, sheet_name=None)

            merged_df = merge_pos_neg(df1, df2)  # dictionary

            merged_name = (
                r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sunaina"
                + os.path.splitext(os.path.basename(file1))[0]
                + "_merged.xlsx"
            )

            with pd.ExcelWriter(
                merged_name, engine="xlsxwriter"
            ) as writer:
                # you can change the file path here to whatever you want
                for sheet, df in tqdm(merged_df.items(), desc="Sheet Progress"):
                    df.to_excel(writer, sheet_name=sheet)

            sheets = list(merged_df.keys())

            for sheet in sheets:
                per_sheet = merged_df[sheet]  # dataframe
                output = duplicate_analysis(per_sheet)  # dataframe
                output_dictionary[sheet] = output  # dictionary

            new_file_name = (
                r"C:\Users\Sirsha\Desktop\LAB WORK NISER\Sunaina"
                + os.path.splitext(os.path.basename(file1))[0]
                + "_diff.xlsx"
            )

            with pd.ExcelWriter(
                new_file_name, engine="xlsxwriter"
            ) as writer:
                # you can change the file path here to whatever you want
                for sheet, merged_df in tqdm(output_dictionary.items(), desc="Output Progress"):
                    merged_df.to_excel(writer, sheet_name=sheet)

        except Exception as e:
            print("This file is giving an error ->" + str(file1) + str(file2))
            print(traceback.format_exc())
            continue


def merge_pos_neg(dic_1, dic_2):
    """
    Input as two dictionaries where each key is the sheet_name and value is the dataframe of that sheet's data.
    """
    keys_1 = list(dic_1.keys())  # ['SAM01', 'SAM02']
    keys_2 = list(dic_2.keys())  # ['SAM01', 'SAM02']

    merged_diction = {}

    for i, j in zip(keys_1, keys_2):
        if len(i) == 2 and len(j) == 2:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df
        if len(i) == 3 and len(j) == 3:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df
        if len(i) == 4 and len(j) == 4:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df
    
        if len(i) == 5 and len(j) == 5:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df

        if len(i) == 6 and len(j) == 6:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df

        if len(i) == 7 and len(j) == 7:
            if i[:-2] == j[:-2]:
                merged_df = pd.concat([dic_1[i], dic_2[j]], ignore_index=True)
                new_key = i
                merged_diction[new_key] = merged_df

    return merged_diction

def duplicate_analysis(per_sheet):
    df = per_sheet.dropna(how="all", axis=1)
    df = df.sort_values("MB NAME")
    
    df.loc[df['AREA'].apply(lambda x: isinstance(x, str)), 'AREA'] = np.nan
    #numeric_df = df[pd.to_numeric(df['AREA'], errors='coerce').notnull()]
    std_df = df.groupby("MB NAME")["AREA"].std().reset_index()
    
    df["z_score"] = df.groupby("MB NAME")["AREA"].transform(calculate_z_scores)
    df["z_score"] = df["z_score"].fillna(0)
    df = df[abs(df["z_score"]) < 1.5]

    mean_df = (
        df.groupby("MB NAME")
        .agg(
            AREA=("AREA", list),
            n=("MB NAME", "size"),
            z_score=("z_score", list),
        )
        .reset_index()
    )

    final_df = pd.concat([mean_df, std_df['AREA'].rename('Std_dev')], axis=1)

    return final_df

def calculate_z_scores(group):
    std_dev = group.std()
    if std_dev != 0:
        return (group - group.mean()) / group.std()
    else:
        return np.nan

if __name__ == "__main__":
    dir1 = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\POS"  # add the path file here to the folder that contains all the data files
    dir2 = r"C:\Users\Sirsha\Desktop\LAB WORK NISER\NEG"
    data_cleaning(dir1, dir2)
