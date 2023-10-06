# This is a sample Python script.

# Press Maiusc+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import constants


def load_xlsx_file(file_excel):
    #load a xlsx file in a DF
    df = pd.read_excel(file_excel)
    return df


def clean_df(df):
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df = df.reset_index(drop=True)
    return df


df_1 = load_xlsx_file(constants.IN_OUT)
df_final = clean_df(df_1)
print(df_final)


if __name__ == '__main__':
    print("Hello world!")

