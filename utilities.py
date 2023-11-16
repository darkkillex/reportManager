import pandas as pd
from openpyxl.styles import Alignment

import constants


def load_xlsx_file(file_excel):
    df = None
    try:
        df = pd.read_excel(file_excel)
        # Operazioni sul dataframe df qui
    except FileNotFoundError:
        print("File Excel not found.")
    except pd.errors.EmptyDataError:
        print("The File Excel is empty.")
    except pd.errors.ParserError:
        print("Error reading Excel file. Make sure the format is correct.")
    except Exception as e:
        print(f"An unknown error has occurred: {str(e)}")
    return df


def autosize_and_center_columns(sheet):
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width

        # Center-align the content in each cell
        for cell in column:
            cell.alignment = Alignment(horizontal='center', vertical='center')


def clean_df(df):
    df.columns = df.iloc[0] # Set the 2nd row as of columns label of df
    df = df.iloc[1:] # delete the 1st useless row of xlsx
    df = df.reset_index(drop=True)
    return df



