import pandas as pd
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


def clean_df(df):
    df.columns = df.iloc[0] # Set the 2nd row as of columns label of df
    df = df.iloc[1:] # delete the 1st useless row of xlsx
    df = df.reset_index(drop=True)
    return df



