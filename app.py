import pandas as pd
from datetime import datetime
import logging
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import constants
import inOut


wb_adr, wb_not_adr = Workbook(), Workbook()


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

def write_on_xlsx_sheet_file(wb, sheets):
    for group_name, group_data in sheets.items():
        sheet = wb.create_sheet(title=group_name)
        # adjust the width of the columns
        for col in range(1, len(group_data.columns) + 1):
            sheet.column_dimensions[sheet.cell(1, col).column_letter].width = 40
        # writa the data
        for row in dataframe_to_rows(group_data, index=False, header=True):
            sheet.append(row)
        # Set to Bold the first line
        for cell in sheet['2']:
            cell.font = Font(bold=True)

def create_IN_OUT_report(wb, sheet, label):
    write_on_xlsx_sheet_file(wb, sheet)
    wb.remove(wb['Sheet'])
    wb.save(label)


# TODO get the common date to add to label of report Excel
def generate_report_date(df):
    # Estrai la prima colonna contenente le date come stringhe
    date_column = df.iloc[:, 0]

    # Inizializza una lista vuota per conservare le date valide
    date_objects = []

    # Converte le stringhe in oggetti datetime, gestendo gli errori
    for date_string in date_column:
        try:
            date_time_obj = datetime.strptime(date_string, "%d/%m/%Y")
            date_objects.append(date_time_obj)
        except ValueError:
            # Se la conversione va in errore, puoi decidere come gestire la data non valida
            # In questo esempio, semplicemente aggiungiamo None alla lista
            date_objects.append(None)

    # Conta le occorrenze di ciascuna data
    date_counts = date_column.value_counts()
    # Trova la data con il conteggio massimo (data comune)
    common_date = date_counts.idxmax()
    return common_date


def run_scripts():
    # let's go work on In_OUT DFs
    df_1 = load_xlsx_file(constants.IN_OUT)
    if df_1 is not None:
        print(df_1)
        print(generate_report_date(df_1))
        df_final = clean_df(df_1)
        df = inOut.create_data_struct(df_final)
        df = inOut.populate_sheets(df)
        create_IN_OUT_report(wb_adr, inOut.adr_sheet, "assets/IN_OUT_ADR.xlsx")
        create_IN_OUT_report(wb_not_adr, inOut.not_adr_sheet, "assets/IN_OUT_NON_ADR.xlsx")


if __name__ == '__main__':
    run_scripts()
    print("IN_OUT_ADR.xlsx and IN_OUT_NON_ADR.xlsx has been generated!")

