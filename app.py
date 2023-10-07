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
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]

        # Elimina la prima riga (Header)
        sheet.delete_rows(1)

        # Imposta i filtri per tutte le colonne
        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = 'A2'
    wb.save(label)


def generate_IN_OUT_report_label(df, label):
    # Extract the first column containing the dates as strings
    if not df.empty:
        cell_value = df.iloc[1, 1] # get the date from
        date_parts = cell_value.split(' ')
        date_string = date_parts[0]
        # Replace the symbol "/" with "-"
        date_string = date_string.replace('/', '-') # now we have the date
        report_label = "assets/" + label + date_string + ".xlsx"

    return report_label


def run_scripts():
    # let's go work on In_OUT DFs
    df_1 = load_xlsx_file(constants.IN_OUT)
    if df_1 is not None:
        df_final = clean_df(df_1)
        df = inOut.create_df_data_struct(df_final)
        inOut.populate_sheets(df)
        create_IN_OUT_report(wb_adr, inOut.adr_sheet, generate_IN_OUT_report_label(df, constants.LABEL_REPORT_IN_OUT_ADR))
        create_IN_OUT_report(wb_not_adr, inOut.not_adr_sheet, generate_IN_OUT_report_label(df, constants.LABEL_REPORT_IN_OUT_NOT_ADR))


if __name__ == '__main__':
    run_scripts()
    print("IN_OUT_ADR.xlsx and IN_OUT_NON_ADR.xlsx has been generated!")

