import pandas as pd
from datetime import datetime
import logging
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import constants
import utilities


wb_adr_in_out, wb_not_adr_in_out, wb_adr_vob_pob, wb_not_adr_vob_pob = Workbook(), Workbook(), Workbook(), Workbook()
adr_sheet_in_out = {}
not_adr_sheet_in_out = {}
adr_sheet_vob_pob = {}
not_adr_sheet_vob_pob = {}


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


def create_report(wb, sheet, label):
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


def generate_report_label(df, label):
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
    df_in_out = load_xlsx_file(constants.IN_OUT)
    df_vob_pob = load_xlsx_file(constants.VOB_POB)

    if df_in_out is not None:
        df_final_in_out = clean_df(df_in_out)
        df_ultimate_in_out = utilities.create_df_data_struct(df_final_in_out, type_of_report=constants.REPORT_IN_OUT)
        utilities.populate_sheets(df_ultimate_in_out,adr_sheet_in_out, not_adr_sheet_in_out)
        create_report(wb_adr_in_out, adr_sheet_in_out,
                      generate_report_label(df_ultimate_in_out, constants.LABEL_REPORT_IN_OUT_ADR))
        create_report(wb_not_adr_in_out, not_adr_sheet_in_out, generate_report_label(df_ultimate_in_out,
                                                                                     constants.LABEL_REPORT_IN_OUT_NOT_ADR))
    if df_vob_pob is not None:
        df_final_vob_pob = clean_df(df_vob_pob)
        df_ultimate_vob_pob = utilities.create_df_data_struct(df_final_vob_pob, type_of_report=constants.REPORT_VOB_POB)
        utilities.populate_sheets(df_ultimate_vob_pob, adr_sheet_vob_pob, not_adr_sheet_vob_pob)
        print(df_ultimate_vob_pob.head())
        create_report(wb_adr_vob_pob, adr_sheet_vob_pob, generate_report_label(df_ultimate_vob_pob,
                                                                               constants.LABEL_REPORT_VOB_POB_ADR))
        create_report(wb_not_adr_vob_pob, not_adr_sheet_vob_pob, generate_report_label(df_ultimate_vob_pob,
                                                                                       constants.LABEL_REPORT_VOB_POB_NOT_ADR))


if __name__ == '__main__':
    run_scripts()
    print("IN_OUT_ADR.xlsx and IN_OUT_NON_ADR.xlsx has been generated!")

