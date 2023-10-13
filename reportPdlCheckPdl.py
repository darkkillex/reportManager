import re
import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
import constants
import utilities

pd.set_option('display.max_columns', None)

def create_df_data_struct_report_pdl_check(df):
    df = df[["Macro area", "Tipologia attività", "Esito check"]]
    return df


def replace_comma_specific_part(string):
    if "CND (spessom,liquidi pen.,tecnografie,ultrasuoni)" in string:
        # Trova la parte specifica tra parentesi
        specific_part = re.search(r'CND \(spessom,liquidi pen.,tecnografie,ultrasuoni\)', string).group(0)
        # Sostituisci le virgole con la barra "/" solo nella parte specifica
        specific_part_replaced = specific_part.replace(',', '/')
        # Sostituisci la parte specifica nella stringa originale
        string = string.replace(specific_part, specific_part_replaced)
    return string


def find_priority_type(types, priority_list):
    types = types.split(",")  # Split the types separated by commas into a list
    priority_type = None

    for priority in priority_list:
        for type in types:
            if type.strip() == priority:
                priority_type = priority
                break
        if priority_type:
            break

    return priority_type


def create_excel_sheet(prioritized_types, df_original, output_file):

    wb = Workbook()

    # Creare fogli Excel separati per ciascuna tipologia di "Esito"
    for esito in df_original['Esito check'].unique():
        ws = wb.create_sheet(title=esito)

        df_excel = pd.DataFrame()

        # Filtrare il DataFrame originale per la tipologia di "Esito"
        df_filtered = df_original[df_original['Esito check'] == esito]

        # Inserire le tipologie di attività nella prima colonna
        df_excel['Tipologia attività'] = prioritized_types

        # Iterare attraverso le "macro aree" univoche
        all_macro_areas = df_original['Macro area'].unique()

        # Inserisci tutte le macro aree come colonne, inizializzate a 0
        for macro_area in all_macro_areas:
            df_excel[macro_area] = 0

        # Eseguire il conteggio delle occorrenze
        counts_df = df_filtered.groupby(['Macro area', 'Tipologia attività']).size().reset_index(name='Counts')

        # Aggiungere i conteggi al DataFrame Excel
        for index, row in counts_df.iterrows():
            macro_area = row['Macro area']
            activity_type = row['Tipologia attività']
            count = row['Counts']

            # Verifica se c'è una corrispondenza tra "macro area" e "tipologia attività"
            if macro_area in df_excel.columns and activity_type in df_excel['Tipologia attività'].values:
                df_excel.loc[df_excel['Tipologia attività'] == activity_type, macro_area] = count

        df_excel = df_excel.replace('', 0)
        # Calcola e aggiungi la riga con la somma delle colonne (tralasciando la prima colonna)
        sums = df_excel.iloc[:, 1:].sum()
        df_sums = pd.DataFrame([['TOTALE'] + sums.tolist()], columns=df_excel.columns)
        df_excel = pd.concat([df_excel, df_sums], ignore_index=True)
        for r in dataframe_to_rows(df_excel, index=False, header=True):
            ws.append(r)

    # Rimuovi il foglio di lavoro predefinito
    wb.remove(wb.active)

    # Salva il Workbook completo in un file Excel
    wb.save(output_file)




def generate_report_label(df):
    # Extract the first column containing the dates as strings
    if not df.empty:
        today = datetime.date.today()
        week = today.strftime("%U")  # Numero della settimana
        year = today.strftime("%Y")  # Anno corrente
        report_label = "assets/report_pdl_check_pdl/" + constants.LABEL_REPORT_PDL_CHECK + week + "-" + year + ".xlsx"

    return report_label




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
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        # Elimina la prima riga (Header)
        sheet.delete_rows(1)
        # Imposta i filtri per tutte le colonne
        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = 'A2'
    wb.save(label)





def run_scripts_report_pdl_check():
    df_pdl_check = utilities.load_xlsx_file(constants.PDL_CHECK)
    df_final_pdl_check = utilities.clean_df(df_pdl_check)
    df_temp_pdl_check = create_df_data_struct_report_pdl_check(df_final_pdl_check)
    df_ultimate_in_out = df_temp_pdl_check.copy()
    df_ultimate_in_out['Tipologia attività'] = df_ultimate_in_out['Tipologia attività'].\
        apply(replace_comma_specific_part)
    df_ultimate_in_out["Tipologia attività"] = df_ultimate_in_out["Tipologia attività"].\
        apply(lambda x: find_priority_type(x, priority_list=constants.LIST_PRIORITY_PDL_AND_CHECK))
    create_excel_sheet(constants.LIST_PRIORITY_PDL_AND_CHECK, df_ultimate_in_out, generate_report_label(df_ultimate_in_out))
