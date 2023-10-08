import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows
import constants
import utilities

pd.set_option('display.max_columns', None)


wb_pdl_check = Workbook()


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
    # Creare un DataFrame vuoto
    df_excel = pd.DataFrame()

    # Inserire le tipologie di attività nella prima colonna
    df_excel['Tipologia attività'] = prioritized_types

    # Iterare attraverso le "macro aree" univoche
    unique_macro_areas = df_original['Macro area'].unique()

    for macro_area in unique_macro_areas:
        # Creare una colonna separata per ogni "macro area"
        df_excel[macro_area] = ""

    # Eseguire il conteggio delle occorrenze
    counts_df = df_original.groupby(['Macro area', 'Tipologia attività']).size().reset_index(name='Counts')

    # Aggiungere i conteggi al DataFrame Excel
    for index, row in counts_df.iterrows():
        macro_area = row['Macro area']
        activity_type = row['Tipologia attività']
        count = row['Counts']

        df_excel.loc[df_excel['Tipologia attività'] == activity_type, macro_area] = count


    df_excel = df_excel.replace('', 0)
    # Calcola e aggiungi la riga con la somma delle colonne (tralasciando la prima colonna)
    sums = df_excel.iloc[:, 1:].sum()
    df_sums = pd.DataFrame([['TOTALE'] + sums.tolist()], columns=df_excel.columns)
    df_excel = pd.concat([df_excel, df_sums], ignore_index=True)

    df_excel.to_excel(output_file, index=False)
    print(df_excel)


def run_scripts_report_pdl_check():
    df_pdl_check = utilities.load_xlsx_file(constants.PDL_CHECK)
    df_final_pdl_check = utilities.clean_df(df_pdl_check)
    df_temp_pdl_check = create_df_data_struct_report_pdl_check(df_final_pdl_check)
    df_ultimate_in_out = df_temp_pdl_check.copy()
    df_ultimate_in_out['Tipologia attività'] = df_ultimate_in_out['Tipologia attività'].\
        apply(replace_comma_specific_part)
    df_ultimate_in_out["Tipologia attività"] = df_ultimate_in_out["Tipologia attività"].\
        apply(lambda x: find_priority_type(x, priority_list=constants.LIST_PRIORITY_PDL_AND_CHECK))
    create_excel_sheet(constants.LIST_PRIORITY_PDL_AND_CHECK, df_ultimate_in_out, 'assets/report_pdl_check_pdl/reportPdlCheck.xlsx')




    print(df_ultimate_in_out.head())
