
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import constants
import inOut


wb_adr, wb_not_adr = Workbook()



def load_xlsx_file(file_excel):
    #load a xlsx file in a DF
    df = pd.read_excel(file_excel)
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


#TODO: review this parto of code

def create_IN_OUT_report(wb, sheets):
    # Scrivi i dati su fogli separati nei due file Excel
    write_on_xlsx_sheet_file(wb_adr, fogli_con_adr)
    write_on_xlsx_sheet_file(wb_not_adr, fogli_senza_adr)

    # Rimuovi il foglio di lavoro predefinito "Sheet" da entrambi i file
    wb_adr.remove(wb_adr['Sheet'])
    wb_not_adr.remove(wb_not_adr['Sheet'])

    # Salva i due nuovi file Excel
    wb_adr.save('C:/Users/Maersk/Desktop/testpy/INOUT_ADR.xlsx')
    wb_not_adr.save('C:/Users/Maersk/Desktop/testpy/INOUT_NON_ADR.xlsx')

#------------------------------------------------------------------------------

# let's go work on In_OUT DFs
df_1 = load_xlsx_file(constants.IN_OUT)
df_final = clean_df(df_1)

df = inOut.create_data_struct(df_final)
df = inOut.populate_sheets(df)





if __name__ == '__main__':
    print(df)

