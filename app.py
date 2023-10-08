import pandas as pd
from datetime import datetime
import logging
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import constants
import utilities
import in_out_vob as iov


def run_scripts():
    # let's go work on DFs
    df_in_out = iov.load_xlsx_file(constants.IN_OUT)
    df_vob_pob = iov.load_xlsx_file(constants.VOB_POB)

    if df_in_out is not None:
        df_final_in_out = iov.clean_df(df_in_out)
        df_ultimate_in_out = utilities.create_df_data_struct(df_final_in_out, type_of_report=constants.REPORT_IN_OUT)
        utilities.populate_sheets(df_ultimate_in_out,iov.adr_sheet_in_out, iov.not_adr_sheet_in_out)
        iov.create_report(iov.wb_adr_in_out, iov.adr_sheet_in_out,
                      iov.generate_report_label(df_ultimate_in_out, constants.LABEL_REPORT_IN_OUT_ADR))
        iov.create_report(iov.wb_not_adr_in_out, iov.not_adr_sheet_in_out, iov.generate_report_label(df_ultimate_in_out,
                                                                                     constants.LABEL_REPORT_IN_OUT_NOT_ADR))
    if df_vob_pob is not None:
        df_final_vob_pob = iov.clean_df(df_vob_pob)
        df_ultimate_vob_pob = utilities.create_df_data_struct(df_final_vob_pob, type_of_report=constants.REPORT_VOB_POB)
        utilities.populate_sheets(df_ultimate_vob_pob, iov.adr_sheet_vob_pob, iov.not_adr_sheet_vob_pob)
        #print(df_ultimate_vob_pob.head())
        iov.create_report(iov.wb_adr_vob_pob, iov.adr_sheet_vob_pob, iov.generate_report_label(df_ultimate_in_out,#in this case we get the df_in_out to get the right date of extraction
                                                                               constants.LABEL_REPORT_VOB_POB_ADR))
        iov.create_report(iov.wb_not_adr_vob_pob, iov.not_adr_sheet_vob_pob, iov.generate_report_label(df_ultimate_in_out,#in this case we get the df_in_out to get the right date of extraction
                                                                                       constants.LABEL_REPORT_VOB_POB_NOT_ADR))


if __name__ == '__main__':
    try:
        run_scripts()
    except Exception as e:
        print(f"Si è verificato un errore: {str(e)}")
    else:
        print("IN_OUT and VOB reports have been generated in the /assets folder!")

