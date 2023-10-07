import constants



def create_df_data_struct(df, type_of_report):
    if type_of_report == "IN_OUT":
        #switch data between col1 and col2
        temp_IN_OUT = df['Momento'].copy()  # Copia i dati di Colonna1 in una variabile temporanea
        df['Momento'] = df['Appaltatore']  # Sovrascrive i dati di Colonna1 con quelli di Colonna2
        df['Appaltatore'] = temp_IN_OUT  # Sovrascrive i dati di Colonna2 con quelli dalla variabile temporanea
        df = df.rename(columns={'Momento': 'Azienda Appaltatrice'})
        df = df.rename(columns={'Appaltatore': 'Momento'})
    if type_of_report == "VOB_POB":
        df = df[constants.LIST_OF_LABELS_VOB_POB]
        df = df.rename(columns={'Appaltatore': 'Azienda Appaltatrice'})


    # cuts the number of characters in the excel sheet label (max 31 char supported)
    df['Azienda Appaltatrice'] = df['Azienda Appaltatrice'].str.slice(0, 31)
    return df


def populate_sheets(df, adr_sheet, not_adr_sheet):
    for group, group_data in df.groupby('Azienda Appaltatrice'):
        temp_adr = group_data[group_data['Tipologia'].str.contains('ADR', case=False, na=False)]
        temp_not_adr = group_data[~group_data['Tipologia'].str.contains('ADR', case=False, na=False)]
        if not temp_adr.empty:
             adr_sheet[group] = temp_adr
        if not temp_not_adr.empty:
             not_adr_sheet[group] = temp_not_adr



