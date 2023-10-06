import pandas
import constants


#("Momento", "Appaltatore", "Tipologia", "Targa", "Sito di giacenza", "Stato")

adr_sheet = {}
not_adr_sheet = {}


def create_data_struct(df):
    #switch data between col1 and col2
    df_temp = df['Momento'].copy()  # Copia i dati di Colonna1 in una variabile temporanea
    df['Momento'] = df['Appaltatore']  # Sovrascrive i dati di Colonna1 con quelli di Colonna2
    df['Appaltatore'] = df_temp  # Sovrascrive i dati di Colonna2 con quelli dalla variabile temporanea

    df = df.rename(columns={'Momento': 'Appaltatore-'})
    df = df.rename(columns={'Appaltatore': 'Momento'})


    # Accorcia il contenuto della prima colonna a 25 caratteri
    df['Appaltatore-'] = df['Appaltatore-'].str.slice(0, 25)
    return df


def populate_sheets(df):
    for group, group_data in df.groupby('Appaltatore-'):
        temp_adr = group_data[group_data['Tipologia'].str.contains('ADR', case=False, na=False)]
        temp_not_adr = group_data[~group_data['Tipologia'].str.contains('ADR', case=False, na=False)]
        if not temp_adr.empty:
             adr_sheet[group] = temp_adr
        if not temp_not_adr.empty:
             not_adr_sheet[group] = temp_not_adr



