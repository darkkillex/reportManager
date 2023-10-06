# This is a sample Python script.

# Press Maiusc+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import constants


def load_xlsx_file(file_excel):
    #load a xlsx file in a DF
    df = pd.read_excel(file_excel)
    return df





if __name__ == '__main__':
    print_hi('PyCharm')


