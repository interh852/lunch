import os
import sys
import json
import openpyxl
import numpy as np
import pandas as pd
from module import orderlunch

PWD = os.getcwd()

# Set the authentication information obtained from GCP
PATH_CREDENTIALS = '{}/../secret/order-lunch-project.json'.format(PWD)

# Read the google spreadsheet keys
with open('../secret/sp_name.json') as f:
    sp_names = json.load(f)
EXCEL_DIR = '{}/../database/shokuraku'.format(PWD)

if __name__ == "__main__":

    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Substruct this month's daily lunch menu
    excel_list = os.listdir(EXCEL_DIR)
    excel_this_month = excel_list[-1]
    lunch_menu = pd.read_excel(os.path.join(
        EXCEL_DIR, excel_this_month), skiprows=46, skipfooter=2, usecols=list(range(12, 36)), header=None).dropna(how='all', axis=1)

    menu_excel = pd.DataFrame(index=[], columns=['Date', 'Menu', 'Price'])
    for i in list(range(0, 45, 11)):
        menu_excel = pd.concat(
            [menu_excel, skrk.tidy_lunch_menu_a_week(lunch_menu[i:(i + 7)])])
    menu_excel['Date'] = menu_excel['Date'].dt.date

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Menu List')

    menu_excel['Date'] = menu_excel['Date'].apply(
        lambda x: x.strftime('%Y/%m/%d'))
    skrk_worksheet.update([menu_excel.columns.values.tolist(
    )] + menu_excel.values.tolist(), value_input_option='USER_ENTERED')
