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

    # Create a schedule for the current month's work
    df = menu_excel.copy()
    df = df.loc[:, ['Date', 'is_holiday']].drop_duplicates(subset='Date')
    df.reset_index(inplace=True)

    # update monthly menu
    df['update_monthly_menu'] = np.nan
    df['update_monthly_menu'].iloc[-5] = 1

    # check order
    df.loc[df['index'] == 1, 'check_order'] = 1

    # update order list
    df.loc[df['index'] == 5, 'update_order_list'] = 1

    # update database
    df_weekday = df.copy()
    df_weekday['is_holiday'] = df_weekday['Date'].map(
        skrk.is_holiday).astype(int)
    df_weekday['is_holiday_lag'] = df_weekday['is_holiday'] - \
        df_weekday['is_holiday'].shift(-1)
    df_weekday.loc[(df_weekday['index'] == 5) & (df_weekday['is_holiday'] == 0) & (
        df_weekday['is_holiday_lag'] == 0), 'update_db'] = 1
    df_weekday.loc[(df_weekday['is_holiday'] == 0) & (
        df_weekday['is_holiday_lag'] == -1), 'update_db'] = 1
    df_weekday['update_db'] = df_weekday['update_db'].shift(-1)
    df_weekday['update_db'].iloc[-2] = 1
    df_weekday = df_weekday[df_weekday['is_holiday']
                            == 0].loc[:, ['update_db']]

    df = df.join(df_weekday).drop(['index'], axis=1)
    df = df.fillna(0)

    menu_list = pd.merge(menu_excel, df, on='Date', how='left')
    menu_list[menu_list.select_dtypes(['float64']).columns] = menu_list.select_dtypes([
        'float64']).apply(lambda x: x.astype('int16'))

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Menu List')

    menu_list['Date'] = menu_list['Date'].apply(
        lambda x: x.strftime('%Y/%m/%d'))
    skrk_worksheet.update([menu_list.columns.values.tolist(
    )] + menu_list.values.tolist(), value_input_option='USER_ENTERED')
