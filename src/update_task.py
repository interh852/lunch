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

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Menu List')

    # Menu List Worksheet of shokuraku spreadsheet
    df = skrk.get_dataframe(skrk_worksheet)
    df['Date'] = pd.to_datetime(df['Date'])
    df['day_name'] = pd.to_datetime(df['Date']).dt.day_name()
    df['Date'] = df['Date'].dt.date

    # drop duplicated schedule
    df['is_holiday'] = df['Date'].map(skrk.is_holiday).astype(int)
    df = df.loc[:, ['Date', 'is_holiday', 'day_name']
                ].drop_duplicates(subset='Date')
    df.reset_index(inplace=True)

    # update monthly menu
    df['update_monthly_menu'] = np.nan
    df['update_monthly_menu'].iloc[-5] = 1

    # check order
    df.loc[df['day_name'] == 'Monday', 'check_order'] = 1

    # update order list
    df.loc[df['day_name'] == 'Monday', 'update_order_list'] = 1

    # update database
    df_weekday = df.copy()
    df_weekday['is_holiday_lag'] = df_weekday['is_holiday'] - \
        df_weekday['is_holiday'].shift(-1)
    df_weekday.loc[(df_weekday['day_name'] == 'Friday') & (
        df_weekday['is_holiday'] == 0) & (df_weekday['is_holiday_lag'] == 0), 'update_db'] = 1
    df_weekday.loc[(df_weekday['is_holiday'] == 0) & (
        df_weekday['is_holiday_lag'] == -1), 'update_db'] = 1
    df_weekday['update_db'] = df_weekday['update_db'].shift(-1)
    df_weekday['update_db'].iloc[-2] = 1
    df_weekday = df_weekday[df_weekday['is_holiday']
                            == 0].loc[:, ['update_db']]

    df_task = df.join(df_weekday).drop(['index', 'day_name'], axis=1)
    df_task = df_task.fillna(0)

    df_task[df_task.select_dtypes(['float64']).columns] = df_task.select_dtypes([
        'float64']).apply(lambda x: x.astype('int16'))

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Task')

    df_task['Date'] = df_task['Date'].apply(lambda x: x.strftime('%Y/%m/%d'))
    skrk_worksheet.update([df_task.columns.values.tolist(
    )] + df_task.values.tolist(), value_input_option='USER_ENTERED')
