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
    SP_NAMES = json.load(f)
EXCEL_DIR = '{}/../database/shokuraku'.format(PWD)


def update_monthly_menu(PATH_CREDENTIALS, SP_NAMES):
    """update monthly menu

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """
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


def update_task(PATH_CREDENTIALS, SP_NAMES):
    """update task

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """
    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Menu List')

    # Menu List Worksheet of shokuraku spreadsheet
    df = skrk.get_dataframe(skrk_worksheet)
    df['Date'] = pd.to_datetime(df['Date']).dt.date

    # drop duplicated schedule
    df['is_holiday'] = df['Date'].map(skrk.is_holiday).astype(int)
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

    df_task = df.join(df_weekday).drop(['index'], axis=1)
    df_task = df_task.fillna(0)

    df_task[df_task.select_dtypes(['float64']).columns] = df_task.select_dtypes([
        'float64']).apply(lambda x: x.astype('int16'))

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Task')

    df_task['Date'] = df_task['Date'].apply(lambda x: x.strftime('%Y/%m/%d'))
    skrk_worksheet.update([df_task.columns.values.tolist(
    )] + df_task.values.tolist(), value_input_option='USER_ENTERED')

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Task')

    df_task['Date'] = df_task['Date'].apply(lambda x: x.strftime('%Y/%m/%d'))
    skrk_worksheet.update([df_task.columns.values.tolist(
    )] + df_task.values.tolist(), value_input_option='USER_ENTERED')


def update_check_order(PATH_CREDENTIALS, SP_NAMES):
    """update_check_order

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """
    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Menu List')

    # Get the data from Menu List Worksheet
    skrk_df_menu_monthly = skrk.get_dataframe(skrk_worksheet)

    TODAY = datetime.now(timezone(timedelta(hours=+9), 'JST')).date()
    # TODAY = datetime.strptime('2020/07/22', '%Y/%m/%d').date()

    skrk_df_menu_weekly = skrk_df_menu_monthly[(skrk_df_menu_monthly['Date'] >= TODAY + timedelta(
        days=7)) & (skrk_df_menu_monthly['Date'] < TODAY + timedelta(days=12))].loc[:, ['Date', 'Menu', 'Price']]

    # Glide UI Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Check Order')

    # Reset Glide UI Worksheet
    cell_list = skrk_worksheet.range('A2:C31')
    for cell in cell_list:
        cell.value = ''
    skrk_worksheet.update_cells(cell_list, value_input_option='USER_ENTERED')

    cell_list = skrk_worksheet.range('D2:J31')
    for cell in cell_list:
        cell.value = 'False'
    skrk_worksheet.update_cells(cell_list, value_input_option='USER_ENTERED')

    # Update Glide UI Worksheet
    skrk_df_menu_weekly['Date'] = skrk_df_menu_weekly['Date'].dt.strftime(
        '%Y/%m/%d')
    skrk_worksheet.update([skrk_df_menu_weekly.columns.values.tolist(
    )] + skrk_df_menu_weekly.values.tolist(), value_input_option='USER_ENTERED')


def update_db(PATH_CREDENTIALS, SP_NAMES):
    """update_db

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """
    # Most recent OrderList
    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Menu Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Check Order')

    # Get the data from Menu Worksheet
    skrk_df_order_list = skrk.get_dataframe(skrk_worksheet)

    # Format the most recent Order List to save it in the database
    skrk_df_database_latest = skrk.convert_order_list_to_database(
        skrk_df_order_list)

    # DataBase
    # Instance of shokuraku spreadsheet
    db = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['database'])

    # Summary Worksheet of database spreadsheet
    db_worksheet = db.get_worksheet(worksheet_name='Summary')

    # Get the data from Summary Worksheet
    db_df = db.get_dataframe(db_worksheet)

    # Update the most recent OrderList in the database spreadsheet
    START_CELL = 'A' + str(len(db_df) + 2)
    db_df_update = skrk_df_database_latest.copy()
    db_df_update['Date'] = db_df_update['Date'].dt.strftime('%Y/%m/%d')
    db_worksheet.update(START_CELL, db_df_update.values.tolist(),
                        value_input_option='USER_ENTERED')


def update_order_list(PATH_CREDENTIALS, SP_NAMES):
    """update_order_list

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """
    # Instance of shokuraku spreadsheet
    db = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['database'])

    # Summary Worksheet of database spreadsheet
    db_worksheet = db.get_worksheet(worksheet_name='Summary')

    # Get the data from Summary Worksheet
    db_df = db.get_dataframe(db_worksheet)

    TODAY = datetime.now(timezone(timedelta(hours=+9), 'JST')).date()
    # TODAY = datetime.strptime('2020/07/23', '%Y/%m/%d').date()

    db_df_today = db_df[(db_df['Date'] >= TODAY + timedelta(days=3)) & (
        db_df['Date'] < TODAY + timedelta(days=8))].drop(['Price', 'Count'], axis=1)

    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Order List Personal')

    skrk_worksheet.clear()

    db_df_today['Date'] = db_df_today['Date'].dt.strftime('%Y/%m/%d')
    skrk_worksheet.update([db_df_today.columns.values.tolist(
    )] + db_df_today.values.tolist(), value_input_option='USER_ENTERED')


def main_func(PATH_CREDENTIALS, SP_NAMES):
    """main function of lunch order system

    Args:
        PATH_CREDENTIALS (dict): authentication information obtained from GCP
        SP_NAMES (dict): the google spreadsheet keys
    """

    # Task Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Task')

    # today's task
    todays_task = get_todays_task(skrk_worksheet)

    if todays_task['update_monthly_menu'] == 1:
        update_monthly_menu(PATH_CREDENTIALS, SP_NAMES)

    if todays_task['check_order'] == 1:
        update_check_order(PATH_CREDENTIALS, SP_NAMES)

    if todays_task['update_order_list'] == 1:
        update_order_list(PATH_CREDENTIALS, SP_NAMES)

    if todays_task['update_db'] == 1:
        update_db(PATH_CREDENTIALS, SP_NAMES)
