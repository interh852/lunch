import os
import json
import pandas as pd
from module import orderlunch
from datetime import datetime, timezone, timedelta

PWD = os.getcwd()

# Set the authentication information obtained from GCP
PATH_CREDENTIALS = '{}/../secret/order-lunch-project.json'.format(PWD)

# Read the google spreadsheet keys
with open('../secret/sp_name.json') as f:
    sp_names = json.load(f)

if __name__ == "__main__":

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
