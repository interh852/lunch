import os
import json
import pandas as pd
from module import orderlunch

PWD = os.getcwd()

# Set the authentication information obtained from GCP
PATH_CREDENTIALS = '{}/../secret/order-lunch-project.json'.format(PWD)

# Read the google spreadsheet keys
with open('../secret/sp_name.json') as f:
    sp_names = json.load(f)

if __name__ == "__main__":

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
