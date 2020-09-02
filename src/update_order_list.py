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
    db = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['database'])

    # Summary Worksheet of database spreadsheet
    db_worksheet = db.get_worksheet(worksheet_name='Summary')

    # Get the data from Summary Worksheet
    db_df = db.get_dataframe(db_worksheet)

    TODAY = datetime.now(timezone(timedelta(hours=+9), 'JST')).date()
    # TODAY = datetime.strptime('2020/07/23', '%Y/%m/%d').date()

    db_df_today = db_df[(db_df['Date'] >= TODAY + timedelta(days=0)) & (
        db_df['Date'] < TODAY + timedelta(days=5))].drop(['Price', 'Count'], axis=1)

    # Instance of shokuraku spreadsheet
    skrk = orderlunch.Shokuraku(
        path_credentials=PATH_CREDENTIALS, spreadsheet_key=sp_names['shokuraku'])

    # Menu List Worksheet of shokuraku spreadsheet
    skrk_worksheet = skrk.get_worksheet(worksheet_name='Order List Personal')

    skrk_worksheet.clear()

    db_df_today['Date'] = db_df_today['Date'].dt.strftime('%Y/%m/%d')
    skrk_worksheet.update([db_df_today.columns.values.tolist(
    )] + db_df_today.values.tolist(), value_input_option='USER_ENTERED')
