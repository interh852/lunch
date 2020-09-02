import gspread
import json
import jpholiday
import pandas as pd
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials


class Shokuraku:
    def __init__(self, path_credentials, spreadsheet_key):
        # Describes two APIs for spreadsheet and googledrive
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/drive']

        # Set the authentication information obtained from GCP
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
            path_credentials, self.scope)

        # Read the google spreadsheet key
        self.spreadsheet_key = spreadsheet_key

    def get_worksheet(self, worksheet_name):
        """
        Get a google spreadsheet worksheet
        """

        # Login to google spreadsheet
        gc = gspread.authorize(self.credentials)

        # Open the Menu sheet in the Shokulaku spreadsheet
        worksheet = gc.open_by_key(
            self.spreadsheet_key).worksheet(worksheet_name)

        return(worksheet)

    def get_dataframe(self, worksheet):
        """
        Fetch all values in google spreadsheet and convert a data frame
        """

        data_cells = worksheet.get_all_values()
        data_cells_df = pd.DataFrame(data_cells, columns=data_cells.pop(0))
        data_cells_df['Date'] = pd.to_datetime(data_cells_df['Date'])

        return(data_cells_df)

    def get_todays_task(self, worksheet):
        """
        Get today's task
        """

        # Get the data from Menu List Worksheet
        df_task = skrk.get_dataframe(worksheet)

        # Substract today's Task
        TODAY = datetime.now(timezone(timedelta(hours=+9), 'JST')).date()
        df_task_today = df_task[df_task['Date']
                                == TODAY].reset_index(drop=True)

        return(df_task_today)

    def convert_order_list_to_database(self, df_order_list):
        """
        conver from Order List to DataBase
        """

        df_database = df_order_list.melt(var_name='Name', value_name='Count', id_vars=[
                                         'Date', 'Menu', 'Price'])
        df_database['Date'] = pd.to_datetime(df_database['Date'])
        df_database = df_database[df_database['Count']
                                  == 'TRUE'].replace('TRUE', 1)
        df_database = df_database[['Date', 'Name', 'Menu', 'Price', 'Count']].sort_values(
            by=['Date', 'Name']).reset_index(drop=True)

        return(df_database)

    def tidy_lunch_menu_a_week(self, df):
        """
        tidy the lunch menu for a week
        """

        # Date
        df_date = df[0:1].T.dropna()
        df_date.rename(columns=dict(
            zip(df_date.iloc[:, [0]].columns, ['Date'])), inplace=True)
        df_date = df_date[df_date['Date'] != 0].reset_index(
            drop=True).rename(index=lambda x: x + 1)

        # Menu
        df_menu = df[1:]
        df_menu_long = pd.concat([df_menu.iloc[:, [0, 1]].rename(index=lambda x: 1, columns=dict(zip(df_menu.iloc[:, [0, 1]].columns, ['Menu', 'Price']))),
                                  df_menu.iloc[:, [2, 3]].rename(index=lambda x: 2, columns=dict(
                                      zip(df_menu.iloc[:, [2, 3]].columns, ['Menu', 'Price']))),
                                  df_menu.iloc[:, [4, 5]].rename(index=lambda x: 3, columns=dict(
                                      zip(df_menu.iloc[:, [4, 5]].columns, ['Menu', 'Price']))),
                                  df_menu.iloc[:, [6, 7]].rename(index=lambda x: 4, columns=dict(
                                      zip(df_menu.iloc[:, [6, 7]].columns, ['Menu', 'Price']))),
                                  df_menu.iloc[:, [8, 9]].rename(index=lambda x: 5, columns=dict(zip(df_menu.iloc[:, [8, 9]].columns, ['Menu', 'Price'])))])

        df_long = pd.concat([df_date, df_menu_long], axis=1)

        return(df_long)

    def is_holiday(self, date):
        holidays = ['01/01', '01/02', '01/03', '08/13',
                    '08/14', '08/15', '12/30', '12/31']
        date_str = '{}/{}'.format(date.strftime('%m'), date.strftime('%d'))

        if (date_str in holidays):
            return True
        else:
            return(jpholiday.is_holiday(date))
