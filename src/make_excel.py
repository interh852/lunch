import os
import sys
import pandas as pd
import openpyxl


if __name__ == "__main__":
    # 定数
    PWD = os.getcwd()
    EXCEL_DIR = '{}/../database/shokuraku'.format(PWD)
    GLIDE_DIR = '{}/../glide'.format(PWD)
    DATABASE = '{}/../database/database.xlsx'.format(PWD)

    # 今月のSpreadSheetの読み込み
    excel_list = os.listdir(EXCEL_DIR)
    excel_this_month = excel_list[-1]
    menu = pd.read_excel(os.path.join(
        EXCEL_DIR, excel_this_month), skiprows=3, usecols=[0, 1], header=None).dropna()
    # データの整理
    menu.columns = ['Date', 'Menu']
    menu['Date'] = pd.DatetimeIndex(menu['Date']).tz_localize(
        'Asia/Tokyo').strftime('%Y/%m/%d')
    menu['Order'] = 'False'

    # 各メンバーの来週のメニューリストのSpreadSheetを作成
    members = pd.read_excel(DATABASE).columns.values
    members = [member for member in members if member not in
               ['Date', 'Etc', 'Total']]

    for member in members:
        menu.to_excel('{}/{}.xlsx'.format(GLIDE_DIR, member),
                      sheet_name='MenuList', index=False)
