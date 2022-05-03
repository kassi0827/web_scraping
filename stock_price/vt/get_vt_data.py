import requests
from bs4 import BeautifulSoup
import openpyxl
import datetime
import xlwings as xw

# 更新対象ファイルのパス
excel_path = '/Users/kashimataichi/play_python/web_scraping/stock_price/vt/vt_data.xlsx'

# ファイルのデータがある最終行と最終列を取得
wb = openpyxl.load_workbook(excel_path)
ws = wb['vt_data']
max_row_plus_one = ws.max_row + 1


# 年月日の転記
dt_now = datetime.datetime.now()
today_yyyy_mm_dd = str(dt_now.year) + '/' + \
    str(dt_now.month) + '/' + str(dt_now.day)
ws.cell(row=max_row_plus_one, column=1).value = today_yyyy_mm_dd


# 対象ページのソースを取得
vt_bloomberg_url = 'https://www.bloomberg.co.jp/quote/VT:US'
res = requests.get(vt_bloomberg_url)

# 取得したページのソースをビューティフルスープでスクレイピング
# 転記の処理はfor文化を検討
soup = BeautifulSoup(res.text, 'html.parser')
elems = soup.find_all('div', attrs={'class': 'cell__value cell__value_'})

# B列始値の転記
ws.cell(row=max_row_plus_one, column=2).value = elems[0].contents[0]

# C列前日終値の転記
ws.cell(row=max_row_plus_one, column=3).value = elems[3].contents[0]

# D列出来高
ws.cell(row=max_row_plus_one, column=4).value = elems[2].contents[0]

# I列日次安値高値レンジ
ws.cell(row=max_row_plus_one, column=9).value = elems[1].contents[0]

# J列52週レンジの転記
ws.cell(row=max_row_plus_one, column=10).value = elems[4].contents[0]

# F列3年トータルリターン
ws.cell(row=max_row_plus_one, column=6).value = elems[19].contents[0]

# G列5年トータルリターン
ws.cell(row=max_row_plus_one, column=7).value = elems[20].contents[0]

# K列純資産額
ws.cell(row=max_row_plus_one, column=11).value = elems[11].contents[0]

elems_others = soup.find_all(
    'div', attrs={'class': 'cell__value cell__value_down'})

# E列1年トータルリターン
ws.cell(row=max_row_plus_one, column=5).value = elems_others[1].contents[0]

# H列年初来リターン
ws.cell(row=max_row_plus_one, column=8).value = elems_others[2].contents[0]

wb.save(excel_path)

xw.Book(r'/Users/kashimataichi/play_python/web_scraping/stock_price/vt/vt_data.xlsx')

print('Transaction completed.')
