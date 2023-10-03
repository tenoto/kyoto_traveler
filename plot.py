#!/usr/bin/env python

# excel のファイル "001631648.xls" を読み込み、"1-2" という名前のシートで、A行が"26京都府"である行をB列以降を順番に読み込み、np.array として格納するコードを python で書いてください。

import xlrd
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import smplotlib

def remove_whitespace(s):
    return ''.join(s.split())

# ファイルパス
file_path = "001631648.xls"

# Excelファイルを開く
workbook = xlrd.open_workbook(file_path)

# シートを取得
sheet = workbook.sheet_by_name("1-2")

# 行数
num_rows = sheet.nrows

# 列数
num_cols = sheet.ncols

data_np = None

# "26京都府"がA列にある行を検索
for i in range(num_rows):
    if sheet.cell_value(i, 0) == "26京都府":
        # B列以降のデータを取得し、NumPy配列に変換
        data_np = np.array(sheet.row_values(i, start_colx=1))

# データが見つからない場合
if data_np is None:
    print("Data not found for '26京都府'.")
else:
    print("Data for '26京都府':")
    print(data_np)

# python の np.array で '' （空白文字列）を削除してください。
kyoto_traveler = (data_np[data_np != '']).astype(float)/1e+6
print(kyoto_traveler)

# python で「月ごとの来客数」をプロットするサンプルコードを書いてください。月の範囲は、2011年1月から2020年12月までとします。

months = pd.date_range(start="2011-01-01", end="2023-07-01", freq='MS')
plt.figure(figsize=(12, 6))
plt.plot(months, kyoto_traveler, marker='o', color='r',markersize=3)
plt.xlabel('Year/Month')
plt.ylabel('Number of Visitors at Kyoto / $10^6$ travelers')
plt.title('Monthly Visitors (Jan 2011 - July 2023)')
plt.grid(True)
plt.tight_layout()
plt.xlim(pd.to_datetime('2011-04-01'), pd.to_datetime('2025-12-31'))
plt.ylim(0.0,3.5)
# python で時系列データを表示する際に、x軸が 2023-01 のような年・月でヒョじされているときに、x軸の表示範囲を指定する方法を教えて下さい。

# python の matplotlib で表示しているマーカーを filled circule にする方法を教えてください。

# グラフを表示
# python で matplotlib で作成した図を pdf で出力するコードを書いてください。
plt.savefig('kyoto_traveler.pdf')

"""
python で以下のコードを書いてください。

まず、以下の２つの np.array が用意されています。
months = pd.date_range(start="2011-01-01", end="2023-07-01", freq='MS')
kyoto_traveler　= np.random.randint(100, 1000, size=len(months))
months は年と月の時刻を、kyoto_traveler は月ごとの訪問者数を示しています。

ここで、訪問者数の推移を X軸は１月から１２月までの１２個に分け、毎年の違いは凡例で区別して表示してください。
"""

cmap = plt.get_cmap("tab20") 

# データを年ごとに分割
data_per_year = {}
for i in range(len(months)):
    year = months[i].year
    if year not in data_per_year:
        data_per_year[year] = []
    data_per_year[year].append(kyoto_traveler[i])

# 月ごとの訪問者数をプロット
plt.figure(figsize=(12, 6))
i = 0
for year, year_data in data_per_year.items():
	plt.plot(np.arange(1, len(year_data) + 1), 
    	year_data, label=f'{year}',
    	color=cmap(i),
    	marker='o',markersize=3)
	i += 1

plt.xlabel('Months')
plt.ylabel('Number of Visitors at Kyoto / $10^6$ travelers')
plt.title('Monthly Visitors (Jan 2011 - July 2023)')
#plt.legend()
plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
plt.grid(True)
plt.tight_layout()

# グラフを表示
plt.savefig('kyoto_traveler_month.pdf')
