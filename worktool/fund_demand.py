# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import locale

try:
	data = pd.read_excel('data.xlsx',encoding='utf8')
	df = pd.DataFrame(data)
	df.columns = df.columns.map(str.strip)
except:
	print('Error: 请提供\'data.xlsx\'文件')
	raise SystemExit

locale.setlocale(locale.LC_ALL,'')
exchange = {
	'HKD':'0.9',
	'USD':'6.92',
	'EUR':'7.4',
	'GBP':'8.62',
	'JPY':'0.062'
}

def get_Exchange(current):
	return float(exchange[current])

def RMB_format(value):
	return (float('%.2f'%value))

def pay_this_week(data,month_start,day_start,month_end,day_end,*,year=2017):
	df = data[(data['出纳付款时间'].isnull()) & (data['要求付款时间'] >= pd.datetime(year,month_start,day_start)) & (data['要求付款时间'] <= pd.datetime(year,month_end,day_end))]
	print('本周总付款订单供%d笔'%len(df))

	df['RMB'] = df['币种'].map(get_Exchange) * df['要求付款金额']
	df['RMB'] = df['RMB'].map(RMB_format)
	return df
# print(len(df1))
if __name__ == '__main__':
	tw_df = pay_this_week(df,4,15,4,22)
	# tw_df.to_excel('result.xlsx')
	# print(type(RMB_format(123)))
	print(tw_df['RMB'].sum())