# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os

try:
	data = pd.read_excel('data.xlsx',encoding='utf8')
	df = pd.DataFrame(data)
	df.columns = df.columns.map(str.strip)
except:
	print('Error: 请提供\'data.xlsx\'文件')
	raise SystemExit

# locale.setlocale(locale.LC_ALL,'')
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

#汇总本周需要付款的订单详情
def pay_this_week(data,month_start,day_start,month_end,day_end,*,year=2017):
	df = data[(data['出纳付款时间'].isnull()) & (data['要求付款时间'] >= pd.datetime(year,month_start,day_start)) & (data['要求付款时间'] <= pd.datetime(year,month_end,day_end))]
	del df['分组：步骤号']
	# df['要求付款时间'] = pd.to_datetime(df['要求付款时间'],format='%Y%m')
	df.index = df['flow_num']
	print('本周总付款订单供%d笔'%len(df))

	df['RMB'] = df['币种'].map(get_Exchange) * df['要求付款金额']
	df['RMB'] = df['RMB'].map(RMB_format)
	return df
	

#对数据进行透视
def paymt_pivot(df):

	#获取总体付款数据
	pv_all = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],aggfunc=[np.sum])
	pv_all.columns = ['应付总额']
	pv_all['date'] = pv_all.index

	#获取账期数据
	pv_credit = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],columns=['账期'],aggfunc=[np.sum],fill_value=0)
	pv_isCredit = pv_credit['sum']['RMB']['是'] 
	df_pv_isCredit = pd.DataFrame(pv_isCredit)
	df_pv_isCredit.columns = ['账期资金']
	df_pv_isCredit['date'] = df_pv_isCredit.index

	#获取ZD数据
	ZD_capital = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],columns=['资金来源'],aggfunc=[np.sum],fill_value=0)
	pv_isZD = ZD_capital['sum']['RMB']['中电资金'] 
	df_pv_isZD = pd.DataFrame(pv_isZD)
	df_pv_isZD.columns = ['中电资金']
	df_pv_isZD['date'] = df_pv_isZD.index

	#获取ZD赎货数据
	ZD_redeem = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],columns=['资金来源'],aggfunc=[np.sum],fill_value=0)
	pv_isRedeem = ZD_redeem['sum']['RMB']['中电赎货'] 
	df_pv_redeem = pd.DataFrame(pv_isRedeem)
	df_pv_redeem.columns = ['中电赎货']
	df_pv_redeem['date'] = df_pv_redeem.index

	#合并处理的数据并以日期列作为索引
	pv = pd.merge(pd.merge(pd.merge(pv_all,df_pv_redeem),df_pv_isZD),df_pv_isCredit,how='left',on='date')
	date = pv.pop('date')
	pv.insert(0,'date',date)
	pv.index = pv['date']
	pv.index.name = '日期'
	del pv['date']
	return pv

def output(df,month_start,day_start,month_end,day_end,*,year=2017):
	tw = paymt_pivot(pay_this_week(df,month_start,day_start,month_end,day_end,year=year))
	tw['跨境资金'] = tw['应付总额'] - tw['中电赎货']
	tw.ix['总计'] = {'应付总额':tw['应付总额'].sum(),'中电赎货':tw['中电赎货'].sum(),'中电资金':tw['中电资金'].sum(),'账期资金':tw['账期资金'].sum(),'跨境资金':tw['跨境资金'].sum()}
	tw.index.name = '日期'
	tw['date'] = tw.index
	tw['date'] = tw['date'].apply(lambda time:pd.Period(time,freq='D') if (type(time) == pd.tslib.Timestamp) else time)
	tw.index = tw['date']
	del tw['date']
	tw.to_csv('this_week_result.csv')
	return tw

if __name__ == '__main__':
	# tw_df = pay_this_week(df,4,15,4,22)
	# tw_pv = paymt_pivot(tw_df)
	tw = output(df,4,15,4,22)
	
	print(tw)
	
	