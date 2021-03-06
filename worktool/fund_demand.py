# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os
import sys

try:
	data = pd.read_excel('data.xlsx',encoding='utf8')
	df = pd.DataFrame(data)
	df.columns = df.columns.map(str.strip)
except:
	print('Error: 请提供\'data.xlsx\'文件')
	raise SystemExit

# locale.setlocale(locale.LC_ALL,'')
exchange = {
	'HKD':'0.87',
	'USD':'6.77',
	'EUR':'8',
	'GBP':'8.88',
	'JPY':'0.062'
}
#选择汇率
def get_Exchange(current):
	return float(exchange[current])
#格式化货币
def RMB_format(value):
	return float('%.2f'%value)
#调整列位置
def change_loc(df,col_to_set,position):
	df_col = df[col_to_set]
	df.drop(labels=[col_to_set],axis=1,inplace=True)	
	df.insert(position,col_to_set,df_col)

#汇总本周需要付款的订单详情
def pay_this_week(data,time_start,time_end):
	
	df = data[(data['出纳付款时间'].isnull()) & (data['要求付款时间'] >= pd.Timestamp(time_start)) & (data['要求付款时间'] <= pd.Timestamp(time_end))]
	
	
	del df['分组：步骤号']
	# df['要求付款时间'] = pd.to_datetime(df['要求付款时间'],format='%Y%m')
	df.index = df['流水号']


	df['RMB'] = df['币种'].map(get_Exchange) * df['要求付款金额']
	# df['RMB'] = df['RMB'].map(RMB_format)
	# print(df)
	return df
	

#对数据进行透视
def paymt_pivot(df,time_start,time_end):

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

	#付款账户
	pay_account = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],columns=['付款公司'],aggfunc=[np.sum],fill_value=0)['sum']['RMB']
	pay_account = pd.DataFrame(pay_account)
	# pay_account.columns = ['精盈实业']
	pay_account['date'] = pay_account.index

	#初始化表格，判断是否有赎货
	if '中电赎货' not in df['资金来源'].values:
		df2 = pd.DataFrame((np.arange(16).reshape(1,16)),columns=df.columns)
		df2['资金来源'] = '中电赎货'
		df2['要求付款金额'] = 0
		df2['要求付款时间'] = pd.datetime(2017,1,1)
		df = df.append(df2,ignore_index=True)
	
	#获取ZD赎货数据
	ZD_redeem = pd.pivot_table(df,index=['要求付款时间'],values=['RMB'],columns=['资金来源'],aggfunc=[np.sum],fill_value=0)
	pv_isRedeem = ZD_redeem['sum']['RMB']['中电赎货'] 
	df_pv_redeem = pd.DataFrame(pv_isRedeem)
	df_pv_redeem.columns = ['中电赎货']
	df_pv_redeem['date'] = df_pv_redeem.index

	
	#合并处理的数据并以日期列作为索引
	pv = pd.merge(pd.merge(pd.merge(pd.merge(pv_all,df_pv_redeem),df_pv_isZD),df_pv_isCredit),pay_account,how='left',on='date')
	pv.set_index('date',inplace=True,drop=True)

	pv['应付供应商资金'] = pv['应付总额'] - pv['中电赎货']
	pv['自有资金'] = pv['应付供应商资金'] - pv['账期资金'] - pv['中电资金']


	pv.ix['总计'] = {'应付总额':pv['应付总额'].sum(),'中电赎货':pv['中电赎货'].sum(),'中电资金':pv['中电资金'].sum(),'账期资金':pv['账期资金'].sum(),'应付供应商资金':pv['应付供应商资金'].sum(),'自有资金':pv['自有资金'].sum(),'精盈实业':pv['精盈实业'].sum(),'跨境国际':pv['跨境国际'].sum()}
	pv.index = pd.Series(pv.index.tolist()).apply(lambda time:pd.Period(time,freq='D') if (type(time) == pd.tslib.Timestamp) else time)
	pv.index.name = '日期'

	#调整列位置
	change_loc(pv, '应付供应商资金', 2)
	change_loc(pv, '自有资金', 5)

	#统计输出

	credit_count = len(df[df['账期'] == '是'])
	zd_count = len(df[df['资金来源'] == '中电资金'])
	redeem_count = len(df[df['资金来源'] == '中电赎货'])
	print('---------------------------')
	print('以下为%s年%s月第 周（%s.%s-%s.%s）已提出计划向的采购付款需求详情。\n'%(time_start[:4],time_start[4:6],time_start[4:6],time_start[-2:],time_end[4:6],time_end[-2:]))
	print('本周总付款订单共%d笔，总额%.2f元（含中电赎货），不排除临时增加采购的需求，实际的资金需求会根据具体的货物情况有变动。'%(len(df),pv.ix['总计']['应付总额']))
	print('本周中电赎货共%d笔，总额%.2f元'%(redeem_count,pv.ix['总计']['中电赎货']))
	print('本周账期结算共%d笔，总额%.2f元，付款优先级最高。'%(credit_count,pv.ix['总计']['账期资金']))
	print('本周已使用中电资金共%d笔，总额%.2f元，优先进行付款。'%(zd_count,pv.ix['总计']['中电资金']))
	print('')
	print('以下为每日的付款需求，请财务部根据账户状况进行付款方案的安排，并通知采购部。')
	print('---------------------------')

	return pv

def output(df,time_start,time_end):
	ptw = pay_this_week(df,time_start,time_end)
	tw = paymt_pivot(ptw,time_start,time_end)
	#格式化数据
	for col in tw.columns:
		tw[col] = tw[col].apply(RMB_format)
	# tw.to_csv('this_week_result.csv')
	tw.to_excel('this_week_result.xlsx')
	# print(tw)
	return tw

if __name__ == '__main__':
	# time_start = input('请输入开始日期(如\'20170101\'):')
	# time_end = input('请输入结束日期(如\'20170101\'):')
	# tw = output(df,time_start,time_end)
	tw = output(df,'20170814','20170818')
	# print(tw.ix['总计']['应付总额'])
	


	