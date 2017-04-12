import pandas as pd
import numpy as np
from pandas import DataFrame
import os


def group_to_count(filename,group_by,column_to_sum,column_to_count,column_to_mean):
	data = pd.read_excel(filename)
	df = pd.DataFrame(data)
	df = df.fillna(value=0)
	group1 = df.groupby(group_by)[column_to_sum].sum()
	group2 = df.groupby(group_by)[column_to_count].count()
	group3 = df.groupby(group_by)[column_to_mean].mean()
	group = pd.concat([group1, group2,group3],axis=1).sort_values(by=column_to_sum,ascending=False)

	group.columns = ['sum of '+column_to_sum,'count of '+column_to_count,'mean of '+column_to_mean]
	return group


def output(filename,group_by1,group_by2,column_to_sum,column_to_count,column_to_mean):
	path = os.getcwd() + '\\analyze\\'
	dirIsExists = os.path.exists(path)
	if not dirIsExists:
		os.makedirs(path)
	else:
		pass

	writer = pd.ExcelWriter(path+filename.split('.')[0]+'_group.xlsx')
	group1 = group_to_count(filename, group_by1, column_to_sum,column_to_count,column_to_mean)
	group2 = group_to_count(filename, group_by2, column_to_sum,column_to_count,column_to_mean)
	
	group1.to_excel(writer,group_by1)
	group2.to_excel(writer,group_by2)
	writer.save()

if __name__ == '__main__':
	
	output('kl_cleanning.xls', 'brand','origin_of_brand','cmt_num','SKUID','price')
	output('kl_sunblock.xls', 'brand','origin_of_brand','cmt_num','SKUID','price')

