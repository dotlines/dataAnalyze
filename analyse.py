import pandas as pd
import numpy as np
from pandas import DataFrame
import os

# result = df.loc[:,['SKUID','price','cmt_num']]
# print(result)

# price_over_199 = df[df.price > 199]
# cmt_over_p199 = price_over_199.cmt_num.sum()
# print(cmt_over_p199)
# print(df[df['price']>200]['price'].value_counts())

#group操作
# group1 = df.groupby('origin_of_brand')['price'].mean()
# group1.to_excel('groupby_country.xls',sheet_name='sheet1')
# print(group1)
# data=DataFrame([{"id":0,"name":'lxh',"age":20,"cp":'lm'},{"id":1,"name":'xiao',"age":40,"cp":'ly'},{"id":2,"name":'hua',"age":4,"cp":'yry'},{"id":3,"name":'be',"age":70,"cp":'old'}])
# data1=DataFrame([{"id":100,"name":'lxh','cs':10},{"id":101,"name":'xiao','cs':40},{"id":102,"name":'hua2','cs':50}])
# data2=DataFrame([{"id":0,"name":'lxh','cs':10},{"id":101,"name":'xiao','cs':40},{"id":102,"name":'hua2','cs':50}])

def group_to_count(filename,group_by,column_to_sum,column_to_count,column_to_mean):
	data = pd.read_excel(filename)
	df = pd.DataFrame(data)
	df = df.fillna(value=0)
	group1 = df.groupby(group_by)[column_to_sum].sum()
	group2 = df.groupby(group_by)[column_to_count].count()
	group3 = df.groupby(group_by)[column_to_mean].mean()
	group = pd.concat([group1, group2,group3],axis=1).sort_values(by=column_to_sum,ascending=False)
	# print(group)
	# print('\n-------------------\n')
	# print(group1.index)
	# print('\n-------------------\n')
	# print(group2.index)
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
	# print(data,'\n-----------\n',data1,'\n-----------\n',data2)
	group1.to_excel(writer,group_by1)
	group2.to_excel(writer,group_by2)
	writer.save()

if __name__ == '__main__':
	# group_by = input("which column you want to group: ")
	# apply_column = input("which column you want to compute: ")
	# group_to_count(df, group_by, apply_column)
	output('kl_cleanning.xls', 'brand','origin_of_brand','cmt_num','SKUID','price')
	output('kl_sunblock.xls', 'brand','origin_of_brand','cmt_num','SKUID','price')

