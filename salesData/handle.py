import pandas as pd

data = pd.read_excel('kl_cleanning.xls')
df = pd.DataFrame(data)
# test = df.ix['price']
# print(test)
v = df
v['category'] = 'cleanning'
df.index = df.SKUID
writer = pd.ExcelWriter('handle.xlsx')
v.to_excel(writer)
writer.save()
print(v)
print(df.columns)
print(df.index)
