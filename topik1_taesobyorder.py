import pandas as pd
from openpyxl import Workbook
wb = Workbook()
from openpyxl.utils.dataframe import dataframe_to_rows
index_list = []
values_list = []
df = pd.read_excel('d:/python/study/TOPIK/data/topik1_after_test.xlsx', encoding='euckr')
a = pd.value_counts(df['태소'].values, sort=True)
print(type(a))
for i in a.index :
    index_list.append(i)
for v in a.values :
    values_list.append(v)
ddf = pd.DataFrame(values_list, index_list)
#print(ddf[1:2])

#for row in dataframe_to_rows(ddf, index=True, header=True):
#    if len(row) > 1:
#        ws3.append(row)
#wb.save("D:/python/study/topik/excel_test.xlsx") #저장
