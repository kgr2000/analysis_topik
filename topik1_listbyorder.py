import pandas as pd
from openpyxl import Workbook
wb = Workbook()
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_excel('d:/python/study/TOPIK/data/topik1_after.xlsx', encoding='euckr')
a = pd.value_counts(df['원형'].values, sort=True)
ws3 = wb.create_sheet(title = "ListByOrder")

index_list = []
values_list = []

for i in a.index :
    index_list.append(i)
for v in a.values :
    values_list.append(v)
ddf = pd.DataFrame(values_list, index_list)
for row in dataframe_to_rows(ddf, index=True, header=True):
    if len(row) > 1:
        ws3.append(row)
wb.save("D:/python/study/topik/excel_test.xlsx") #저장

#    전체 데이터 프레임을 for문으로 돌려서 키 값과 일치하면 새로운 프레임에 태소를 추가
