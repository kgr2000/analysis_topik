import pandas as pd
from openpyxl import Workbook
wb = Workbook()
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_excel('d:/python/study/TOPIK/data/topik1_after_test.xlsx', encoding='euckr')
#a = pd.value_counts(df['원형'].values, sort=True)
#for v in a.index :
#    전체 데이터 프레임을 for문으로 돌려서 키 값과 일치하면 새로운 프레임에 태소를 추가
import pandas as pd
from openpyxl import Workbook
wb = Workbook()
from openpyxl.utils.dataframe import dataframe_to_rows
index_list = []
values_list = []
df = pd.read_excel('d:/python/study/TOPIK/data/topik1_after_test.xlsx', encoding='euckr')
a = pd.value_counts(df['태소'].values, sort=True)
#print(type(int(a['힘든'])))
#######################
inf_dict = {}
for i in range(len(df)) :
#for i in range(100) :
    inf = df.iloc[i].원형
    pum = df.iloc[i].품사
    tae = df.iloc[i].태소
    mun = df.iloc[i].문장
    if inf not in inf_dict.keys() :
        inf_dict[inf] = [pum, tae]
    if inf in inf_dict.keys() :
        if tae in inf_dict[inf] :
            continue
        elif len(inf_dict[inf]) < 6 :
#            if a[tae]
            inf_dict[inf].append(tae)
        elif len(inf_dict[inf]) > 5 :
            for l in range(1,5) :
                if a[tae] > a[inf_dict[inf][1:]].values.any() :
                    if a[inf_dict[inf][l]] < a[tae] :
                        inf_dict[inf][l] = tae
#                        inf_dict[inf][l].replace(inf_dict[inf][l], tae)
                        break




print(inf_dict)
#if inf in inf_dict :
#    for l in range(1,5) :
#        if a[tae] > a[inf_dict[inf][1:]].values.any() :
#            if a[inf_dict[inf][l]] < a[tae] :
#                print('yo')
