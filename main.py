# -*- coding: utf-8 -*-
import pandas as pd
import re

filename = 'LISTA JULHO CAMPANHA MENSAL.xlsx'

in_file = "in/" +  filename
out_file = "out/" +  filename


df = pd.read_excel(in_file,header=1)
df.index.names = ['id']
df.rename_axis('id')
print(df.index)
new_df = pd.DataFrame(df['APLICABILIDADE'].str.split(',').tolist(), index=df.index).stack()
new_df = new_df.reset_index([0, 'id'])
new_df.columns = ['id', 'APLICABILIDADE']
final_df = pd.merge(df, new_df,left_index=True, right_on='id')
grouped_df = final_df.groupby(by='APLICABILIDADE_y')

writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
for key, item in grouped_df:
    # []:*?/\
    new_sheet = re.sub('[\[\]:*?///]',' ', key)
    new_sheet = re.sub('[ ]+',' ', new_sheet)
    new_sheet = (new_sheet[:28] + '...') if len(new_sheet) > 31 else new_sheet
    # print(grouped_df.get_group(key))
    grouped_df.get_group(key).to_excel(writer,sheet_name=new_sheet,encoding='utf8')

writer.save()
# pd.concat([Series(row['var2'], row['var1'].split(','))
#                     for _, row in a.iterrows()]).reset_index()
# print(grouped_df)
