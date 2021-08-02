import pandas as pd
import re
pd.set_option('display.width', None)
f = pd.ExcelFile('C:\history\python\結果対照表.xlsx')
# print(f.sheet_names)
# print(f.parse(sheet_name = '0719-10').fillna(method='ffill'))
df = f.parse(sheet_name = '0719-6(2)')
df_list = list(df)
# if df.loc[df["B"].isin(1)]:
#     for row in df[5:].iterrows():  # 数据模板，从第六行开始加载数据。
#         print('---------')
#         print(row)

for index, row in df[4:].iterrows():  # 根据数据模板，从第六行开始加载数据。
    if(str(row[1]).isdigit()):        # 根据数据模板，项目行名为:全半角数字
        for col_index in range(len(df_list)):          # 遍历获取每行中每一列的单元格值
            if re.match(r'TOA', str(df.loc[index, df_list[col_index]])):
                print(df.loc[index, df_list[col_index]])
                print(df.loc[index, df_list[col_index+2]])
            # if pd.isnull(df.loc[index, col]) or not re.match(r'TOA', df.loc[index, col]):
            #     continue
            # elif df.loc[index+1, col] in ('L','H','再'):
            #     print('---------')
            #     print(df.loc[index, col])
            #     print(df.loc[index+1, col])
    else:
        continue
#     for col in row:
#         print('---------')
#         print(col)
        # print(df.loc[col, 1])
        # if pd.isnull(df.loc[col, 1]):
    # for col in list(df):
    #     print(row[1])
        # print(list(row))
        # print('---------')
        # print(df.loc[5, col])
        # print(df.loc[row[0], col])
        
        
        # If this cell is empty all in the same row too.
        # if pd.isnull(df.loc[prow, 'Album Name']):  
        #     continue
        #  # If a cell and next one are empty, take previous valor. 
        # elif pd.isnull(df.loc[prow, col]) and pd.isnull(df.loc[row[0], col]):
        #     df.loc[prow, col] = df.loc[pprow, col]
        # pprow = prow
        # prow = row[0]
