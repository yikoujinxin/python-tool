import pandas as pd
import re
pd.set_option('display.width', None)
f = pd.ExcelFile('C:\history\python\結果対照表.xlsx')
result = {}

for name in f.sheet_names[:10]:
    check_list = []
    df = f.parse(sheet_name = name)
    df_list = list(df)
    for index, row in df[4:].iterrows():  # 根据数据模板，从第六行开始加载数据。
        if(str(row[1]).isdigit()):        # 根据数据模板，项目行名为:全半角数字
            for col_index in range(len(df_list)):          # 遍历获取每行中每一列的单元格值
                if re.match(r'TOA', str(df.loc[index, df_list[col_index]])) and df.loc[index, df_list[col_index+2]] in ('L','H','再'):
                    print(df.loc[index, df_list[col_index]])
                    print(df.loc[index, df_list[col_index+2]])
                    check_list.append({df.loc[index, df_list[col_index]]:df.loc[index, df_list[col_index+2]]})
        else:
            continue
    result[name] = check_list
print("result",result)
