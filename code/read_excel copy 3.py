import io
import re
import pandas as pd
import msoffcrypto

pd.set_option('display.width', None)

def find_check_result():
    f = pd.ExcelFile('C:\history\python\結果対照表.xlsx')
    result = {}
    output_result = {}
    check_result = read_check_result()
    for name in f.sheet_names:
        check_list = []
        df = f.parse(sheet_name = name)
        df_list = list(df)
        for index, row in df[4:].iterrows():  # 根据数据模板，从第六行开始加载数据。
            if(str(row[1]).isdigit()):        # 根据数据模板，项目行名为:全半角数字
                for col_index in range(len(df_list)):          # 遍历获取每行中每一列的单元格值
                    if re.match(r'TOA', str(df.loc[index, df_list[col_index]])) and df.loc[index, df_list[col_index+2]] in ('L','H','再'):
                        check_list.append({df.loc[index, df_list[col_index]]:df.loc[index, df_list[col_index+2]]})
                        for check_item in check_result:
                            if check_item == str(df.loc[index, df_list[col_index]]):
                                output_result[check_item] = df.loc[index, df_list[col_index+2]]
                                output_result['sheet_name'] = name
            else:
                continue
        result[name] = check_list
    # print("result",result)
    print("output_result",output_result)

def read_check_result():
    check_result_list = []
    file_path = 'C:\history\python\東京PCR検査記録表.xlsx'
    decrypted = io.BytesIO()

    with open(file_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password="pcr-ht-n1")  # Use password
        file.decrypt(decrypted)

    of = pd.ExcelFile(decrypted)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for index, row in odf[4:].iterrows():
            if(str(row[1]).isdigit()) and not pd.isnull(odf.loc[index, odf_list[4]]):
                check_result_list.append(odf.loc[index, odf_list[4]])
    return check_result_list

# def find_check_result(reservation_code):
find_check_result()

