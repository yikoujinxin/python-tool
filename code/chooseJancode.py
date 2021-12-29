import io
import re
import sys
import pandas as pd
# import msoffcrypto
# import json
from openpyxl import load_workbook
pd.set_option('display.width', None)

def insert_smaregi_prices(file_path):
    insert_smaregi_list = []
    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for index, row in odf[0:].iterrows():
            if(row[6]) == 0 or pd.isnull(odf.loc[index, odf_list[6]]):
                insert_smaregi_list.append(["insert into dv_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values('"+
                        str(odf.loc[index, odf_list[4]])+"','"+str(odf.loc[index, odf_list[3]])+"','"+
                        str(odf.loc[index, odf_list[5]])+"','販売価格の計算を使用して、購入価格が見つかりませんでした');"
                    ])
            else:
                insert_smaregi_list.append(["insert into dv_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values('"+
                        str(odf.loc[index, odf_list[4]])+"','"+str(odf.loc[index, odf_list[3]])+"','"+
                        str(odf.loc[index, odf_list[6]])+"','');"
                    ])
    data_frame= pd.DataFrame(insert_smaregi_list)  
    data_frame.to_csv("C:\\pdf\\发注数据\\更新数据\\insert_smaregi_sql.csv",index=False,sep=",",encoding='cp932')

if __name__ == '__main__':
    file_path = sys.argv[1]
    insert_smaregi_prices(file_path)

