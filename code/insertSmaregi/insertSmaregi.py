import io
import re
import sys
import pandas as pd
import MySQLdb
# import msoffcrypto
# import json
from openpyxl import load_workbook
pd.set_option('display.width', None)

host="toa-cloud-test.cjrkfow6klcg.ap-northeast-1.rds.amazonaws.com"
user="shop_auction"
password="xjM2VxxIJZGHhlImnNt2yNntYGSLZBeG"
dbname="shop_auction"

def insert_smaregi1_prices(file_path, dir_path, db, cursor):
    insert_smaregi_list = []
    insert_zero_smaregi_list = []
    insert_zero_smaregi_list.append("商品名"+","+"JANコード"+","+"価格"+","+"種類")
    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for index, row in odf[0:].iterrows():
            insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            insert_cmd = insert_cmd + str(odf.loc[index, odf_list[4]])+"\",\""+str(odf.loc[index, odf_list[3]])+"\",\""+str(odf.loc[index, odf_list[6]])+"\",\"1\");"
            if(row[6]) != 0 and not pd.isnull(odf.loc[index, odf_list[6]]) and not pd.isnull(odf.loc[index, odf_list[3]]):
                insert_smaregi_list.append(insert_cmd+"\n")
            else:
                insert_zero_smaregi_list.append(str(odf.loc[index, odf_list[4]])+","+str(odf.loc[index, odf_list[3]])+","+str(odf.loc[index, odf_list[6]])+",1")
    
    with open(dir_path + "/" + "insert_zero_smaregi1.csv", mode="w",encoding='cp932',errors="ignore") as f:
        for ele in insert_zero_smaregi_list:
            f.write(ele + '\n')

    with open(dir_path + "/" + "insert_smaregi1.csv", mode="w",encoding='cp932',errors="ignore") as f:
        f.writelines(insert_smaregi_list)

def insert_priority_smaregi2_prices(file_path, dir_path, db, cursor):
    of = pd.ExcelFile(file_path)
    odf = of.parse(sheet_name = "振分表②")
    odf_list = list(odf)
    priority_smaregi2 = []
    priority_zero_smaregi2 = []
    priority_zero_smaregi2.append("商品名"+","+"JANコード"+","+"価格"+","+"種類")
    for index, row in odf[0:1462].iterrows():
        insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
        insert_cmd = insert_cmd + str(odf.loc[index, odf_list[6]])+"\",\""+str(odf.loc[index, odf_list[8]])+"\",\""+str(odf.loc[index, odf_list[14]])+"\",\"2\");"
        if(row[13]) != 0 and not pd.isnull(odf.loc[index, odf_list[13]]) and not pd.isnull(odf.loc[index, odf_list[8]]):
            priority_smaregi2.append(insert_cmd+"\n")
        else:
            priority_zero_smaregi2.append(str(odf.loc[index, odf_list[6]])+","+str(odf.loc[index, odf_list[8]])+","+str(odf.loc[index, odf_list[14]])+","+"2")
    
    with open(dir_path + "/" + "insert_zero_smaregi2.csv", mode="w",encoding='cp932',errors="ignore") as f:
        for ele in priority_zero_smaregi2:
            f.write(ele + '\n')

    with open(dir_path + "/" + "insert_smaregi2.csv", mode="w",encoding='cp932',errors="ignore") as f:
        f.writelines(priority_smaregi2)

def insert_toa_smaregi3(file_path, dir_path, db, cursor):
    of = pd.ExcelFile(file_path)
    odf = of.parse(sheet_name = "マスター")
    odf_list = list(odf)
    priority_smaregi3 = []
    priority_zero_smaregi3 = []
    priority_zero_smaregi3.append("商品名"+","+"JANコード"+","+"価格"+","+"種類")
    for index, row in odf[1:].iterrows():
        insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
        insert_cmd = insert_cmd + str(odf.loc[index, odf_list[1]])+"\",\""+str(odf.loc[index, odf_list[2]])+"\",\""+str(odf.loc[index, odf_list[3]])+"\",\"3\");"
        if(row[3]) != 0 and not pd.isnull(odf.loc[index, odf_list[3]]) and not pd.isnull(odf.loc[index, odf_list[2]]):
            priority_smaregi3.append(insert_cmd+"\n")
        else:
            priority_zero_smaregi3.append(str(odf.loc[index, odf_list[1]])+","+str(odf.loc[index, odf_list[2]])+","+str(odf.loc[index, odf_list[3]])+","+"3")
    
    with open(dir_path + "/" + "insert_zero_smaregi3.csv", mode="w",encoding='cp932',errors="ignore") as f:
        for ele in priority_zero_smaregi3:
            f.write(ele + '\n')

    with open(dir_path + "/" + "insert_smaregi3.csv", mode="w",encoding='cp932',errors="ignore") as f:
        f.writelines(priority_smaregi3)

def insert_toa_smaregi4(file_path, dir_path, db, cursor):
    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        priority_smaregi4 = []
        priority_zero_smaregi4 = []
        priority_zero_smaregi4.append("商品名"+","+"JANコード"+","+"価格"+","+"種類")
        for index, row in odf[1:].iterrows():
            insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            insert_cmd = insert_cmd + str(odf.loc[index, odf_list[2]])+"\",\""+str(odf.loc[index, odf_list[7]])+"\",\""+str(odf.loc[index, odf_list[8]])+"\",\"4\");"
            if(row[8]) != 0 and not pd.isnull(odf.loc[index, odf_list[8]]) and not pd.isnull(odf.loc[index, odf_list[7]]):
                priority_smaregi4.append(insert_cmd+"\n")
            else:
                priority_zero_smaregi4.append(str(odf.loc[index, odf_list[2]])+","+str(odf.loc[index, odf_list[7]])+","+str(odf.loc[index, odf_list[8]])+","+"4")
    
    with open(dir_path + "/" + "insert_zero_smaregi4.csv", mode="w",encoding='cp932',errors="ignore") as f:
        for ele in priority_zero_smaregi4:
            f.write(ele + '\n')

    with open(dir_path + "/" + "insert_smaregi4.csv", mode="w",encoding='cp932',errors="ignore") as f:
        f.writelines(priority_smaregi4)

if __name__ == '__main__':
    #python insertSmaregi.py C:\pdf smaregi1.xlsx smaregi2.xlsx smaregi3.xlsx smaregi4.xlsx
    dir_path = sys.argv[1]
    filePath1 = dir_path + "/" + sys.argv[2]
    filePath2 = dir_path + "/" + sys.argv[3]
    filePath3 = dir_path + "/" + sys.argv[4]
    filePath4 = dir_path + "/" + sys.argv[5]

    db=MySQLdb.connect(host,user,password,dbname,charset="utf8")
    cursor=db.cursor()
    insert_smaregi1_prices(filePath1, dir_path, db, cursor)
    print("------ smaregi导出商品数据 ------- 完了")
    insert_priority_smaregi2_prices(filePath2, dir_path, db, cursor)
    print("------ 森井仕入れ商品数据 ------- 完了")
    insert_toa_smaregi3(filePath3, dir_path, db, cursor)
    print("------ 东亚商业制品的商品数据 ------- 完了")
    insert_toa_smaregi4(filePath4, dir_path, db, cursor)
    print("------ 东亚开发部的商品数据 ------- 完了")
    

