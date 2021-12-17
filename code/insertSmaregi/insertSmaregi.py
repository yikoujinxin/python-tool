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

def insert_smaregi1_prices(file_path, db, cursor):
    insert_smaregi_list = []
    insert_zero_smaregi_list = []
    insert_zero_smaregi_list.append(["商品名","JANコード","価格","種類"])
    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for index, row in odf[0:].iterrows():
            smaregi_sql = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            smaregi_sql = smaregi_sql + str(odf.loc[index, odf_list[4]])+"\",\""+str(odf.loc[index, odf_list[3]])+"\",\""+str(odf.loc[index, odf_list[6]])+"\",\"1\");"
            if(row[6]) != 0 and not pd.isnull(odf.loc[index, odf_list[6]]) and not pd.isnull(odf.loc[index, odf_list[3]]):
                insert_smaregi_list.append(smaregi_sql)
            else:
                insert_zero_smaregi_list.append([str(odf.loc[index, odf_list[4]]),str(odf.loc[index, odf_list[3]]),str(odf.loc[index, odf_list[6]]),"1"])
    zero_smaregi1_data_frame= pd.DataFrame(insert_zero_smaregi_list)  
    zero_smaregi1_data_frame.to_csv("C:\\pdf\\sql\\insert_zero_smaregi1.csv",index=False,sep=",",encoding='cp932')

    smaregi1_data_frame= pd.DataFrame(insert_smaregi_list)  
    smaregi1_data_frame.to_csv("C:\\pdf\\sql\\insert_smaregi1.csv",index=False,sep=",",encoding='cp932')


def insert_priority_smaregi2_prices(file_path, db, cursor):
    of = pd.ExcelFile(file_path)
    odf = of.parse(sheet_name = "振分表②")
    odf_list = list(odf)
    priority_smaregi2 = []
    priority_zero_smaregi2 = []
    priority_zero_smaregi2.append(["商品名","JANコード","価格","種類"])
    for index, row in odf[0:1462].iterrows():
        insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
        insert_cmd = insert_cmd + str(odf.loc[index, odf_list[6]])+"\",\""+str(odf.loc[index, odf_list[8]])+"\",\""+str(odf.loc[index, odf_list[14]])+"\",\"2\");"
        if(row[13]) != 0 and not pd.isnull(odf.loc[index, odf_list[13]]) and not pd.isnull(odf.loc[index, odf_list[8]]):
            priority_smaregi2.append(insert_cmd)
        else:
            priority_zero_smaregi2.append([str(odf.loc[index, odf_list[6]]),str(odf.loc[index, odf_list[8]]),str(odf.loc[index, odf_list[14]]),"2"])
    zero_smaregi2_data_frame= pd.DataFrame(priority_zero_smaregi2)
    with open("C:\\pdf\\sql\\insert_zero_smaregi2.csv", mode="w",encoding='cp932',errors="ignore") as f:
        zero_smaregi2_data_frame.to_csv(f,index=False)

    smaregi2_data_frame= pd.DataFrame(priority_smaregi2)
    with open("C:\\pdf\\sql\\insert_smaregi2.csv", mode="w",encoding='cp932',errors="ignore") as f:
        smaregi2_data_frame.to_csv(f,index=False)

def insert_toa_smaregi3(file_path, db, cursor):
    of = pd.ExcelFile(file_path)
    odf = of.parse(sheet_name = "マスター")
    odf_list = list(odf)
    priority_smaregi3 = []
    priority_zero_smaregi3 = []
    priority_zero_smaregi3.append(["商品名","JANコード","価格","種類"])
    for index, row in odf[1:].iterrows():
        insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
        insert_cmd = insert_cmd + str(odf.loc[index, odf_list[1]])+"\",\""+str(odf.loc[index, odf_list[2]])+"\",\""+str(odf.loc[index, odf_list[3]])+"\",\"3\");"
        if(row[3]) != 0 and not pd.isnull(odf.loc[index, odf_list[3]]) and not pd.isnull(odf.loc[index, odf_list[2]]):
            priority_smaregi3.append(insert_cmd)
        else:
            priority_zero_smaregi3.append([str(odf.loc[index, odf_list[1]]),str(odf.loc[index, odf_list[2]]),str(odf.loc[index, odf_list[3]]),"3"])
    zero_smaregi3_data_frame= pd.DataFrame(priority_zero_smaregi3)
    with open("C:\\pdf\\sql\\insert_zero_smaregi3.csv", mode="w",encoding='cp932',errors="ignore") as f:
        zero_smaregi3_data_frame.to_csv(f,index=False)

    smaregi3_data_frame= pd.DataFrame(priority_smaregi3)
    with open("C:\\pdf\\sql\\insert_smaregi3.csv", mode="w",encoding='cp932',errors="ignore") as f:
        smaregi3_data_frame.to_csv(f,index=False)

def insert_toa_smaregi4(file_path, db, cursor):
    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        priority_smaregi4 = []
        priority_zero_smaregi4 = []
        priority_zero_smaregi4.append(["商品名","JANコード","価格","種類"])
        for index, row in odf[1:].iterrows():
            insert_cmd = "insert into dv_clone_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            insert_cmd = insert_cmd + str(odf.loc[index, odf_list[2]])+"\",\""+str(odf.loc[index, odf_list[7]])+"\",\""+str(odf.loc[index, odf_list[8]])+"\",\"4\");"
            if(row[8]) != 0 and not pd.isnull(odf.loc[index, odf_list[8]]) and not pd.isnull(odf.loc[index, odf_list[7]]):
                priority_smaregi4.append(insert_cmd)
            else:
                priority_zero_smaregi4.append([str(odf.loc[index, odf_list[2]]),str(odf.loc[index, odf_list[7]]),str(odf.loc[index, odf_list[8]]),"4"])
    zero_smaregi4_data_frame= pd.DataFrame(priority_zero_smaregi4)
    with open("C:\\pdf\\sql\\insert_zero_smaregi4.csv", mode="w",encoding='cp932',errors="ignore") as f:
        zero_smaregi4_data_frame.to_csv(f,index=False)

    smaregi4_data_frame= pd.DataFrame(priority_smaregi4)
    with open("C:\\pdf\\sql\\insert_smaregi4.csv", mode="w",encoding='cp932',errors="ignore") as f:
        smaregi4_data_frame.to_csv(f,index=False)

if __name__ == '__main__':
    # filePath1 = sys.argv[1] + "/" + sys.argv[2]
    # filePath2 = sys.argv[1] + "/" + sys.argv[3]
    # filePath3 = sys.argv[1] + "/" + sys.argv[4]
    # filePath4 = sys.argv[1] + "/" + sys.argv[5]

    db=MySQLdb.connect(host,user,password,dbname,charset="utf8")
    cursor=db.cursor()
    insert_smaregi1_prices("C:\pdf\smaregi1.xlsx", db, cursor)
    print("------1------- 完了")
    insert_priority_smaregi2_prices("C:\pdf\smaregi2.xlsx", db, cursor)
    print("------2------- 完了")
    insert_toa_smaregi3("C:\pdf\smaregi3.xlsx", db, cursor)
    print("------3------- 完了")
    insert_toa_smaregi4("C:\pdf\smaregi4.xlsx", db, cursor)
    print("------4------- 完了")
    

