import io
import re
import sys
import pandas as pd
import MySQLdb
from openpyxl import load_workbook
pd.set_option('display.width', None)

host="toa-cloud-test.cjrkfow6klcg.ap-northeast-1.rds.amazonaws.com"
user="shop_auction"
password="xjM2VxxIJZGHhlImnNt2yNntYGSLZBeG"
dbname="shop_auction"

def insert_smaregi1_prices(file_path, dir_path, file_path2, db, cursor):
    insert_smaregi_list = []
    insert_zero_smaregi_list = []
    insert_zero_smaregi_list.append("商品名"+","+"JANコード"+","+"価格"+","+"種類")
    of = pd.ExcelFile(file_path)
    of2 = pd.ExcelFile(file_path2)

    preparation_data = pd.read_excel(file_path2, sheet_name = "preparation", usecols=[0])
    preparation_data_list = preparation_data.values.tolist()
    preparation_data_array = []
    for i in range(len(preparation_data_list)):
        preparation_data_array.append(str(preparation_data_list[i][0]))
    print("preparation_data_array: ",preparation_data_array)
    odf_preparation = of2.parse(sheet_name = "preparation")
    odf_preparation_list = list(odf_preparation)

    priority_data = pd.read_excel(file_path2, sheet_name = "priority", usecols=[0])
    priority_data_list = priority_data.values.tolist()
    priority_data_array = []
    for i in range(len(priority_data_list)):
        priority_data_array.append(str(priority_data_list[i][0]))
    print("priority_data_array: ",priority_data_array)
    odf_priority = of2.parse(sheet_name = "priority")
    odf_priority_list = list(odf_priority)

    for index_priority, row_priority in odf_priority[0:].iterrows():
        if(row_priority[3]) != 0 and not pd.isnull(odf_priority.loc[index_priority, odf_priority_list[3]]) :
            insert_priority_cmd = "insert into dv_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            insert_priority_cmd = insert_priority_cmd + str(odf_priority.loc[index_priority, odf_priority_list[1]]).replace("\"", "")+"\",\""+str(odf_priority.loc[index_priority, odf_priority_list[0]])+"\",\""+str(odf_priority.loc[index_priority, odf_priority_list[3]])+"\",\"1\");"
            print("原価优先商品: ", insert_priority_cmd)
            insert_smaregi_list.append(insert_priority_cmd+"\n")

    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for index, row in odf[0:].iterrows():
            insert_cmd = "insert into dv_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
            insert_cmd = insert_cmd + str(odf.loc[index, odf_list[4]]).replace("\"", "")+"\",\""+str(odf.loc[index, odf_list[3]])+"\",\""+str(odf.loc[index, odf_list[6]])+"\",\"1\");"
            if(row[6]) != 0 and not pd.isnull(odf.loc[index, odf_list[6]]) and not pd.isnull(odf.loc[index, odf_list[3]]) and str(odf.loc[index, odf_list[3]]) not in priority_data_array:
                insert_smaregi_list.append(insert_cmd+"\n")
            elif str(odf.loc[index, odf_list[3]]) in preparation_data_array:
                for index_preparation, row_preparation in odf_preparation[0:].iterrows():
                    if(row_preparation[3]) != 0 and not pd.isnull(odf_preparation.loc[index_preparation, odf_preparation_list[3]]) and str(odf_preparation.loc[index_preparation, odf_preparation_list[0]]) == str(odf.loc[index, odf_list[3]]):
                        insert_preparation_cmd = "insert into dv_smaregi_tbl(commodity_name,commodity_jancode,smaregi_price,remark) values(\""
                        insert_preparation_cmd = insert_preparation_cmd + str(odf_preparation.loc[index_preparation, odf_preparation_list[1]]).replace("\"", "")+"\",\""+str(odf_preparation.loc[index_preparation, odf_preparation_list[0]])+"\",\""+str(odf_preparation.loc[index_preparation, odf_preparation_list[3]])+"\",\"1\");"
                        print("原価異常商品: ", insert_preparation_cmd)
                        insert_smaregi_list.append(insert_preparation_cmd+"\n")
            else:
                insert_zero_smaregi_list.append(str(odf.loc[index, odf_list[4]])+","+str(odf.loc[index, odf_list[3]])+","+str(odf.loc[index, odf_list[6]])+",1")

    with open(dir_path + "/" + "insert_smaregi1.sql", mode="w",encoding='cp932',errors="ignore") as f:
        f.writelines(insert_smaregi_list)   
    for cmd in insert_smaregi_list:     
        cursor.execute(cmd)
        db.commit()
    
    with open(dir_path + "/" + "insert_zero_smaregi1.csv", mode="w",encoding='cp932',errors="ignore") as f:
        for ele in insert_zero_smaregi_list:
            f.write(ele + '\n')

if __name__ == '__main__':
    #python insertSmaregi.py C:\pdf smaregi1.xlsx smaregi2.xlsx smaregi3.xlsx smaregi4.xlsx
    dir_path = input("输入文件读取地址: ")
    filePath1 = input("输入Smaregi文件名: ")
    filePath2 = input("输入处理文件名: ")
    file_path_smaregi = dir_path + "/" + filePath1
    file_path_exception = dir_path + "/" + filePath2
    print("------ 导入smaregi商品数据中 ------- ")
    db=MySQLdb.connect(host,user,password,dbname,charset="utf8")
    cursor=db.cursor()
    insert_smaregi1_prices(file_path_smaregi, dir_path, file_path_exception, db, cursor)
    print("------ smaregi导出商品数据 ------- 完了")
    

