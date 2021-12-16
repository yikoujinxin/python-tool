import io
import re
import sys
import pandas as pd
# import msoffcrypto
import MySQLdb
from openpyxl import load_workbook
pd.set_option('display.width', None)
host="toa-cloud-test.cjrkfow6klcg.ap-northeast-1.rds.amazonaws.com"
user="shop_auction"
password="xjM2VxxIJZGHhlImnNt2yNntYGSLZBeG"
dbname="shop_auction"

def update_summary_prices(file_path, db, cursor):
    of = pd.ExcelFile(file_path)
    odf2 = of.parse(sheet_name = "进货价格缺失")
    odf2_list = list(odf2)
    prices_list = []
    for index, row in odf2[0:].iterrows():
        if not pd.isnull(odf2.loc[index, odf2_list[5]]):
            count_sql = "select COUNT(1) from dv_summary_tbl where summary_id = " + str(row[0]) + ";"
            count = cursor.execute(count_sql)
            # print(count_sql, count)
            if int(count) > 0:
                prices_list.append("\""+str(row[0])+"\"")
                # print("进价含税: ",str(row[0]),str(row[5]))
                prices_cmd = "update dv_summary_tbl set summary_price = " + str(row[5]) + " where summary_price is null and summary_id = " + str(row[0]) + ";"
                print(prices_cmd)
                cursor.execute(prices_cmd)
                db.commit()
            else:
                print("[prices_cmd] summary_id not exit: ", str(row[0]))

    print("prices_list: ",tuple(prices_list))

def update_summary_jancode(file_path, db, cursor):
    of4 = pd.ExcelFile(file_path)
    odf4 = of4.parse(sheet_name = "jancode缺失")
    odf4_list = list(odf4)
    jancode_list = []
    for index, row in odf4[1:].iterrows():
        if not pd.isnull(odf4.loc[index, odf4_list[2]]):
            count_sql = "select COUNT(1) from dv_summary_tbl where summary_id = " + str(row[0]) + ";"
            count = cursor.execute(count_sql)
            # print(count_sql, count)
            if int(count) > 0:
                jancode_list.append("\""+str(row[0])+"\"")
                jancode_cmd = "update dv_summary_tbl set commodity_jancode = " + str(int(row[2])) + " where commodity_jancode is null and summary_id = " + str(row[0]) + ";"
                print(jancode_cmd)
                cursor.execute(jancode_cmd)
                db.commit()
            else:
                print("[jancode_cmd] summary_id not exit: ", str(row[0]))
    print("jancode_list: ",tuple(jancode_list))


def cal_summary_income(db, cursor):
    cal_summary_list = []
    find_cmd = "SELECT summary_id, summary_price, summary_quantity FROM dv_summary_tbl WHERE (summary_income IS NULL OR summary_income <= 0) AND summary_quantity > 0 AND summary_price > 0;"
    cursor.execute(find_cmd)
    summary_result = cursor.fetchall()
    for (summary_id, summary_price, summary_quantity) in cursor:
        # print(summary_id, summary_price, summary_quantity,float(summary_price)*float(summary_quantity))
        cal_summary_list.append(str(summary_id))
        cal_cmd = "update dv_summary_tbl set summary_income = " + str(float(summary_price)*float(summary_quantity)) + " where summary_id = " + str(summary_id) + ";"
        print(cal_cmd)
        cursor.execute(cal_cmd)
        db.commit()
    print("cal_summary_list: ", cal_summary_list)

def cal_summary_price(db, cursor):
    price_summary_list = []
    find_cmd = "SELECT summary_id, summary_tax, summary_cost_price, summary_quantity FROM dv_summary_tbl WHERE (summary_income IS NULL OR summary_income <= 0) AND summary_quantity > 0 AND (summary_price IS NULL OR summary_price <= 0) AND summary_cost_price >0;"
    cursor.execute(find_cmd)
    summary_result = cursor.fetchall()
    for (summary_id, summary_tax, summary_cost_price, summary_quantity) in cursor:
        price_summary_list.append(str(summary_id))
        summary_price = float(summary_cost_price) + float(summary_tax)*float(summary_cost_price)
        cal_price_cmd = "update dv_summary_tbl set summary_price = " + str(summary_price) + " ,summary_income = " + str(float(summary_price)*float(summary_quantity)) + " where summary_id = " + str(summary_id) + ";"
        print(cal_price_cmd)
        cursor.execute(cal_price_cmd)
        db.commit()
    print("price_summary_list: ", price_summary_list)

if __name__ == '__main__':
    file_path = sys.argv[1]
    db=MySQLdb.connect(host,user,password,dbname,charset="utf8")
    cursor=db.cursor()
    # update_summary_prices(file_path, db, cursor)
    # print("------------------------------------------------")
    # update_summary_jancode(file_path, db, cursor)
    print("=================================================")
    cal_summary_income(db, cursor)
    print("#################################################")
    # cal_summary_price(db, cursor)

