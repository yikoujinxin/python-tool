import io
import os
import re
import sys
import zipfile
import pandas as pd
# import msoffcrypto
# import json
from openpyxl import load_workbook
pd.set_option('display.width', None)

def select_store_names(file_path):
    f = open(file_path, "r",encoding='utf-8_sig')
    content = f.read()
    for x in content.split('</li>'):
        matchObj = re.match(r'.*TOAmart(.*)店.*', x)
        if matchObj:
            shop_name = "C://pdf//test//" + matchObj.group(1)+"店"
            print(shop_name)
            folder = os.path.exists(shop_name)
            if not folder:
                os.makedirs(shop_name)

def file2Zip(zip_file_name:str, file_names:list):
    with zipfile.ZipFile(zip_file_name, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for fn in file_names:
            parent_path, name = os.path.split(fn)
            zf.write(fn, arcname=name)


def rename_store_names(file_path):
    file_list = os.listdir(file_path+"\\")
    prefix_name = os.path.dirname(file_path+"\\").split("\\")[-1]+"_"
    zip_name = "C:\\pdf\\test\\"+os.path.dirname(file_path+"\\").split("\\")[-1]+".zip"
    n=0
    files = []
    for i in file_list:
        old_name=file_path+os.sep+file_list[n]
        new_name=file_path+os.sep+prefix_name+file_list[n]
        os.rename(old_name, new_name)
        print(old_name, new_name)
        files.append(new_name)
        n+=1
    print("zipName: ", zip_name)
    file2Zip(zip_name, files)

if __name__ == '__main__':
    file_path = sys.argv[1]
    # output_write_result = find_check_result()
    # write_excel(output_write_result)
    # read_good_names()
    # select_store_names(file_path)
    rename_store_names(file_path)

