import re
import pandas as pd
from openpyxl import load_workbook
pd.set_option('display.width', None)


def read_check_result(file_path):
    check_result_list = []
    temp = []

    of = pd.ExcelFile(file_path)
    for name in of.sheet_names:
        odf = of.parse(sheet_name = name)
        odf_list = list(odf)
        for i in range(2, 34, 4):
            for index, row in odf[4:].iterrows():
                if re.match(r'TPN', str(odf.loc[index, odf_list[i]])) and not pd.isnull(odf.loc[index, odf_list[i]]):
                    temp.append(odf.loc[index, odf_list[i]])
        check_result_list.append([name, len(temp)])
        temp = []
                
    return check_result_list

if __name__ == '__main__':
    result = read_check_result("C://filepath")
    for name, length in result:
        if length != 0:
            print(length, name)
