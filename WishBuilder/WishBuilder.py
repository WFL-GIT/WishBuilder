from datetime import datetime
from xlrd import xldate_as_tuple
import requests
import xlrd,os



WishInterface = "https://www.baidu.com"

# try:
#     data = xlrd.open_workbook('/Users/wangfulong/PycharmProjects/WishBuilder/upload/wish.xlsx')
#     table = data.sheets()[0]
#     tables=[ ]
# except:
#     print("当前不存在待读取文件Wish.xlsx 请先上传")

def import_excel(table,tables):
    i = 1
    for rown in range(table.nrows):
        array = {'行号':'','staff_num': '', 'wish_val': '','note':'','Status_code':''}
        array['行号'] = i
        array['staff_num'] = int(table.cell_value(rown, 0))
        array['wish_val'] = int(table.cell_value(rown, 1))
        print(i,table.cell_value(rown,2),type(table.cell_value(rown,2)))
        if table.cell_value(rown, 2) == '':
            array['note'] = '/'
        else:
            array['note'] = table.cell_value(rown, 2)
        wishresponse = requests.post(url=WishInterface,data=array)
        array['Status_code'] = wishresponse.status_code
        tables.append(array)
        i = i + 1
    return tables




