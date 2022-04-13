from urllib.request import urlopen
import xlrd
import xlwt
from bs4 import BeautifulSoup as bs
from xlutils.copy import copy
import re, datetime
import os


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")


# 暂时并没有使用这个方法
def read_excel_xls(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
        print()


if __name__ == '__main__':

    r = urlopen("https://www.boc.cn/sourcedb/whpj/")
    c = r.read()
    bs_obj = bs(c, "html")
    t = bs_obj.find_all("table")[1]
    all_tr = t.find_all("tr")
    all_tr.pop(0)  # 删除第一个元素。
    value1 = []
    rmbDate = ""
    for r in all_tr:
        all_td = r.find_all("td")
        match = re.search('\d{4}.\d{2}.\d{2}', all_td[6].text)
        date = datetime.datetime.strptime(match.group(), '%Y.%m.%d').date()
        newDate = str(date)[:-3]
        rmbDate = newDate

        if all_td[0].text == "港币":
            row = ["HKD" + all_td[0].text, float(all_td[5].text) / 100, newDate]
            value1.append(row)
        elif all_td[0].text == "美元":
            row = ["USD" + all_td[0].text, float(all_td[5].text) / 100, newDate]
            value1.append(row)
        elif all_td[0].text == "日元":
            row = ["JPY" + all_td[0].text, float(all_td[5].text) / 100, newDate]
            value1.append(row)
        elif all_td[0].text == "欧元":
            row = ["EUR" + all_td[0].text, float(all_td[5].text) / 100, newDate]
            value1.append(row)

            # print(f'{all_td[0].text}    中行折算价:{float(all_td[5].text) / 100}  日期:{newDate}  ')

    RMB = ["CNY人民币", 1, rmbDate]
    value1.append(RMB)

    book_name_xls = '各地汇率表.xls'

    sheet_name_xls = '货币汇率'

    value_title = [["货币", "中行折算价", "日期"], ]
    # 判断文件是否存在
    if os.path.exists("各地汇率表.xls"):
        write_excel_xls_append(book_name_xls, value1)
    else:
        write_excel_xls(book_name_xls, sheet_name_xls, value_title)
        write_excel_xls_append(book_name_xls, value1)
