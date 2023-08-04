from xlrd import open_workbook
import re
from xlutils.copy import copy
import datetime
import xlwt


def readExcel(filename, n):
    excel = open_workbook(filename)
    sheet = excel.sheets()[0]
    col = sheet.col_values(n)  # n=12:waring content ; n=44:state of kaidan
    return col


if __name__ == '__main__':
    col_content = readExcel("C:\\Users\\yangzh\\Desktop\\up\\20190826-python\\告警统计0822-0823.xlsx", 12)
    col_event = readExcel("C:\\Users\\yangzh\\Desktop\\up\\20190826-python\\告警统计0822-0823.xlsx", 44)
    # writeExcelRule(coll)
    col_truecontent = col_content[1:]
    col_trueevent = col_event[1:]
    print(col_trueevent)