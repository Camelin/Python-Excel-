import re
import time
from datetime import datetime
import xlrd
import os, sys

from xlrd import xldate_as_tuple

data = xlrd.open_workbook(r'F:\\h合全.xlsx')
sheet = data.sheet_by_index(0)
nrow = sheet.nrows

for i in range(1, nrow):
    name = sheet.cell(i, 0).value
    date = sheet.cell(i, 2).value
    ctype = sheet.cell(i,2).ctype
    if ctype==3:
        date = datetime(*xldate_as_tuple(date, 0))
        date = date.strftime('%Y-%d-%m')
    year=date.split("-")[0]
    month = date.split("-")[2]
    filename=name+".txt"
    content = sheet.cell(i, 3).value
    basedir='F:\\采集文件\\16\\'+year+"\\"+month+"\\"
    try:
        f=open(basedir+ filename, 'w',encoding='utf-8')
    except:
        os.makedirs(basedir)
        f = open(basedir + filename, 'w', encoding='utf-8')
    f.write(content)
    f.close()
