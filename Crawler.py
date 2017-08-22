# -*- coding:utf-8 -*-

import sys
import re
import xlrd
import xlwt



f = file('text.txt')
data = f.read()
f.close()


raw_list = re.findall(r'.+：统\d+，武\d+，智\d+，政\d+，魅\d+',data)
result = []

for i in xrange(0, len(raw_list)-1):

    j = raw_list[i].split('，')
    j.append(j[0].split('：'))

    k = []
    k.append(j[-1][0])
    k.append(j[-1][1][3:])

    for m in xrange(1, 5):
        k.append(j[m][3:])

    result.append(k)


'''
book = xlrd.open_workbook(r'sheet.xlsx')
table = book.sheet_by_index(0)


for i in xrange(0, len(result)-1):
    for j in xrange(0, len(result[i])-1):
        .write(i, j, result[i][j], xlwt.set_style('Times New Roman',220, True))

print table.xlrd.name, table.xlrd.nrows, table.xlrd.ncols
'''

book = xlwt.Workbook()
sheet1 = book.add_sheet(u'sheet',cell_overwrite_ok=True)


for i in xrange(0, len(result)-1):
    for j in xrange(0, len(result[i])-1):
        sheet1.write(i,j,result[i][j])


book.save('result')
