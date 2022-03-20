
# -*- coding: UTF-8 -*-

import xlrd  #引入模块

#打开文件，获取excel文件的workbook（工作簿）对象
workbook=xlrd.open_workbook("./job.xls")  #文件路径


#获取所有sheet的名字
names=workbook.sheet_names()
print(names) #['普通职位', '行政执法类', '专业技术类']

#  #通过sheet索引获得sheet对象
#  worksheet=workbook.sheet_by_index(0)
#  print(worksheet)  #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>
#
#  #通过sheet名获得sheet对象
total=0
for sheet_name in names:
    worksheet=workbook.sheet_by_name(sheet_name)
    print(worksheet) #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>

    '''对sheet对象进行操作'''
    name=worksheet.name  #获取表的姓名
    print(name) #各省市

    nrows=worksheet.nrows  #获取该表总行数
    print(nrows)  #32

    ncols=worksheet.ncols  #获取该表总列数
    print(ncols) #13

    title=[]
    for i in range(nrows): #循环打印每一行
        row=worksheet.row_values(i)
        # print(row)
        if i==1:
            title=row
        if   '本科' in row[9] and('应届' not in row[12]) and (('不限' in row[16] and '三级' not in row[16] ) or '生物' in row[16]) :
            print('------------------------------')
            for x in range(0, 21):
                a=title[x]+"("+str(x)+"):"+str(row[x])
                print(a)
            total+=1

print("total",total)
