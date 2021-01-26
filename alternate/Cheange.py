import xlwt,xlrd
from xlutils.copy import copy
data = xlrd.open_workbook('excel_test.xls',formatting_info=True)
excel = copy(wb=data) # 完成xlrd对象向xlwt对象转换
excel_table = excel.get_sheet(0) # 获得要操作的页
table = data.sheets()[0]
nrows = table.nrows # 获得行数
ncols = table.ncols # 获得列数
values = ["E","X","C","E","L"] # 需要写入的值
for value in values:
    excel_table.write(nrows,1,value) # 因为单元格从0开始算，所以row不需要加一
    nrows = nrows+1
excel.save('excel_test.xls')