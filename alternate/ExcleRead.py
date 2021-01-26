import socket

import xlrd
import xlwt
import urllib.request as request
from urllib import error
from bs4 import BeautifulSoup


# 函数
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height

    # borders= xlwt.Borders()
    # borders.left= 6
    # borders.right= 6
    # borders.top= 6
    # borders.bottom= 6

    style.font = font
    # style.borders = borders

    return style


workbook = xlrd.open_workbook(r'excleData.xlsx')
# 获取所有sheet
sheets_name = workbook.sheet_names()  # [u'sheet1', u'sheet2']
sheet1_name = sheets_name[0]
# 得到要处理的数据
sheet1_obj = workbook.sheet_by_name(sheet1_name)

# 循环处理数据
excle_data = []
for x in range(0, sheet1_obj.nrows):
    item = []
    deal_url = sheet1_obj.cell_value(x, 0)
    target_url = sheet1_obj.cell_value(x, 1)
    item.append(deal_url)
    item.append(target_url)
    excle_data.append(item)

print('--------------------------------------------------------------------------------------------------------')
print('--------------------------------------------------------------------------------------------------------')
print('--------------------------------------------------------------------------------------------------------')
print(excle_data)
# 导出excle表格
f = xlwt.Workbook()  # 创建工作簿
my_sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
row0 = ['测试网页地址', '期望href', '搜索到相关内容如下', '测试结果']
# 生成第一行
for i in range(0, len(row0)):
    my_sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))

for x in range(0, len(excle_data)):
    for y in range(0, len(excle_data[x])):
        print(excle_data[x][y])
        print(type(excle_data[x][y]))
        my_sheet1.write(x+1, y, excle_data[x][y], set_style('Times New Roman', 220, True))

f.save('excleTest.xls')  # 保存文件