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


workbook = xlrd.open_workbook(r'urldata.xlsx')
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
    # 开始爬虫程序
    print('开始爬取' + deal_url + '内容：')
    try:
        # 伪装header
        headers = ('User-Agent','Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11')
        opener = request.build_opener()
        opener.addheaders = [headers]
        origin_bytes = opener.open(deal_url, timeout=100).read()
    except socket.timeout as e:
        print('爬取失败，网站异常：')
        item.append('')
        item.append('网站连接超时')
        excle_data.append(item)
        continue

    html_string = origin_bytes.decode('utf-8')
    # origin_bytes = request.urlopen(url=deal_url, timeout=50,headers=ua_headers).read()
    # html_string = origin_bytes.decode('utf-8')

    print('爬取成功！')
    print('检索rel=alternate：')
    print('检索结果如下：')
    soup = BeautifulSoup(html_string, 'html.parser')
    search_result = soup.find_all(rel='alternate')
    paragraphs = []
    for x in search_result:
        paragraphs.append(str(x))

    # search_result_string = ' '.join(search_result)
    item.append(paragraphs)
    test_result = 'PASS'
    # 检索结果不唯一，不通过
    if len(search_result) != 1:
        print('测试结果：FAIL')
        test_result = 'FAIL'
        item.append('FAIL')
    else:
        for link_data in search_result:
            print('期望的href:' + target_url)
            href = link_data.get('href')
            print('检测到的href:' + href)
            if href != target_url:
                print('测试结果：FAIL')
                test_result = 'FAIL'
                item.append('FAIL')
    if test_result == 'PASS':
        print('测试结果：PASS')
        item.append('PASS')
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

f.save('result_report.xlsx')  # 保存文件
