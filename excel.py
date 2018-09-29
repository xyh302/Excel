import os
import xlrd
import xlwt
from xlutils.copy import copy

dirname = 'C:/Users/shentuxd/Desktop/信息部'
files = os.listdir(dirname)

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("8.15-8.31")

sheet.write_merge(0, 0, 0, 14, "")

rows = [u'序号', u'姓名', u'8.15周三', u'8.16周四', u'8.17周五', u'8.20周一',
        u'8.21周二', u'8.22周三', u'8.23周四', u'8.24周五', u'8.27周一', u'8.28周二',
        u'8.29周三', u'8.30周四', u'8.31周五']
for i in range(len(rows)):
    sheet.write(1, i, rows[i])  # 将上面的rows写入excel

time_lists = ['2018/08/15', '2018/08/16', '2018/08/17', '2018/08/20', '2018/08/21',
              '2018/08/22', '2018/08/23', '2018/08/24', '2018/08/27', '2018/08/28',
              '2018/08/29', '2018/08/30', '2018/08/31']

for i, file in enumerate(files):
    data = xlrd.open_workbook(file, formatting_info=True)  # 打开其中一个excel表
    table = data.sheets()[0]  # 提取第一个表格,sheet1
    col1 = table.col_values(0)  # 提取第一列
    filename = os.path.splitext(file)[0]  # 当前文件的文件名(无后缀)
    
    sheet.write_merge(2*i+2, 2*i+3, 0, 0, i+1)
    sheet.write_merge(2*i+2, 2*i+3, 1, 1, filename)
    
    date = '2018/07/31'
    count = 0
    quantity = 0
    for index,j in enumerate(col1):  # 在一个人的表格中for循环
        if '2018' in j:  # 找出日期的单元格
            time = table.cell(index, 5).value  # 得到当前单元格的时间
            number = col1.count(j)  # 得到相同日期的数量
            if j != date:
                start = time
                end = "00:00:00"
                date = j
            if j == date and time > end:
                end = time
                count += 1  # 如果j==date，则count+1直到==number，此时end为最后的时间
                if count == number:
                    count = 0  # count == number时，归零count
                    #print(date, start, end, number)
                    
                    # 如果将要导入表没有原始表读到的日期，就直接把原始表读到的日期写进导入表对应的位置
                    if j[6:].replace("/", ".") != rows[quantity+2][:4]:  
                        print(j,file)
                        loc = list(map(lambda x: x[:4], rows)).index(j[6:].replace("/", "."))
                        sheet.write(2*i+2, loc, start[:5])
                        sheet.write(2*i+3, loc, end[:5])
                        continue
                    sheet.write(2*i+2, 2+quantity, start[:5])
                    sheet.write(2*i+3, 2+quantity, end[:5])
                    quantity += 1
                    
workbook.save('8月份打卡记录.xls')
print('xlsx格式表格写入成功')
