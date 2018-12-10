import os
import xlrd
import xlwt
from xlutils.copy import copy

dirname = 'C:/Users/shentuxd/Desktop/ITC'
files = os.listdir(dirname)
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("11.01-11.30")

sheet.write_merge(0, 0, 0, 23, "")  # 第四个参数代表要合并的单元格有多少列,19代表有20列

rows = [u'序号', u'姓名', '11.01周四', '11.02周五', '11.05周一', '11.06周二', '11.07周三',
        '11.08周四', '11.09周五', '11.12周一', '11.13周二', '11.14周三','11.15周四', '11.16周五',
        '11.19周一', '11.20周二', '11.21周三', '11.22周四', '11.23周五', '11.26周一', '11.27周二',
        '11.28周三', '11.29周四', '11.30周五']
for i in range(len(rows)):
    sheet.write(1, i, rows[i])  # 将上面的rows写入excel,就是列名

time_lists = ['2018/11/03', '2018/11/04', '2018/11/10', '2018/11/11', '2018/11/17', '2018/11/18',
              '2018/11/24', '2018/11/25']

for i, file in enumerate(files):
    data = xlrd.open_workbook(dirname+'/'+file, formatting_info=True)  # 打开其中一个excel表
    table = data.sheets()[0]  # 提取第一个表格,sheet1
    col1 = table.col_values(0)  # 提取第一列
    filename = os.path.splitext(file)[0]  # 当前文件的文件名(无后缀)
    print(filename,i)
    
    sheet.write_merge(2*i+2, 2*i+3, 0, 0, i+1)
    sheet.write_merge(2*i+2, 2*i+3, 1, 1, filename)
    
    date = '2018/10/31'
    count = 0
    quantity = 0
    for index, j in enumerate(col1):  # 在一个人的表格中for循环
        if '2018' in j:  # 找出日期的单元格
            time = table.cell(index, 5).value  # 得到当前单元格的时间
            number = col1.count(j)  # 得到相同日期的数量
            if j in time_lists:  # 如果j是休息日，则跳过这次循环
                continue
            if j != date:
                start = time
                end = "00:00:00"
                date = j
            if j == date and time > end:
                end = time
                count += 1  # 如果j==date，则count+1直到==number，此时end为最后的时间
                if count == number:
                    count = 0  # count == number时，归零count

                    excel_date = j[5:].replace("/", ".")  # 取j中的日期部分，11.11

                    loc = list(map(lambda x: x[:5], rows)).index(excel_date)  # 找到rows中excel_date的位置,就是第几列
                    print(j,file,i,loc,"ex")
                    sheet.write(2*i+2, loc, start[:5])
                    sheet.write(2*i+3, loc, end[:5])
                    
workbook.save('11月份打卡记录.xls')
print('xlsx格式表格写入成功')
