import os
import xlrd
import xlwt
from xlutils.copy import copy

dirname = 'C:/Users/shentuxd/Desktop/ITC/ITC'
files = os.listdir(dirname)
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("9.3-9.30")

sheet.write_merge(0, 0, 0, 22, "")

rows = [u'序号', u'姓名', '9.03周一', '9.04周二', '9.05周三', '9.06周四', '9.07周五', '9.10周一', '9.11周二', '9.12周三', '9.13周四', '9.14周五',
        '9.17周一', '9.18周二', '9.19周三', '9.20周四', '9.21周五', '9.25周二', '9.26周三', '9.27周四', '9.28周五', '9.29周六', '9.30周日']
for i in range(len(rows)):
    sheet.write(1, i, rows[i])  # 将上面的rows写入excel

time_lists = ['2018/09/01', '2018/09/02', '2018/09/08', '2018/09/09', '2018/09/15', '2018/09/16', '2018/09/22', '2018/09/23', '2018/09/24']

for i, file in enumerate(files):
    data = xlrd.open_workbook(dirname+'/'+file, formatting_info=True)  # 打开其中一个excel表
    table = data.sheets()[0]  # 提取第一个表格,sheet1
    col1 = table.col_values(0)  # 提取第一列
    filename = os.path.splitext(file)[0]  # 当前文件的文件名(无后缀)
    print(filename,i)
    if filename[-1] == '2':
        i -= 1
    else:
        sheet.write_merge(i+2, i+3, 0, 0, i/2+1)  # 写序号
        sheet.write_merge(i+2, i+3, 1, 1, filename)  # 写人员姓名
    
    date = '2018/08/31'
    count = 0
    quantity = 0
    for index,j in enumerate(col1):  # 在一个人的表格中for循环
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
                    
                    # 如果将要导入表没有原始表读到的日期，就直接把原始表读到的日期写进导入表对应的位置，然后跳过这次循环
                    if j[6:].replace("/", ".") != rows[quantity+2][:4]:  
                        loc = list(map(lambda x: x[:4], rows)).index(j[6:].replace("/", "."))
                        print(j,file,i,loc)
                        sheet.write(i+2, loc, start[:5])
                        sheet.write(i+3, loc, end[:5])
                        continue
                    sheet.write(i+2, 2+quantity, start[:5])
                    sheet.write(i+3, 2+quantity, end[:5])
                    quantity += 1
                    
workbook.save('9月份打卡记录.xls')
print('xlsx格式表格写入成功')
