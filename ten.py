import os
import xlrd
import xlwt
from xlutils.copy import copy

dirname = 'C:/Users/shentuxd/Desktop/ITC'
files = os.listdir(dirname)
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("10.8-10.31")

sheet.write_merge(0, 0, 0, 19, "")  # 第四个参数代表要合并的单元格有多少列,19代表有20列

rows = [u'序号', u'姓名', '10.08周一', '10.09周二', '10.10周三', '10.11周四', '10.12周五',
        '10.15周一', '10.16周二', '10.17周三', '10.18周四', '10.19周五','10.22周一', '10.23周二',
        '10.24周三', '10.25周四', '10.26周五', '10.29周二', '10.30周三', '10.31周四']
for i in range(len(rows)):
    sheet.write(1, i, rows[i])  # 将上面的rows写入excel,就是列名

time_lists = ['2018/10/01', '2018/10/02', '2018/10/03', '2018/10/04', '2018/10/05', '2018/10/06',
              '2018/10/07', '2018/10/13', '2018/10/14', '2018/10/20', '2018/10/21', '2018/10/27', '2018/10/28']

for i, file in enumerate(files):
    data = xlrd.open_workbook(dirname+'/'+file, formatting_info=True)  # 打开其中一个excel表
    table = data.sheets()[0]  # 提取第一个表格,sheet1
    col1 = table.col_values(0)  # 提取第一列
    filename = os.path.splitext(file)[0]  # 当前文件的文件名(无后缀)
    print(filename,i)
    if filename[-1] == '2':  # 如果文件名最后一个字符是'2',i索引值-1
        i -= 1
    else:
        sheet.write_merge(i+2, i+3, 0, 0, i/2+1)  # 写序号,第i+2行,第i+3列
        sheet.write_merge(i+2, i+3, 1, 1, filename)  # 写人员姓名
        #write_merge(x, x + h, y, w + y, string, sytle)
        #x表示行，y表示列，w表示跨列个数，h表示跨行个数，string表示要写入的单元格内容，style表示单元格样式。
        #注意，x，y，w，h，都是以0开始计算的。
    
    date = '2018/09/31'
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

                    excel_date = j[5:].replace("/", ".")
                    # 如果将要导入表没有原始表读到的日期，就直接把原始表读到的日期写进导入表对应的位置，然后跳过这次循环
                    if excel_date != rows[quantity+2][:5]:  
                        loc = list(map(lambda x: x[:5], rows)).index(excel_date)  # 找到rows中excel的位置,就是第几列
                        print(j,file,i,loc)
                        sheet.write(i+2, loc, start[:5])
                        sheet.write(i+3, loc, end[:5])
                        continue
                    sheet.write(i+2, 2+quantity, start[:5])
                    sheet.write(i+3, 2+quantity, end[:5])
                    quantity += 1
                    
workbook.save('10月份打卡记录.xls')
print('xlsx格式表格写入成功')
