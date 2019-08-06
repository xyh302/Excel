import os
import xlrd
import xlwt
from datetime import datetime, timedelta
from LastMonth import Month

dirname = os.getcwd()
files = os.listdir(dirname)
files = [i for i in files if i.endswith('xls') or i.endswith('xlsx')]
workbook = xlwt.Workbook()

month = Month()
lastmonth = str(month.last_month)
sheet = workbook.add_sheet(lastmonth + ".01-" + lastmonth + "." + str(month.last_month_last_day))

alist = month.get_last_month_one_weekday(5) + month.get_last_month_one_weekday(6)
time_lists = alist + ['2019/01/01', '2019/02/04', '2019/02/05', '2019/02/06',
                      '2019/02/07', '2019/02/08', '2019/04/05', '2019/05/01',
                      '2019/06/07', '2019/09/13', '2019/10/01', '2019/10/02',
                      '2019/10/03', '2019/10/04', '2019/10/07']


def get_weekday(list):
    week_day_dict = {
        0: '周一',
        1: '周二',
        2: '周三',
        3: '周四',
        4: '周五',
        5: '周六',
        6: '周天',
    }
    datetime_list = [i[5:].replace('/', '.') + week_day_dict[
        datetime.strptime(i, '%Y/%m/%d').weekday()] for i in list]
    return datetime_list


def get_rows_and_time(time_lists):
    time_allday = month.get_last_month_all()
    time_lists = [i for i in time_lists if i not in [
        '2019/02/02', '2019/02/03', '2019/09/29', '2019/10/12']]  # 去除补加班的周末日期
    time_workday = [i for i in time_allday if i not in time_lists]
    workday = get_weekday(time_workday)
    return workday, time_lists


workday, time_lists = get_rows_and_time(time_lists)  # 得到工作日和休息日的日期

rows = [u'序号', u'姓名'] + workday

sheet.write_merge(0, 0, 0, len(rows) - 1, "")  # 第四个参数代表要合并的单元格有多少列,19代表有20列

for i in range(len(rows)):
    sheet.write(1, i, rows[i])  # 将上面的rows写入excel,就是列名


def get_time(ctime):  # 因系统差,所有时间减去2分钟
    ctime = (datetime.strptime(ctime, '%H:%M') - timedelta(minutes=0)).strftime("%H:%M")  # 在这里不用减时间
    return ctime


for i, file in enumerate(files):
    data = xlrd.open_workbook(dirname + '/' + file, formatting_info=True)  # 打开其中一个excel表
    table = data.sheets()[0]  # 提取第一个表格,sheet1
    col1 = table.col_values(0)  # 提取第一列
    filename = os.path.splitext(file)[0]  # 当前文件的文件名(无后缀)
    # print(filename, i)
    if filename[-1] == '1':
        i -= 1
    else:
        sheet.write_merge(i + 2, i + 3, 0, 0, i / 2 + 1)  # 写序号
        sheet.write_merge(i + 2, i + 3, 1, 1, filename)  # 写人员姓名
    # sheet.write_merge(2 * i + 2, 2 * i + 3, 0, 0, i + 1)
    # sheet.write_merge(2 * i + 2, 2 * i + 3, 1, 1, filename)

    date = '2018/12/31'
    count = 0
    quantity = 0
    for index, j in enumerate(col1):  # 在一个人的表格中for循环
        if '2019' in j:  # 找出日期的单元格
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
                    #
                    # loc = list(map(lambda x: x[:5], rows)).index(excel_date)  # 找到rows中excel_date的位置,就是第几列
                    # # print(j, file, i, loc, "ex")
                    # sheet.write(2 * i + 2, loc, get_time(start[:5]))
                    # sheet.write(2 * i + 3, loc, get_time(end[:5]))

                    # 如果将要导入表没有原始表读到的日期，就直接把原始表读到的日期写进导入表对应的位置，然后跳过这次循环
                    if excel_date != rows[quantity + 2][:5]:
                        loc = list(map(lambda x: x[:5], rows)).index(excel_date)
                        print(j, file, i, loc)
                        sheet.write(i + 2, loc, start[:5])
                        sheet.write(i + 3, loc, end[:5])
                        continue
                    sheet.write(i + 2, 2 + quantity, start[:5])
                    sheet.write(i + 3, 2 + quantity, end[:5])
                    quantity += 1

workbook.save(lastmonth + '月份打卡记录.xls')
print('xls格式表格写入成功')
