import datetime
import calendar

def get_last_month_one_weekday(day_of_week):
    # 获取上个月全部周几

    day_of_week_list = {
        "mon": 0,
        "tue": 1,
        "wed": 2,
        "thu": 3,
        "fri": 4,
        "sta": 5,
        "sun": 6,
    }

    if type(day_of_week) == str:
        day_of_week = day_of_week_list[day_of_week]

    last_month_thus = []
    year = int(datetime.datetime.today().strftime("%Y"))
    now_month = int(datetime.datetime.today().strftime("%m"))
    if now_month == 1:
        last_month = 12
        year -= 1
    else:
        last_month = now_month - 1

    last_month_last_day = calendar.monthrange(year, last_month)[1]

    for i in range(1, 8):
        iday = datetime.datetime.strptime(
            (str(year) + "-" + str(last_month) + "-" + str(i)), '%Y-%m-%d')
        if iday.weekday() == day_of_week:  # 0-6,周一到周末,可根据需要自行调整
            first_weekday = int(iday.strftime("%d"))
            last_month_thus.append(iday.strftime("%Y-%m-%d"))
            break

    while first_weekday <= last_month_last_day - 7:
        first_weekday += 7
        last_month_thus.append(datetime.datetime.strptime(
             (str(year) + "-" + str(last_month) + "-" + str(first_weekday)), '%Y-%m-%d').strftime("%Y-%m-%d"))


    return last_month_thus


# print(get_last_month_one_weekday("tue"))
# print(get_last_month_one_weekday(0))
# print(get_last_month_one_weekday(1))
# print(get_last_month_one_weekday(2))
# print(get_last_month_one_weekday(3))
# print(get_last_month_one_weekday(4))
# print(get_last_month_one_weekday(5))
# print(get_last_month_one_weekday(6))
