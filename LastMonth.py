import datetime
import calendar


class Month:
    last_month_thus = []
    last_month_all = []
    year = int(datetime.datetime.today().strftime("%Y"))
    now_month = int(datetime.datetime.today().strftime("%m"))

    if now_month == 1:
        last_month = 12
        year -= 1
    else:
        last_month = now_month - 1

    last_month_last_day = calendar.monthrange(year, last_month)[1]

    def get_last_month_one_weekday(self, day_of_week):
        # 获取上个月全部周几
        for i in range(1, 8):
            iday = datetime.datetime.strptime(
                (str(self.year) + "/" + str(self.last_month) + "/" + str(i)), '%Y/%m/%d')
            if iday.weekday() == day_of_week:  # 0-6,周一到周末,可根据需要自行调整
                first_weekday = int(iday.strftime("%d"))
                self.last_month_thus.append(iday.strftime("%Y/%m/%d"))
                break

        while first_weekday <= self.last_month_last_day - 7:
            first_weekday += 7
            self.last_month_thus.append(datetime.datetime.strptime(
                (str(self.year) + "/" + str(self.last_month) + "/" + str(first_weekday)),
                '%Y/%m/%d').strftime("%Y/%m/%d"))

        return self.last_month_thus

    def get_last_month_all(self):
        for i in range(1,self.last_month_last_day+1):
            self.last_month_all.append(datetime.datetime.strptime(
                (str(self.year) + "/" + str(self.last_month) + "/" + str(i)),
                '%Y/%m/%d').strftime("%Y/%m/%d"))

        return self.last_month_all


if __name__ == '__main__':
    print(Month.last_month)
    print(Month().get_last_month_all())
# print(get_last_month_one_weekday("tue"))
# print(get_last_month_one_weekday(0))
# print(get_last_month_one_weekday(1))
# print(get_last_month_one_weekday(2))
# print(get_last_month_one_weekday(3))
# print(get_last_month_one_weekday(4))
# print(get_last_month_one_weekday(5))
# print(get_last_month_one_weekday(6))
