# -*- coding: utf-8 -*-
import os
import sys
import time
import calendar
from datetime import datetime, timedelta

# 读取的excel地址
from auto_install import import_tools

inputExcel = "/Users/mylo/Desktop/6月李怡萱read.xls"
# 统计完写入的excel地址
outputExcel = "/Users/mylo/Desktop/OverWorkExcel.xls"
# 第几个sheet需要被统计 从0开始计算
sheetIndex = 0
# 加班餐补单位费用
overWorkPay = 30
# 节假日 例:"2018-03-07"
holidays = []
# 额外工作日 例: "2018-03-31"
extra_workdays = []


# 返回时间戳
def get_timestamp(timestr):
    return int(time.mktime(get_time(timestr)))


# 获取时间
def get_time(timestr):
    try:
        return time.strptime(timestr, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return time.strptime(timestr, "%Y-%m-%d")


# 格式化时间
def get_standard_format_date(unformated_date):
    from xlrd import xldate_as_tuple
    if str(unformated_date).__contains__("."):
        # 修改过时间，时间变成了float
        d = datetime(*xldate_as_tuple(unformated_date, 0))
        return d.strftime("%Y-%m-%d %H:%M:%S")
    else:
        return unformated_date


# 判断是否是周末或节假日
def is_weekend(timestamp):
    # 中国时区！
    date = datetime.utcfromtimestamp(timestamp) + timedelta(hours=8)
    if str(date.date()) in holidays:
        return True
    elif str(date.date()) in extra_workdays:
        return False
    else:
        return date.weekday() >= 5


# sheet格式
def sheet_style(sheet, defaults, style):
    for i in range(len(defaults)):
        sheet.col(i).width = defaults[i][0]
        sheet.write(0, i, defaults[i][1], style)


def upload_time_list(random_time):
    d = datetime.strptime(random_time, "%Y-%m-%d")
    time_list = []
    year = int(d.year)
    month = int(d.month)
    # 获取当月第一天的星期和当月的总天数
    weekdays, monthRange = tuple(calendar.monthrange(year, month))
    # 获取当月的第一天
    first_day = datetime(year=year, month=month, day=1)
    last_day = datetime(year=year, month=month, day=monthRange)

    temp_day = first_day
    while temp_day <= last_day:  # 感谢怡萱！
        time_list.append([temp_day, "", 0])
        temp_day = temp_day + timedelta(days=1)
    return time_list


def calculate():
    import xlrd
    import xlwt
    from imp import reload
    reload(sys)
    sys.setdefaultencoding('utf8')
    # 打开文件
    workbook = xlrd.open_workbook(inputExcel)

    try:
        # 读取的sheet
        the_read_sheet = workbook.sheet_by_index(sheetIndex)
    except IndexError:
        print "角标越界了，老铁"
        return

        # 获取员工编号
    cols = the_read_sheet.col_values(1)
    # 员工缓存的打卡字典
    staff = {}

    # tempId = 0
    # tempid = sheet.cell(427, 1).value.encode('utf-8')
    for over_work_index in range(len(cols)):
        if cols[over_work_index].isdigit():
            daka = staff.get(cols[over_work_index])
            if cols[over_work_index] in staff:
                daka.append(the_read_sheet.row_values(over_work_index))
            else:
                staff[cols[over_work_index]] = [the_read_sheet.row_values(over_work_index)]

    # 加班的缓存字典
    over_work_dict = {}
    for sta in staff:
        over_work_time = 0
        over_work_detail = []
        over_work_days = []
        # for index in range(len(staff.get(tempid))):
        for over_work_index in range(len(staff.get(sta))):
            staff_day = staff.get(sta)[over_work_index]
            origin_time = get_standard_format_date(staff_day[3])
            on_duty_time = get_standard_format_date(staff_day[4])
            off_duty_time = get_standard_format_date(staff_day[5])
            origin_ts = get_timestamp(origin_time)
            on_duty_ts = get_timestamp(on_duty_time)
            off_duty_ts = get_timestamp(off_duty_time)

            last_off_duty_ts = 0
            if over_work_index > 0:
                last_off_duty_time = get_standard_format_date(staff.get(sta)[over_work_index - 1][5])
                last_off_duty_ts = get_timestamp(last_off_duty_time)
            t = datetime.strptime(str(staff_day[3]), "%Y-%m-%d")

            # 判断下班时间，如果第二天上班时间早于6点，也认为是前一天的下班时间
            if over_work_index < len(staff.get(sta)) - 1:
                next_on_duty_time = get_standard_format_date(staff.get(sta)[over_work_index + 1][4])
                next_origin_time = get_standard_format_date(staff.get(sta)[over_work_index + 1][3])
                next_on_duty_ts = get_timestamp(next_on_duty_time)
                next_origin_ts = get_timestamp(next_origin_time)
                if next_on_duty_ts - next_origin_ts < 6 * 3600:
                    off_duty_ts = next_on_duty_ts

            # 判断是否是周六日,并且中午13点之前到公司,且正常下班(工作时间超过8个小时,或下班时间晚于6点)
            if is_weekend(on_duty_ts) & (on_duty_ts - origin_ts < 13 * 3600) & (
                                off_duty_ts - on_duty_ts >= 8 * 3600 or off_duty_ts - origin_ts >= 18 * 3600):
                over_work_detail.append(str(staff_day[3]) + "中午")
                over_work_days.append(t)
                over_work_time += 1

            # 判断下班时间大于晚上8点
            if off_duty_ts - origin_ts > 20 * 3600:
                # 当天工作时长大于11小时 或前一天晚上10点之后回家
                if (off_duty_ts - on_duty_ts >= 11 * 3600) | (origin_ts - last_off_duty_ts <= 2 * 3600):
                    over_work_detail.append(str(staff_day[3]) + "晚上")
                    over_work_days.append(t)
                    over_work_time += 1

            over_work_dict[sta] = [staff.get(sta)[0][2], over_work_time, over_work_detail, over_work_days]

    # 判断文件是否存在，存在就删了
    if os.path.exists(outputExcel):
        try:
            os.remove(outputExcel)
        except WindowsError:
            print "要写入的文件正在打开状态,请关闭后重试"
            return

    # 写入表中
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    style = xlwt.easyxf('align: wrap yes, vert centre, horiz center;')
    # 加班表
    the_read_sheet = book.add_sheet('over_work_sheet', cell_overwrite_ok=True)
    default_reading = [[4000, "人员编号"], [4000, "姓名"], [6000, "加班次数"]]
    sheet_style(the_read_sheet, default_reading, style)
    # 加班表详情
    detail_sheet = book.add_sheet('over_work_detail_sheet', cell_overwrite_ok=True)
    default_detail = [[4000, "人员编号"], [4000, "姓名"], [6000, "加班详情"]]
    sheet_style(detail_sheet, default_detail, style)
    # 上报表
    upload_sheet = book.add_sheet('upload_sheet', cell_overwrite_ok=True)
    default_upload = [[4000, "日期&时间"], [6000, "加班原因"], [12000, "用餐人姓名"], [4000, "费用(￥)"], [4000, "备注"]]
    sheet_style(upload_sheet, default_upload, style)

    over_work_index = 1
    detail_index = 1
    for over_work_key in over_work_dict:
        # 第一个sheet放置加班姓名和时长
        name = str(over_work_dict.get(over_work_key)[0])
        times = str(over_work_dict.get(over_work_key)[1])
        the_read_sheet.write(over_work_index, 0, str(over_work_key).decode('utf-8'), style)
        the_read_sheet.write(over_work_index, 1, name.decode('utf-8'), style)
        the_read_sheet.write(over_work_index, 2, times.decode('utf-8'), style)
        over_work_index += 1

        detail = over_work_dict.get(over_work_key)[2]
        detail_sheet.write(detail_index, 0, str(over_work_key).decode('utf-8'), style)
        detail_sheet.write(detail_index, 1, name.decode('utf-8'), style)
        for detail_item_index in range(len(detail)):
            detail_sheet.write(detail_index, 2, detail[detail_item_index].decode('utf-8'), style)
            detail_index += 1

    # 上报表的绘制
    temp_date = staff.get(staff.keys()[0])[0][3]
    canlender_list = upload_time_list(temp_date)
    for sta in over_work_dict:
        over_work_item = over_work_dict.get(sta)
        for i in range(len(canlender_list)):
            for item in over_work_item[3]:
                if item == canlender_list[i][0]:
                    if over_work_item[0] in canlender_list[i][1]:
                        canlender_list[i][1] = canlender_list[i][1].replace(over_work_item[0],
                                                                            over_work_item[0] + "(中午/晚上)")
                    else:
                        canlender_list[i][1] = canlender_list[i][1] + "," + over_work_item[0]
                    canlender_list[i][2] = canlender_list[i][2] + 1
    upload_index = 1
    for item in canlender_list:
        t = item[0].strftime("%Y-%m-%d")
        upload_sheet.write(upload_index, 0, t.decode('utf-8'), style)
        upload_sheet.write(upload_index, 1, "咖啡项目加班", style)
        name_list = item[1]
        if len(name_list) > 1:
            name_list = name_list[1:]
        upload_sheet.write(upload_index, 2, name_list, style)
        upload_sheet.write(upload_index, 3, overWorkPay * item[2], style)
        upload_sheet.write(upload_index, 4, "共" + str(item[2]) + "人次", style)
        upload_index += 1

    book.save(outputExcel)
    print "统计已完成，请在" + outputExcel + "中查看加班统计信息"


if __name__ == '__main__':
    import_tools()
    calculate()
