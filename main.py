import xlrd
import datetime

# 文件名
FILE_NAME = r"我的课表.xlsx"
# 开学时间 年/月/日
START_YEAR = 2024; START_MONTH = 2; START_DAY = 26
# 课程开始时间
CLASS_START_TIME = ["080000", "085500", "100000", "105500", "133000", "142500", "153000", "162500", "182000", "190500", "200000", "204500"]
# 一节课程持续时间
CLASS_LAST_MINUTES = 45

def getEndTime(s):
    hour = int(s[0:2])
    minute = int(s[2:4])
    second = int(s[4:6])
    minute += CLASS_LAST_MINUTES
    if minute >= 60:
        minute -= 60
        hour += 1
        if hour >= 24:
            hour -= 24
            # day += 1
    return "%02d" % hour + "%02d" % minute + "%02d" % second

class DataType:
    # 13节课的开始时间常数
    def __init__(self, id, name, startWeek, endWeek, specialWeek, week, startTime, endTime, loc):
        self.id = id
        self.name = name
        self.startWeek = startWeek
        self.endWeek = endWeek
        self.specialWeek = specialWeek
        self.week = week
        self.startTime = startTime
        self.endTime = endTime
        self.loc = loc
    def getStartDate(self):
        startDay = datetime.date(year=START_YEAR, month=START_MONTH, day=START_DAY)
        deltaDay = (self.startWeek - 1) * 7 + self.week - 1
        return (startDay + datetime.timedelta(days=deltaDay)).strftime("%Y%m%d")
    def getEndDate(self):
        startDay = datetime.date(year=START_YEAR, month=START_MONTH, day=START_DAY)
        deltaDay = (self.endWeek - 1) * 7 + self.week - 1
        return (startDay + datetime.timedelta(days=deltaDay)).strftime("%Y%m%d")

    def getStartTime(self):
        return CLASS_START_TIME[self.startTime - 1]
    def getEndTime(self):
        return getEndTime(CLASS_START_TIME[self.endTime - 1])
    def __repr__(self):
        return 'class content: \nid: ' + str(self.id) + '\nname: ' + str(self.name) + '\nweeks: ' + str(self.startWeek) + '~' + str(self.endWeek) + '(' +  str(self.specialWeek) + ')' + '\nweek: ' + str(self.week) + '\ntime: ' + str(self.startTime) + '~' + str(self.endTime) + '\nloc: ' + str(self.loc) + '\n'
events = []
def getName(s):
    idx = 0
    for c in s:
        if c == '*' or c == '[':
            break
        idx += 1
    return s[0: idx]
def getStartWeek(s):
    idx = 0
    for c in s:
        if c == '-' or c == r'周':
            break
        idx += 1
    return int(s[0: idx])
def getEndWeek(s):
    idx = 0
    time = 0
    left = 0
    for c in s:
        if c == '-' or c == r'周':
            time += 1
            if time == 1:
                if c == r'周':
                    return int(s[0: idx])
                left = idx + 1
            else:
                break
        idx += 1
    return int(s[left: idx])
def getSpecialWeek(s):
    if s.count(r'单'):
        return 1
    elif s.count(r'双'):
        return 0
    return -1
def getWeek(s):
    return int(s[2])
def getTime(s):
    idx = start = end = 0;
    for c in s:
        idx += 1
        if c == r'第':
            start = idx
        elif c == r'节':
            end = idx
    return int(s[start: end - 1])
def getLocation(s):
    ss = s.split('-')
    if len(ss) < 2:
        return s
    return ss[1] + ss[2]
def decodeColumn(column):
    for s in column:
        if len(s) > 0:
            ss = s.split('\n')
            id  = name = weeks = week = time = loc = ""
            for sss in ss:
                if len(sss) > 0:
                    t = sss.split('-')
                    if t[0][0:2].islower():
                        id = t[0]
                        name = t[1]
                    else:
                        t = sss.split(',')
                        if len(t) > 1:
                            weeks = ""
                            for info in t:
                                if info.count(r'周') > 0:
                                    weeks += info + ','
                                elif info.count(r'星期') > 0:
                                    week = info
                                elif len(info) > 0:
                                    idx = 0
                                    cnt = 0
                                    for c in info:
                                        idx += 1
                                        if c == r'节':
                                            cnt += 1
                                            if cnt == 2:
                                                break
                                    time = info[0: idx]
                                    loc = info[idx:]
                            items = weeks.split(',')
                            for wks in items:
                                if len(wks) > 0:
                                    times = time.split('-')
                                    tmp = DataType(id, getName(name), getStartWeek(wks), getEndWeek(wks), getSpecialWeek(wks), getWeek(week), getTime(times[0]), getTime(times[1]), getLocation(loc))
                                    notIn = True
                                    for event in events:
                                        if tmp.__dict__ == event.__dict__:
                                            notIn = False
                                            break
                                    if notIn:
                                        events.append(tmp)
wkst = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]
def weekToUpperCase(week):
    return wkst[week - 1]
def OutputEvent(event):
    f.write("BEGIN:VEVENT\n")
    f.write("DESCRIPTION:" + event.id + "\n")
    f.write("DTEND;TZID=Asia/Shanghai:" + event.getStartDate() + "T" + event.getEndTime() + "\n")
    f.write("DTSTART;TZID=Asia/Shanghai:" + event.getStartDate() + "T" + event.getStartTime() + "\n")
    f.write("LOCATION:" + event.loc + "\n");
    if event.specialWeek == -1:
        f.write("RRULE:FREQ=WEEKLY;INTERVAL=1;" + "UNTIL=" + event.getEndDate() + "T" + event.getEndTime() + "Z;BYDAY=" + weekToUpperCase(event.week) + "\n")
    else:
        f.write("RRULE:FREQ=WEEKLY;INTERVAL=2;" + "UNTIL=" + event.getEndDate() + "T" + event.getEndTime() + "Z;BYDAY=" + weekToUpperCase(event.week) + "\n")
    f.write("SUMMARY:" + event.name + "\n")
    f.write("BEGIN:VALARM\n")
    f.write("TRIGGER:-PT30M\n")
    f.write("END:VALARM\n")
    f.write("END:VEVENT\n")

# Main
data = xlrd.open_workbook(FILE_NAME)
f = open("Timetable.ics", "w")
f.write("BEGIN:VCALENDAR\n")
f.write("CALSCALE:GREGORIAN\n")
f.write("VERSION:2.0\n")
table = data.sheet_by_index(0)
for i in range(2, 9):
    decodeColumn(table.col_values(i))
for event in events:
    OutputEvent(event)
f.write("END:VCALENDAR\n")
f.close()