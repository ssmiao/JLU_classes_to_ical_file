import re
import xlrd
from uuid import uuid1
from icalendar import Calendar, Event
from datetime import datetime
from datetime import date
from dateutil.relativedelta import relativedelta

start_monday = date(2019,8,26)

def read_excel(day):
    xlsfile = r'学生课程表.xls'

    book = xlrd.open_workbook(xlsfile)
    sheet_name=book.sheet_names()[0]
    sheet1=book.sheet_by_name(sheet_name)  #通过sheet名字来获取，当然如果你知道sheet名字了可以直接指定  

    day_data = sheet1.col_values(day+2)[3:-1]   #获得从第四行开始直到结束的数据列表  

    return day_data

class Course:
    def __init__(self,string):
        self.string = string

        self.subject = re.search(r'^(.*?)◇' , string).group()[0:-1]
        self.time_location = re.search(r'◇(.*)◇' , string).group()[1:-1].split("◇")
        self.time = self.time_location[0]
        
        try:
            self.building = self.time_location[1].split("#")[0]
            self.location = self.time_location[1].split("#")[1]
        except IndexError:      #This means the file not specify the location or building,the common situation is "地点待定".
            self.building = self.location = self.time_location[1]
        
        self.teacher = re.search(r'◇.(.*?)$' , string).group()[1:]
        self.course_str = [self.subject,self.teacher,self.time,self.building,self.location]

def main():
    dayclasses = []
    for index in range(10):
        if(index != 1 and index != 4 and index != 6):       #JLU class excel exists null col ->col D,G,I
            data = read_excel(index)
            for course_no in range(1,5):
                if data[course_no] != "":
                    if(re.search(r'\n',data[course_no]) != None):   #This means there are serval courses in the same time at different week.
                        same_time_course = data[course_no].split('\n')
                        for course in same_time_course:
                            each_course = Course(course)
                            dayclasses.append(each_course.course_str+[data[0]])

                    else:
                        each_course = Course(data[course_no])
                        dayclasses.append(each_course.course_str+[data[0]])
        write_ics(dayclasses)


def write_ics(allclasses):

    cal = Calendar()
    cal['version'] = '2.0'
    cal['prodid'] = '-//JLU//Farthing//CN'
    cal['calscale'] = 'GREGORIAN'
    cal['method'] = 'PUBLISH'

    dict_week = {'星期一': 0, '星期二': 1, '星期三': 2, '星期四': 3, '星期五': 4, '星期六': 5, '星期日': 6}
    dict_day = {1: relativedelta(hours=8, minutes=00), 2:relativedelta(hours=8,minutes=55),
                3: relativedelta(hours=10, minutes=00),4:relativedelta(hour=10,minutes=55),
                5: relativedelta(hours=13, minutes=30),6:relativedelta(hours=14, minutes=25),
                7: relativedelta(hours=15,minutes=30),8:relativedelta(hours=16,minutes=25),
                9: relativedelta(hours=18, minutes=30),10:relativedelta(hours=19,minutes=25),
                11:relativedelta(hours=20,minutes=20)}

##############################################################

    #string example= ['马克思主义基本原理概论', '孙慧', '9,10,11节{第1-14周}', '前卫-经信教学楼', 'F区第一阶梯']

    for string in allclasses:
        day = string[-1]
        
        try:
            info_day = re.search(r'^.{1,15}节',string[2]).group()[0:-1].split(',')
            info_week = re.search(r'第(\d+)-(\d+)周',string[2]).group()[1:-1].split('-')
        except:
            print(string)

        dtstart_date = start_monday + relativedelta(weeks=(int(info_week[0]) - 1)) + relativedelta(
            days= (dict_week[day]))

        for class_num in info_day:
            event = Event()
            
            dtstart_datetime = datetime.combine(dtstart_date, datetime.min.time())
            dtstart = dtstart_datetime + dict_day[int(class_num)]
            dtend = dtstart + relativedelta(minutes=45)
            
            if string[2].find('|') == -1:
                  interval = 1
            else:
                  interval = 2
                  if(re.search(r'\|双周',string[2]) != None):
                    dtstart = (dtstart + relativedelta(weeks=1))
     
            event.add('uid', str(uuid1()) + '@FarthingJLU')
            event.add('summary', string[0])
            event.add('dtstamp', datetime.now())
            event.add('dtstart', dtstart)
            event.add('dtend', dtend)

            event.add('rrule',
                        {'freq': 'weekly', 'interval': interval,
                         'count': int(info_week[1]) - int(info_week[0]) + 1})

            event.add('location', string[4]+string[3])
            cal.add_component(event)

        print(info_day)
        print(info_week)

    with open('output.ics', 'w+', encoding='utf-8') as file:
        file.write(cal.to_ical().decode('utf-8'.replace('\r\n', '\n')).replace('\n','').strip())
    return cal

if (__name__ == '__main__'):
    main()
