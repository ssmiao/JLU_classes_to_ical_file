import re
import xlrd
from uuid import uuid1
from icalendar import Calendar, Event
from datetime import datetime
from datetime import date
from dateutil.relativedelta import relativedelta

def upset(string):
    #string = '线性代数A◇5,6节{第1-16周}◇前卫-经信教学楼#F区第二阶梯◇宋东哲'
    subject = re.search(r'^.{1,15}◇' , string).group()[0:-1]
    time = re.search(r'◇.{1,20}◇' , string).group()[1:-1]
    building = re.search(r'◇.{1,20}#' , string).group()[1:-1]
    location = re.search(r'#.{1,20}◇' , string).group()[1:-1]
    teacher = re.search(r'◇.{1,5}$' , string).group()[1:]

    return [subject,teacher,time,building,location]

def read_excel(day):
    xlsfile = r'学生课程表.xls'

    book = xlrd.open_workbook(xlsfile)
    sheet_name=book.sheet_names()[0]
    sheet1=book.sheet_by_name(sheet_name)  #通过sheet名字来获取，当然如果你知道sheet名字了可以直接指定  

    day_data = sheet1.col_values(day+2)[3:-1]   #获得第1行的数据列表  
    return day_data

def main():
    i = 0
    dayclasses = []
    while (i < 10):
        if(i != 1 and i != 4 and i != 6):       #JLU class excel exists null col
            data = read_excel(i)

            n = 1

            while(n < 6):
                if data[n] != "":
                    if(re.search(r'\n',data[n]) != None):
                        interval = data[n].split('\n')
                        for inter_week in interval:
                            data_upset = upset(inter_week)
                            data_upset.append(data[0])
                            dayclasses.append(data_upset)

                    else:
                        data_upset = upset(data[n])
                        data_upset.append(data[0])
                        dayclasses.append(data_upset)
                n= n + 1
            write_ics(dayclasses)
        i =i+1
        

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
    start_monday = date(2018,3,5)
##############################################################


    
    #string example= ['马克思主义基本原理概论', '孙慧', '9,10,11节{第1-14周}', '前卫-经信教学楼', 'F区第一阶梯']

    for string in allclasses:
        day = string[-1]
        

        info_day = re.search(r'^.{1,15}节',string[2]).group()[0:-1].split(',')
        info_week = re.search(r'第(\d+)-(\d+)周',string[2]).group()[1:-1].split('-')

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
