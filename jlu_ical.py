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

    #print ('Your name is ' + sheet1.cell_value(1,4)+'\n') 

    day_data = sheet1.col_values(day+2)[3:-1]   #获得第1行的数据列表  
    return day_data

def main():
    i = 0
    allclasses = []
    while (i < 10):
        if(i != 1 and i != 4 and i != 6):       #JLU class excel exists null col
            data = read_excel(i)
            #print(data[0])
            #if(data != ''):
            n = 1
            #print(data[0]+'\n')
            while(n < 6):
                if data[n] != "":
                    if(re.search(r'\n',data[n]) != None):
                        interval = data[n].split('\n')
                        for inter_week in interval:
                            data_upset = upset(inter_week)
                            data_upset.append(data[0])
                            allclasses.append(data_upset)
                            #allclasses.append((upset(inter_week)).append(data[0]))
                            #data_upset = (upset(inter_week))
                            #write_ics(data_upset,data[0])

                    else:
                        data_upset = upset(data[n])
                        data_upset.append(data[0])
                        allclasses.append(data_upset)
                        #print(upset(data[n]))
                        #data_upset = (upset(data[n]))

                        #write_ics(data_upset,data[0])
                n= n + 1
                #print('\n')
        i =i+1
    write_ics(allclasses)

def write_ics(allclasses):
    #print(allclasses)

    cal = Calendar()
    cal['version'] = '2.0'
    cal['prodid'] = '-//CQUT//Syllabus//CN'
    dict_week = {'星期一': 0, '星期二': 1, '星期三': 2, '星期四': 3, '星期五': 4, '星期六': 5, '星期日': 6}
    dict_day = {1: relativedelta(hours=8, minutes=00), 2:relativedelta(hours=8,minutes=55),
                3: relativedelta(hours=10, minutes=00),4:relativedelta(hour=10,minutes=55),
                5: relativedelta(hours=13, minutes=30),6:relativedelta(hours=14, minutes=25),
                7: relativedelta(hours=15,minutes=30),8:relativedelta(hours=16,minutes=25),
                9: relativedelta(hours=18, minutes=30),10:relativedelta(hours=19,minutes=25),
                11:relativedelta(hours=20,minutes=20)}
    start_monday = date(2017,3,5)
##############################################################

    #print(string)
    
    #string example= ['马克思主义基本原理概论', '孙慧', '9,10,11节{第1-14周}', '前卫-经信教学楼', 'F区第一阶梯']
    #print(string)
    for string in allclasses:
        day = string[-1]
        
        #event = Event()
        info_day = re.search(r'^.{1,15}节',string[2]).group()[0:-1].split(',')
        #print(info_day)
        info_week = re.search(r'第(\d+)-(\d+)周',string[2]).group()[1:-1].split('-')

        dtstart_date = start_monday + relativedelta(weeks=(int(info_week[0]) - 1)) + relativedelta(
            days= (dict_week[day]+1))

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


        ###单双周导入存在问题

            event.add('rrule',
                        {'freq': 'weekly', 'interval': interval,
                         'count': int(info_week[1]) - int(info_week[0]) + 1})


            event.add('location', string[4]+string[3])
            cal.add_component(event)

        print(info_day)
        print(info_week)

    with open('output.ics', 'w+', encoding='utf-8') as file:
        file.write(cal.to_ical().decode('utf-8'.replace('\r\n', '\n').strip()))
    return cal

if (__name__ == '__main__'):
    main()
