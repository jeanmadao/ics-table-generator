import os
import requests
import datetime
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()
SCHEDULE_ICS_URL = os.getenv('SCHEDULE_ICS_URL')

wb = Workbook()

ws = wb.active

ws.append(["Date", "Start", "End", "Hours"])
ws.row_dimensions[1].font = Font(bold=True,
                                 color='FFFFFF',
                                 name='Arial')

ws.row_dimensions[1].fill = PatternFill("solid", fgColor="4285F4")


def parse_timedate(ICS_line):
    ICS_line, timedate = ICS_line.split(':')
    date, time = timedate.split('T')
    year = date[:4]
    month = date[4:6]
    day = date[6:8]
    hour = time[:2]
    minute = time[2:4]
    second = time[4:6]
    ICS_line, timezone = ICS_line.split('=')

    return datetime.date(int(year), int(month), int(day)), datetime.time(int(hour), int(minute), int(second))

ICS = requests.get(SCHEDULE_ICS_URL).text.split('\r\n')

lines_nb = len(ICS)
workdays = []
i = 0
while i < lines_nb:
    if ICS[i].startswith('DTSTART;'):
        start_date, start_time = parse_timedate(ICS[i])
        if start_date.month == 7:
            while not ICS[i].startswith('DTEND;'):
                i += 1
            end_date, end_time = parse_timedate(ICS[i])
            workdays.append((start_date, start_time, end_time))
    i += 1

workdays.sort(key=lambda workday: workday[0])

for workday in workdays:
    start_date, start_time, end_time = workday
    ws.append([start_date.strftime('%d/%m/%Y'), start_time, end_time, datetime.datetime.combine(datetime.date.min, end_time) - datetime.datetime.combine(datetime.date.min, start_time)])

wb.save("kafei.xlsx")
