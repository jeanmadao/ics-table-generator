import datetime

def parse_timedate(ics_line):
    ics_line, timedate = ics_line.split(':')
    date, time = timedate.split('T')
    year = date[:4]
    month = date[4:6]
    day = date[6:8]
    hour = time[:2]
    minute = time[2:4]
    second = time[4:6]
    ics_line, timezone = ics_line.split('=')

    return datetime.date(int(year), int(month), int(day)), datetime.time(int(hour), int(minute), int(second))
