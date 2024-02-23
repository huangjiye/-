import datetime
from icalendar import Calendar, Event
import pandas as pd

def crate_calendar(cal,v_summary,v_description,start_time,end_time):
    event1 = Event()
    event1.add('summary', v_summary)
    event1.add('description', v_description)
    event1.add('dtstart', start_time)
    event1.add('dtend', end_time)
    cal.add_component(event1)



cal = Calendar()

cal.add('version', '2.0')
cal.add('prodid', '-//My calendar generator//')

data = pd.read_excel("/Users/parker/Downloads/3class.xlsx", sheet_name="Sheet1")

for i in data.index:
    if pd.notnull(data.loc[i, 'am']):
        v_summary = ""
        v_description = ""
        if data.loc[i, 'am'].find("考试") == -1:
            v_summary = data.loc[i, 'am'].split("/")[0] + "-" + data.loc[i, 'am'].split("/")[1]
            v_description = data.loc[i, 'am'].split("/")[2]
        else:
            v_summary =data.loc[i, 'am']
            v_description =data.loc[i, 'am']
        start_time = data.loc[i, 'date']
        start_time = start_time.replace(hour=8)
        stop_time = start_time.replace(hour=11, minute=35)
        crate_calendar(cal,v_summary,v_description,start_time,stop_time)

    if pd.notnull(data.loc[i, 'pm']):
        v_summary = ""
        v_description = ""
        if data.loc[i, 'pm'].find("考试") == -1:
            v_summary = data.loc[i, 'pm'].split("/")[0] + "-" + data.loc[i, 'pm'].split("/")[1]
            v_description = data.loc[i, 'pm'].split("/")[2]
        else:
            v_summary =data.loc[i, 'pm']
            v_description =data.loc[i, 'pm']
        start_time = data.loc[i, 'date']
        start_time = start_time.replace(hour=13, minute=30)
        stop_time = start_time.replace(hour=17, minute=5)
        crate_calendar(cal,v_summary,v_description,start_time,stop_time)

    if pd.notnull(data.loc[i, 'pm2']):
        v_summary = ""
        v_description = ""
        if data.loc[i, 'pm2'].find("考试") == -1:
            v_summary = data.loc[i, 'pm2'].split("/")[0] + "-" + data.loc[i, 'pm2'].split("/")[1]
            v_description = data.loc[i, 'pm2'].split("/")[2]
        else:
            v_summary =data.loc[i, 'pm2']
            v_description =data.loc[i, 'pm2']
        start_time = data.loc[i, 'date']
        start_time = start_time.replace(hour=18, minute=00)
        stop_time = start_time.replace(hour=21, minute=10)
        crate_calendar(cal,v_summary,v_description,start_time,stop_time)


#
# # 保存为ICS文件
with open('calendar.ics', 'wb') as f:
    f.write(cal.to_ical())



