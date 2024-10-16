import datetime as dt

# date at time of running script
today = dt.datetime.now()
weekday_num = int(today.strftime('%w'))
this_week = today.strftime('%U')
this_sun = (today - dt.timedelta(weekday_num)).strftime('%m/%d/%Y')
week_start = "week #" + this_week +  ", beginning Sunday " + this_sun

#print("Is this timesheet for " + week_start + "? Enter 'y' if yes, otherwise enter the start date of the week to submit.")
#week_input = input()
week_input = 40

days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
day_count = 0
day_counter = (today + dt.timedelta(days = day_count)).strftime('%m/%d/%Y')
service_data = []

def activity_collector():
    print('Hours served on ' + d + '?')
    new_hours = input()
    if new_hours == '':
        new_hours = 0
    print('Activities?')
    new_activities = input()
    if new_activities == '':
        new_activities = ''
    service_data.append([day_counter, [new_hours], [new_activities]])
    day_count + 1

for d in days:
    activity_collector()

print(service_data)
