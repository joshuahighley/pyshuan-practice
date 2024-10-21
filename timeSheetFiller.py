import datetime as dt

# date at time of running script
today = dt.datetime.now()
weekday_num = int(today.strftime('%w'))
this_week = today.strftime('%U')
this_sun = (today - dt.timedelta(weekday_num)).strftime('%m/%d/%Y')
week_start = "week #" + this_week +  ", beginning Sunday " + this_sun

print("Is this timesheet for " + week_start + "? Enter 'y' if yes, otherwise enter the start date of the week to submit.")
week_input = input()

if week_input == 'y':
    that_sun = this_sun
else:
    days_between = (7*(int(this_week) - int(week_input) + 1)) + weekday_num
    that_sun = (today - dt.timedelta(days = days_between)).strftime('%m/%d/%Y')


# Timesheet data collection
print("Timesheet for: " + that_sun)

#days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
days = ['Sunday', 'Monday']


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

# Testing lists that work with the browser elements of the timesheet
data_0 = [['10/15/2024', ['8'], ['Testing']], ['10/15/2024', ['8'], ['Groovin']]]
data = ['10/15/2024', '5', 'Testing', '10/15/2024', '8', 'Moovin', '10/16/2024', '7.5', 'Groovin', '10/17/2024', '9', 'Approvin']
days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday']

#counters
input_counter = 74      #input id's begin at 74, up to 87 for saturday activity
ts_date_counter = 0     #sunday = 0th index of the week

# this will eventually use service_data list
def ha_filler():
    hours = browser.find_element(By.ID, 'input_' + str(input_counter))
    hours.send_keys(data[1+ts_date_counter])
    input_counter = input_counter + 1

    activity = browser.find_element(By.ID, 'input_' + str(input_counter))
    activity.send_keys(data[2+ts_date_counter])
    input_counter = input_counter + 1

    ts_date_counter = ts_date_counter + 3

# using print to test what was being sent when
# the browser elements increment in the hours portion, unlike in the contact info section
#    print('hours = input_' + str(input_counter))
#    print(data[1 + ts_date_counter])
#    input_counter = input_counter + 1
#    print('activity = input_' + str(input_counter))
#    print(data[2 + ts_date_counter])
#    input_counter = input_counter + 1
#    ts_date_counter = ts_date_counter + 3

# loop to execute the hours 
for d in days:
    ha_filler()
    

print('date counter: ' + str(ts_date_counter))

# Write input data (my hours worked) to a CSV
service_data_df = pd.DataFrame(service_data_m)
ts_csv = pd.read_csv('highleyHours.csv')
pd.DataFrame.to_csv(service_data_df, 'highleyHours.csv', index=False, header=False, mode='a')
