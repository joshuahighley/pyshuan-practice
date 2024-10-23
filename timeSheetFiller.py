import time
from selenium.webdriver.common.by import By
from selenium import webdriver
import datetime as dt
import pandas as pd
import pyinputplus as pyip


# date at time of running script
today = dt.datetime.now()
weekday_num = int(today.strftime('%w'))
this_week = str(int(today.strftime('%U')) + 1)
this_sun = (today - dt.timedelta(weekday_num))  #.strftime('%m/%d/%Y')
week_message = "week #" + this_week +  ", beginning Sunday " + this_sun.strftime('%m/%d/%Y')


# COLLECTING USER TIMESHEET DATA
# user input of week #
print("Is this timesheet for " + week_message + "? Enter 'y' if yes, otherwise enter the start date of the week to submit.")
week_input = input()

if week_input == 'y':
    ts_week = this_sun
else:
    days_between = (7*(int(this_week) - int(week_input))) + weekday_num
    ts_week = (today - dt.timedelta(days = days_between))   #.strftime('%m/%d/%Y')

print("Timesheet for: " + ts_week.strftime('%m/%d/%Y'))


days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
#days = ['Sunday', 'Monday'] #this one's for testing so I don't have to do 7 days
day_count = 0
day_counter = (ts_week + dt.timedelta(days = int(day_count))).strftime('%m/%d/%Y')
service_data = []
service_data_m = []


# loop through days of the week
#   adds date, hours, and activity to list
def activity_collector():
    global day_count
    global day_counter

    print(d + ' hours?')
    new_hours = pyip.inputNum(blank=True)
    if new_hours == '':
        new_hours = 0
    

    elif new_hours == 0:
        print('No hours recorded')

    print('Activities?')
    new_activities = input()
    if new_activities == '':
        new_activities = 'None'
    
    date_counter = (ts_week + dt.timedelta(days = int(day_count))).strftime('%m/%d/%Y')
    print(date_counter)
    day_count = day_count + 1

    service_data.append(day_counter)
    service_data.append(new_hours)
    service_data.append(new_activities)
    service_data_m.append([day_counter, new_hours, new_activities])

for d in days:
    activity_collector()

print(service_data)


# ENTERING COLLECTED DATA
# selenium browser info
browser = webdriver.Chrome()
browser.maximize_window()
browser.get('https://form.jotform.com/volunteercaprw/timesheet')
time.sleep(1)


# my contact info
my_first = browser.find_element(By.ID, 'first_3')
my_first.send_keys('Joshua')
my_last = browser.find_element(By.ID, 'last_3')
my_last.send_keys('Highley')
my_email = browser.find_element(By.ID, 'input_4')
my_email.send_keys('jhighley@caprw.org')

# supervisor's contact info
sup_first = browser.find_element(By.ID, 'first_6')
sup_first.send_keys('Jeremy')
sup_last = browser.find_element(By.ID, 'last_6')
sup_last.send_keys('Taylor')
sup_email = browser.find_element(By.ID, 'input_5')
sup_email.send_keys('jtaylor@caprw.org')

#group affiliation dropdown menu
#my_group = browser.find_element(By.ID, '')

week_field = browser.find_element(By.ID, 'lite_mode_46')
week_field.send_keys(ts_week.strftime('%m/%d/%Y'))


# incrementally fill timesheet from the created list 'service_data'
input_counter = 74
ts_date_counter = 0

def ha_filler():
    global input_counter
    global ts_date_counter

    hours = browser.find_element(By.ID, 'input_' + str(input_counter))
    #volunteer timesheet limits hours input to 8 hours/day :(
    if float(service_data[1+ts_date_counter]) < 8:
        hours.send_keys(str(service_data[1+ts_date_counter]))
    else:
        hours.send_keys('8')
    input_counter = input_counter + 1

    activity = browser.find_element(By.ID, 'input_' + str(input_counter))
    activity.send_keys(str(service_data[2+ts_date_counter]))
    input_counter = input_counter + 1

    ts_date_counter = ts_date_counter + 3

for d in days:
    ha_filler()

print(service_data_m)


# moving csv business down here because errors tend to happen above
# so I'd rather not write on the csv many times if I have to redo a timesheet
# get & write to the csv file
service_data_df = pd.DataFrame(service_data_m)

ts_csv = pd.read_csv('highleyHours.csv')
pd.DataFrame.to_csv(service_data_df, 'highleyHours.csv', index=False, header=False, mode='a')


#review entered data before allowing filler to continue
r = ''
while r != 'y':
    print("Is the entered information correct? Enter 'Y' when finished editing.")
    r = input()
    if r =='y':
        break


print('This browser will close in 15 seconds...')
print('(Just hit submit, still testing the automated submission process.)')

time.sleep(15)


print("Writing data to 'highleyHours.csv'...")
print('Done')
print("Timesheet for: " + ts_week.strftime('%m/%d/%Y') + "submitted.")
#
## END CSV TEST
