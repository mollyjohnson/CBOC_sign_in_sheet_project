#!/usr/bin/python3 

# Author: Molly Johnson
# Date: 3/24/21
# Description: Creates cboc signin excel file for
# checking in clinics for one month. Will adjust 
# for the days of the month, weekends, and federal holidays

# import openpyxl, datetime, and calendar
from openpyxl import Workbook
from datetime import datetime
import calendar

# create workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "cboc_signin_sheet"
ws.sheet_properties.tabColor = "1072BA"

# TEST
my_string = '2021-03-27'

# Create date object in given time format yyyy-mm-dd
my_date = datetime.strptime(my_string, "%Y-%m-%d")

print(my_date)
print('Type: ',type(my_date))

print('Day of Month', my_date.day)

#to get name of day (in number) from date
print('Day of Week (number): ', my_date.weekday())

# to get name of day from date
print('Day of Week (name): ', calendar.day_name[my_date.weekday()])

#TEST

# save workbook to excel file
wb.save('cboc_signin_sheet.xlsx')    