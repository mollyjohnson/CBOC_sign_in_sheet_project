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

# TEST_START
startDateStr = '2022-04-27'

# Create date object in format yyyy-mm-dd
startDateObj = datetime.strptime(startDateStr, "%Y-%m-%d")

print(startDateObj)
print('Type: ',type(startDateObj))

print('Day of Month', startDateObj.day)

#to get name of day (in number) from date
print('Day of Week (number): ', startDateObj.weekday())

# to get name of day from date
print('Day of Week (name): ', calendar.day_name[startDateObj.weekday()])

# to get name of month from date
print('Month name: ', calendar.month_name[startDateObj.month])
#TEST_END

# save workbook to excel file
wb.save('cboc_signin_sheet.xlsx')   