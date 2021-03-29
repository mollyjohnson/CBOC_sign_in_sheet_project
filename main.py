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

# create cboc/core names and other constants
jb = "Jesse Brown"
cp = "Crown Point"
he = "Hoffman Est."
la = "LaSalle"
au = "Aurora"
jo = "Joliet"
ka = "Kankakee"
ol = "Oak Lawn"
fz = "Frozen"
t = "Tech"
toa = "Time of Arrival"

# get start date of the month from user
startDateStr = input("\nEnter month's start date in the format mm-dd-yyyy: ")
# Create date object in format mm-dd-yyyy
startDateObj = datetime.strptime(startDateStr, "%m-%d-%Y")

#print(startDateObj)
#print('Type: ',type(startDateObj))
print('Day of Month: ', startDateObj.day)
#get year from date
print('Year: ', startDateObj.year)
#to get name of day (in number) from date
print('Day of Week (number): ', startDateObj.weekday())
# to get name of day from date
print('Day of Week (name): ', calendar.day_abbr[startDateObj.weekday()])
# to get name of month from date
print('Month name: ', calendar.month_name[startDateObj.month])

# save workbook to excel file
wb.save('cboc_signin_sheet.xlsx')   