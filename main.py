#!/usr/bin/python3 

# Author: Molly Johnson
# Date: 3/24/21
# Description: Creates cboc signin excel file for
# checking in clinics for one month. Will adjust 
# for the days of the month, weekends, and federal holidays

# openpyxl imports
from openpyxl import Workbook

# create workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "cboc_signin_sheet"
ws.sheet_properties.tabColor = "1072BA"



# save workbook to excel file
wb.save('cboc_signin_sheet.xlsx')    