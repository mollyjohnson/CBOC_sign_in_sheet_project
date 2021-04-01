#!/usr/bin/python3 

# Author: Molly Johnson
# Date: 3/24/21
# Description: Creates cboc signin excel file for
# checking in clinics for one month. Will adjust 
# for the days of the month, weekends, and federal holidays

####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
####################################################################

# import openpyxl, datetime, and calendar
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import calendar
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# create constants
JB = "Jesse Brown"
CP = "Crown Point"
HE = "Hoffman Est."
LA = "LaSalle"
AU = "Aurora"
JO = "Joliet"
KA = "Kankakee"
OL = "Oak Lawn"
FZ = "Frozen"
TCH = "Tech"
TOA = "Time of Arrival"
MID_DATE = 15
CBOC_COL_WIDTH = 10.86
# cell border values
thin = Side(border_style = "thin", color = "000000")
double = Side(border_style = "double", color = "000000")
thick = Side(border_style = "thick", color = "001C54")
    
####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
####################################################################
def createHeader(ws, startDate, endDate, startDateObj):
    headerBorderLeft = Border(top = thick , left = thick, right = None, bottom = thick) 
    headerBorderRight = Border(top = thick , left = None, right = thick, bottom = thick)  
    headerBorderMid = Border(top = thick, left = None, right = None, bottom = thick)
    
    # font values
    headerFont = Font(name = 'Times New Roman', size = 28, bold = True)    
    
    #set header alignment to center, font to Times New Roman and size to 28
    ws['A1'].alignment = Alignment(horizontal = 'center')
    ws['A1'].font = headerFont
    
    # set border at far left and far right of header merged cells
    ws['A1'].border = headerBorderLeft
    ws['AE1'].border = headerBorderRight
    
    # set border at middle header merged cells
    for row in ws.iter_rows(min_row = 1, max_row = 1, min_col = (startDate + 1), max_col = (endDate - 1)):
        for cell in row:
            cell.border = headerBorderMid 
            
    # create header and merge cells A1 through AE1
    ws['A1'] = ("Month/Year: " + str(calendar.month_name[startDateObj.month]) + " " + str(startDateObj.year))
    data = ws['A1'].value 
    ws.merge_cells('A1:AE1')
    ws['A1'] = data 
    
def getDatetimeObj(startDateStr):
    dateTimeObj = datetime.strptime(startDateStr, "%m-%d-%Y")
    #get day number from date
    print('Day of Month: ', dateTimeObj.day)
    #get year from date
    print('Year: ', dateTimeObj.year)
    #to get name of day (in number) from date
    print('Day of Week (number): ', dateTimeObj.weekday())
    # to get name of day from date
    print('Day of Week (name): ', calendar.day_abbr[dateTimeObj.weekday()])
    # to get name of month from date
    print('Month name: ', calendar.month_name[dateTimeObj.month])
    return dateTimeObj
    
####################################################################
### Function Title: main()
### Arguments:
### Returns:
### Description: 
####################################################################
def main():
    # create workbook and 1st sheet
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "1-15"
    ws1.sheet_properties.tabColor = "1072BA"

    # create 2nd sheet
    ws2 = wb.create_sheet("16-End", 1)
    ws2.sheet_properties.tabColor = "1072BA"
    
    # get start date of the month from user
    startDateStr = input("\nEnter month's start date in the format mm-dd-yyyy: ")
    # Create date object in format mm-dd-yyyy
    startDateObj = getDatetimeObj(startDateStr)
    
    #to iterate to next date/day name
    #print('Next date (num) of week: ', (startDateObj.day + 1))
    #print('Next day of week (name): ', calendar.day_abbr[(startDateObj.weekday()) + 1])
    
    cbocNameBorder = Border(top = thick , left = thick, right = thick, bottom = thick) 
    cbocNameFont = Font(name = 'Times New Roman', size = 10, bold = True)
    ws1['A2'].font = cbocNameFont
    ws1['A2'].border = cbocNameBorder
    ws1.column_dimensions['A'].width = CBOC_COL_WIDTH
    
    # create header for both sheets
    createHeader(ws1, startDateObj.day, MID_DATE, startDateObj)
    createHeader(ws2, (MID_DATE + 1), 31, startDateObj)
    
    # save workbook to excel file and exit
    wb.save('cboc_signin_sheet.xlsx')   

if __name__ == "__main__":
    main()