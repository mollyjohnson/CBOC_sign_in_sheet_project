#!/usr/bin/python3 

# Author: Molly Johnson
# Date: 3/24/21
# Description: Creates cboc signin excel file for
# checking in clinics for every month in one year. Will adjust 
# for the days of the month, weekends, and federal holidays

# import openpyxl, datetime, and calendar
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import calendar
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
import os

# create "constants"
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
CBOC_COL_WIDTH = 12.5
HEADER_ROW_HEIGHT = 45
CBOC_ROW_HEIGHT = 20
DATE_ROW_HEIGHT = 22
TECH_TOA_ROW_HEIGHT = 27
CBOC_NAME_AND_FROZEN_ROW_HEIGHT = 21
SPACER_ROW_HEIGHT = 10
NUM_ROWS = 27
HEADER_ROW = 1
HEADER_AND_LABELS_COL = 1
CBOC_COL = 1
CBOC_ROW = 2
DATE_ROW = 3
TECH_TOA_ROW = 4
CBOC_NAME_AND_FROZEN_ROW_START = 5
CBOC_NAME_AND_FROZEN_ROWS = [5,6,8,9,11,12,14,15,17,18,20,21,23,24,26,27]
SPACER_ROWS = [7,10,13,16,19,22,25]
NAMES = [OL,HE,AU,KA,LA,JO,JB,CP] 
WEEKEND_AND_HOL_COL_WIDTH = 2.5
# set "constant" cell border values
thin = Side(border_style = "thin", color = "000000")
double = Side(border_style = "medium", color = "000000")
thick = Side(border_style = "thick", color = "001C54")
cbocNameBorder = Border(top = thick , left = thick, right = thick, bottom = thick) 
cbocNameFont = Font(name = 'Times New Roman', size = 10, bold = True)
dateBorder = Border(left = thick, right = thick, bottom = thick)
bigSpaceBorder = Border(left = thick, right = thick)
spacerBorder = Border(left = thick, right = thick)
cbocNameOnlyBorder = Border(top = double, left = thick, right = thick, bottom = thin)
frozenOnlyBorder = Border(left = thick, right = thick, bottom = double)
bottomRowBorderCBOCName = Border(left = thick, right = thick, bottom = thick)
dateFont = Font(name = 'Calibri', size = 11, bold = True)
techFont = Font(name = 'Calibri', size = 9)
toaFont = Font(name = 'Calibri', size = 5)
dateBorderLeft = Border(top = thick, left = thick, bottom = thick)
dateBorderRight = Border(top = thick, right = thick, bottom = thick)
techInfoBorderLeft = Border(left = thick, right = thin, bottom = double)
techInfoBorderRight = Border(right = thick, bottom = double)
sigBorderTopLeft = Border(top = double, left = thick, right = thin, bottom = thin)
sigBorderBottomLeft = Border(left = thick, right = thin, bottom = double)
sigBorderTopRight = Border(top = double, right = thick, bottom = thin)
sigBorderBottomRight = Border(right = thick, left = thin, bottom = double)
spacerBorderLeft = Border(right = thin, left = thick)
spacerBorderRight = Border(right = thick, left = thin)
weekendAndHolFillColor = PatternFill(fill_type = "solid", start_color = "BFBFBF", end_color = "BFBFBF")

####################################################################
### Function Title: setWeekendStyle()
### Arguments: worksheet, the current col, the current row
### Returns: nothing
### Description: adjusts the width of the columns to narrower, grays out
### all cells in the column, and places x's in the signature/time areas,
### since no CBOC's are received on weekends or federal holidays
####################################################################
def setWeekendAndHolStyle(ws, curCol, curRow):
    endRow = NUM_ROWS
    spacerRows = [7,10,13,16,19,22,25]
    # adjust fill color and col width for all rows in the column
    while(curRow <= endRow):
        ws.cell(row = curRow, column = curCol).fill = weekendAndHolFillColor      
        ws.cell(row = curRow, column = curCol + 1).fill = weekendAndHolFillColor      
        ws.column_dimensions[get_column_letter(curCol)].width = WEEKEND_AND_HOL_COL_WIDTH
        ws.column_dimensions[get_column_letter(curCol + 1)].width = WEEKEND_AND_HOL_COL_WIDTH

        # place "X's" in the signature/time spaces so they can't be marked in. skip spacer rows
        if((curRow >= 5) and (curRow not in spacerRows)):
            ws.cell(row = curRow, column = curCol).value = "X"
            ws.cell(row = curRow, column = curCol + 1).value = "X"
            ws.cell(row = curRow, column = curCol).font = Font(name = 'Calibri', size = 15, bold = True)
            ws.cell(row = curRow, column = curCol + 1).font = Font(name = 'Calibri', size = 15, bold = True)
        curRow += 1

####################################################################
### Function Title: isWeekend()
### Arguments: day name
### Returns: boolean
### Description: checks if the day name passed in is a weekend day
### name. if yes returns true, otherwise returns false
####################################################################
def isWeekend(dayName):
    if(dayName == "SAT" or dayName == "SUN"):
        return True
    return False

####################################################################
### Function Title: createSigBorders()
### Arguments: worksheet, end column
### Returns: nothing
### Description: adjusts the borders for all of the signature cells
####################################################################
def createSigBorders(ws, endCol):
    startCol = 2
    startRow = 5
    endRow = NUM_ROWS
    curCol = startCol
    endCbocNameRow = 26
    endSpacerRow = 25

    while (curCol <= endCol):
        # set cboc name row borders for both cols
        curRow = startRow
        while (curRow <= endCbocNameRow):
            # set left col border
            ws.cell(row = curRow, column = curCol).border = sigBorderTopLeft
            # set right col border
            ws.cell(row = curRow, column = curCol + 1).border = sigBorderTopRight
            curRow += 3

        # set frozen name row borders for both cols
        curRow = 6
        while (curRow <= endRow):
            # set left col border
            ws.cell(row = curRow, column = curCol).border = sigBorderBottomLeft
            # set right col border
            ws.cell(row = curRow, column = curCol + 1).border = sigBorderBottomRight
            if(curRow == endRow):
                # set left col border bottom to thick if last row
                ws.cell(row = curRow, column = curCol).border = Border(right = thin, bottom = thick)
                # set right col border bottom to thick if last row
                ws.cell(row = curRow, column = curCol + 1).border = Border(right = thick, bottom = thick)
            curRow += 3

        # set spacer row borders for both cols
        curRow = 7
        while (curRow <= endSpacerRow):
            # set left col border
            ws.cell(row = curRow, column = curCol).border = spacerBorderLeft
            # set right col border
            ws.cell(row = curRow, column = curCol + 1).border = spacerBorderRight
            curRow += 3    

        # increment cur col by 2 to move onto next date
        curCol += 2

####################################################################
### Function Title: mergeDate()
### Arguments: worksheet, start row, start column
### Returns: nothing
### Description: merges two date info cells
####################################################################
def mergeDateInfo(ws, startRow, startCol):
    data = ws.cell(row = startRow, column = startCol).value
    ws.merge_cells(start_row = startRow, start_column = startCol, end_row = startRow, end_column = startCol + 1)
    ws.cell(row = startRow, column = startCol).value = data


####################################################################
### Function Title: setTechInfo()
### Arguments: worksheet, current column, current row
### Returns: nothing
### Description: adjusts the font/border for the tech and time of
### arrival cells
####################################################################
def setTechInfo(ws, curCol, curRow):
    ws.cell(row = curRow + 2, column = curCol).value = TCH
    ws.cell(row = curRow + 2, column = curCol).font = techFont
    ws.cell(row = curRow + 2, column = curCol).border = techInfoBorderLeft
    

    ws.cell(row = curRow + 2, column = curCol + 1).value = TOA
    ws.cell(row = curRow + 2, column = curCol + 1).font = toaFont
    ws.cell(row = curRow + 2, column = curCol + 1).border = techInfoBorderRight
    ws.cell(row = curRow + 2, column = curCol + 1).alignment = Alignment(wrap_text=True)

####################################################################
### Function Title: setDateInfo()
### Arguments: worksheet, current column, the day of the week number (dayDate),
### the day of the week name, the date number (i.e. date of the month), current row
### Returns: nothing
### Description: sets the font/alignment/border for the day name and date of month
####################################################################
def setDateInfo(ws, curCol, dayDate, dayName, dateNum, curRow):
    # set day name
    ws.cell(row = curRow, column = curCol).value = dayName
    mergeDateInfo(ws, curRow, curCol)
    ws.cell(row = curRow, column = curCol).font = dateFont
    ws.cell(row = curRow, column = curCol).alignment = Alignment(horizontal='center')
    ws.cell(row = curRow, column = curCol).border = dateBorderLeft
    ws.cell(row = curRow, column = curCol + 1).border = dateBorderRight

    # set date num
    ws.cell(row = curRow + 1, column = curCol).value = dateNum
    mergeDateInfo(ws, curRow + 1, curCol)
    ws.cell(row = curRow + 1, column = curCol).font = dateFont
    ws.cell(row = curRow + 1, column = curCol).alignment = Alignment(horizontal='center')
    ws.cell(row = curRow + 1, column = curCol).border = dateBorderLeft
    ws.cell(row = curRow + 1, column = curCol + 1).border = dateBorderRight
    
####################################################################
### Function Title: isHoliday()
### Arguments: holiday dates list, date num (date of the month)
### Returns: boolean
### Description: checks if a specific date is in the previously
### calculated holiday dates list. if yes returns true otherwise false
###################################################################
def isHoliday(holidayDates, dateNum):
    if(dateNum in holidayDates):
        return True
    return False
    
####################################################################
### Function Title: createDateCols()
### Arguments: worksheet, end column, date time object, start date,
### day date (date in the week number), holiday dates list
### Returns: the day date (day in the week number) so can be used for second sheet
### Description: creates/sets the width/borders/font/alignment for each cell
### in each date column. adjusts for if that date is a federal holiday or weekend
###################################################################
def createDateCols(ws, endCol, dateTimeObj, startDate, dayDate, holidayDates):
    #to get name of day (in number) from date
    # to get name of day from date
    dateNum = startDate
    dayName = calendar.day_abbr[dayDate]
    dayName = dayName.upper()

    # make start col and row 2 since dates start after CBOC col and header row
    curRow = 2
    curCol = 2
    #regColWidth = 3.67
    regColWidth = 4.5 
    while (curCol <= (endCol * 2)):
        
        ws.column_dimensions[get_column_letter(curCol)].width = regColWidth 

        # set the day date and name for the date/name cells
        setDateInfo(ws, curCol, dayDate, dayName, dateNum, curRow)
        #set tech info (tech, arrival time)
        setTechInfo(ws, curCol, curRow)

        
        
        # increment to get to second part of each date column
        curCol += 1
        ws.column_dimensions[get_column_letter(curCol)].width = regColWidth

        # check if is weekend or holiday
        if((isWeekend(dayName) == True) or (isHoliday(holidayDates, dateNum) == True)):
            setWeekendAndHolStyle(ws, curCol - 1, curRow)

        # increment to get to first column of new date
        curCol += 1
        # increment the day and date
        dayDate += 1
        dateNum += 1
        #if day number is > 6, i.e. you've reached end of week, start week days over
        if (dayDate > 6):
            dayDate = 0
        dayName = calendar.day_abbr[dayDate]
        dayName = dayName.upper()

    return dayDate
        

####################################################################
### Function Title: createCBOCCOL()
### Arguments: worksheet
### Returns: nothing
### Description: creates the column/font/border for the cboc column
### (which has all cboc names and a row for frozens from each)
###################################################################
def createCBOCCol(ws):
    # create cboc col border and font
    
    # set cboc col width
    ws.column_dimensions[get_column_letter(CBOC_COL)].width = CBOC_COL_WIDTH
    
    # set cboc and date border and font
    ws.cell(row=2, column=CBOC_COL).font = cbocNameFont
    ws.cell(row=2, column=CBOC_COL).border = cbocNameBorder
    ws.cell(row=2, column=CBOC_COL).value = "CBOC/CORE"
    ws.cell(row=3, column=CBOC_COL).border = dateBorder
    ws.cell(row=3, column=CBOC_COL).font = cbocNameFont
    ws.cell(row=3, column=CBOC_COL).value = "Date"
    ws.cell(row=4, column=CBOC_COL).border = bigSpaceBorder

    # put in border/font for cboc name only rows
    i = 5
    j = 0
    while (i <= 26):
        ws.cell(row = i, column = CBOC_COL).font = cbocNameFont
        ws.cell(row = i, column = CBOC_COL).border = cbocNameOnlyBorder
        ws.cell(row = i, column = CBOC_COL).value = NAMES[j]
        i += 3
        j += 1

    # put in frozen rows
    i = 6
    while (i <= 27):
        ws.cell(row = i, column = CBOC_COL).font = cbocNameFont
        ws.cell(row = i, column = CBOC_COL).border = frozenOnlyBorder
        ws.cell(row = i, column = CBOC_COL).value = FZ
        if(i == 27):
            ws.cell(row = i, column = CBOC_COL).border = bottomRowBorderCBOCName
        i += 3

    # put in spacer rows
    i = 7
    while (i <= 25):
        ws.cell(row = i, column = CBOC_COL).border = spacerBorder
        i += 3
    
####################################################################
### Function Title: setRowHeights()
### Arguments: worksheet
### Returns: nothing
### Description: adjusts all row heights depending on which row it is
### (header row, cboc row, date row, tech/time of arrival row, spacer rows,
### cboc name/frozens row)
####################################################################
def setRowHeights(ws):
    ws.row_dimensions[HEADER_ROW].height = HEADER_ROW_HEIGHT
    ws.row_dimensions[CBOC_ROW].height = CBOC_ROW_HEIGHT 
    ws.row_dimensions[DATE_ROW].height = DATE_ROW_HEIGHT
    ws.row_dimensions[TECH_TOA_ROW].height = TECH_TOA_ROW_HEIGHT
    
    for rowNum in CBOC_NAME_AND_FROZEN_ROWS:
        ws.row_dimensions[rowNum].height = CBOC_NAME_AND_FROZEN_ROW_HEIGHT

    for rowNum in SPACER_ROWS:
        ws.row_dimensions[rowNum].height = SPACER_ROW_HEIGHT

####################################################################
### Function Title: validUserInput()
### Arguments: user input (a 4-digit year in string format)
### Returns: boolean
### Description: checks if the user entered a 4-digit int, and if it's
### a year greater than or equal to the current year (2021). returns
### true if valid, false otherwise
####################################################################
def validUserInput(userInput):
    # check that length of user input string is correct
    if(len(userInput) != 4):
        return False
    
    # check that first two chars are digits, mid char
    # is / or -, and last 2 chars are digits.
    i = 0
    while(i < 4):
        if(userInput[i].isdigit() == False):
                return False
        i += 1
    
    # check that year is between 2021 and 2099
    if (int(userInput) < 2021):
        return False
    
    # otherwise met all requirements, return true
    return True

####################################################################
### Function Title: getStartDate()
### Arguments: none
### Returns: start date (in format mm-dd-yyyy) and the year input by
### user (in format yyyy)
### Description: gets input from the user. checks if is a valid year
### in formay yyyy and if is greater than or equal to 2021 (current year).
### if input is invalid, prints error message to the user and keeps asking
### for input until valid input is received. will then use the user's input
### year to create a valid start date that can be used with strptime to get
### a datetime object (which must be in format mm-dd-yyyy). always starts with
### jan 1 of whatever year requested by the user since will create an entire
### year's worth of cboc sheets starting in january
####################################################################
def getStartDate():
    # get start date of the month from user
    i = 0
    while(i == 0):
        userInputYear = input("\nEnter year in the format yyyy: ")
        if(validUserInput(userInputYear) == True):
            i = 1
        else:
            print("Your entry was invalid or a previous year. Enter current year in the format yyyy:")

    # reformat start date input into string for datetime in format 01-01-20yy
    startDate = "01-01-" + userInputYear
    return startDate, userInputYear

####################################################################
### Function Title: createHeader()
### Arguments: worksheet, start row, start column, end row, end column,
### start date object
### Returns: nothing
### Description: creates/merges the header cell with month name/year
####################################################################
def createHeader(ws, startRow, startCol, endRow, endCol, startDateObj):
    # create header border formatting
    headerBorderLeft = Border(top = thick , left = thick, right = None, bottom = thick) 
    headerBorderRight = Border(top = thick , left = None, right = thick, bottom = thick)  
    headerBorderMid = Border(top = thick, left = None, right = None, bottom = thick)
    
    # font values
    headerFont = Font(name = 'Times New Roman', size = 28, bold = True)    
    
    ws.cell(row = startRow, column = startCol).alignment = Alignment(vertical = 'bottom')
    ws.cell(row = startRow, column = startCol).alignment = Alignment(horizontal = 'center')
    ws.cell(row = startRow, column = startCol).font = headerFont
    
    # set border at far left and far right of header merged cells
    ws.cell(row = startRow, column = startCol).border = headerBorderLeft
    ws.cell(row = endRow, column = endCol).border = headerBorderRight
    
    # set border at middle header merged cells
    for row in ws.iter_rows(min_row = startRow, max_row = endRow, min_col = (startCol + 1), max_col = (endCol - 1)):
        for cell in row:
            cell.border = headerBorderMid 
            
    # create header and merge cells A1 through AE1
    ws.cell(row = startRow, column = startCol).value =  ("Month/Year: " + (str(calendar.month_name[startDateObj.month]) + " " + str(startDateObj.year)).upper())
    data = ws.cell(row = startRow, column = startCol).value
    ws.merge_cells(start_row = startRow, start_column = startCol, end_row = endRow, end_column = endCol)
    ws.cell(row = startRow, column = startCol).value = data

####################################################################
### Function Title: getDateTimeObj()
### Arguments: start date (format mm-dd-yyyy)
### Returns: datetime object
### Description: returns the datetime object created using strptime
### from the date provided in mm-dd-yyyy format
####################################################################   
def getDatetimeObj(startDate):
    dateTimeObj = datetime.strptime(startDate, "%m-%d-%Y")
    #print start date string
    #print(startDate)
    #get day number from date
    #print('Day of Month: ', dateTimeObj.day)
    #get year from date
    #print('Year: ', dateTimeObj.year)
    #to get name of day (in number) from date
    #print('Datetime day of Week (number): ', dateTimeObj.weekday())
    # to get name of day from date
    #print('Datetime day of Week (name): ', calendar.day_abbr[dateTimeObj.weekday()])
    # to get name of month from date
    #print('Month name: ', calendar.month_name[dateTimeObj.month])
    return dateTimeObj
    
####################################################################
### Function Title: calcFedHolidays()
### Arguments: datetime object
### Returns: list of federal holiday dates for a given month
### Description: calculates all federal holiday dates for a given month
### (month provided by the datetime object). will adjust for if the
### federal holiday occurs on a saturday or sunday as needed
####################################################################    
def calcFedHolidays(dateTimeObj):
    holidayDates = []
    curDate = 1
    endDate = calendar.monthrange(dateTimeObj.year, dateTimeObj.month)[1]
    monthName = calendar.month_name[dateTimeObj.month]
    monthName = monthName.upper()
    dayDate = dateTimeObj.weekday()
    dayName = calendar.day_abbr[dayDate]
    dayName = dayName.upper()
    mlkMondays = 0
    washBdayMondays = 0
    memorialDayLastMonInMayDate = 0
    laborDayMon = 0
    columbusDayMon = 0
    thanksgivingThurs = 0

    # loop through every day in the month and add holiday dates to the list as appropriate
    while(curDate <= endDate):
        if(monthName == "JANUARY"):
            if(curDate == 1):
                # if new year's occurs on a sat, will be taken care of in dec
                # if new year's occurs on a sun, push it to mon
                if(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                elif(dayName != "SAT"):
                    holidayDates.append(curDate)
            if(dayName == "MON"):
                mlkMondays += 1
                # mlk bday falls on 3rd monday of january
                if(mlkMondays == 3):
                    holidayDates.append(curDate)
        if(monthName == "FEBRUARY"):
            if(dayName == "MON"):
                washBdayMondays += 1
                # washington's bday falls on 3rd monday of the month
                if(washBdayMondays == 3):
                    holidayDates.append(curDate)
        # no fed hols for march or april
        if(monthName == "MAY"):
            # keep replacing last monday in may w the current one,
            # such that the last monday in may will be the last value
            # assigned to the variable (which can then be added later)
            if(dayName == "MON"):
                memorialDayLastMonInMayDate = curDate
        if(monthName == "JUNE"):
            # juneteenth holiday
            if(curDate == 19):
                # if 19th occurs on a sat, add fri to list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 19th holiday occurs on a sun, add mon to list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
        if(monthName == "JULY"):
            # 4th of july holiday
            if(curDate == 4):
                # if 4th holiday occurs on a sat, add fri to list.
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 4th holiday occurs on a sun, add mon to list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
        # no fed hols for august
        if(monthName == "SEPTEMBER"):
            if(dayName == "MON"):
                laborDayMon += 1
                # if is first mon of the month, that's labor day
                if(laborDayMon == 1):
                    holidayDates.append(curDate)
        if(monthName == "OCTOBER"):
            if(dayName == "MON"):
                columbusDayMon += 1
                # if is second mon of month, that's columbus day
                if(columbusDayMon == 2):
                    holidayDates.append(curDate)
        if(monthName == "NOVEMBER"):
            # veteran's day celebrated 11th of nov
            if(curDate == 11):
                # if 11h is a sat, add fri to hols list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 11th is a sun, add mon to hols list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
            if(dayName == "THU"):
                thanksgivingThurs += 1
                # if is 4th thurs of month, that's thanksgiving
                if(thanksgivingThurs == 4):
                    holidayDates.append(curDate)
        if(monthName == "DECEMBER"):
            # if new year's occurs on a sat, add fri dec 31 to hols
            if((curDate == 31) and (dayName == "FRI")):
                holidayDates.append(curDate) 
            # 25th = christmas
            if(curDate == 25):
                # if 25th is on a sat, add fri to hols list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 25th is on a sun, add mon to hols list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)

        curDate += 1
        dayDate += 1
        #if day number is > 6, i.e. you've reached end of week, start week days over
        if (dayDate > 6):
            dayDate = 0
        dayName = calendar.day_abbr[dayDate]
        dayName = dayName.upper()

    # if month is may, add last monday for memorial day
    if(monthName == "MAY"):
        holidayDates.append(memorialDayLastMonInMayDate)

    return holidayDates

####################################################################
### Function Title: main()
### Arguments: none
### Returns: none
### Description: has loop for each month of the given year chosen by
### the user. creates workbook and 2 worksheets per month. calls other
### functions used to get user info, get the datetime object, calculate
### the num of days in the month, calculate the federal holidays, set
### all row and column info for all cells in the cboc signin sheet,
### creates a folder for each year's worksheets (unless the folder
### already exists), and saves all excel spreadsheets using the month
### number/name formatted so that it will have them in correct month
### order in the folder (jan - dec for the given year).
####################################################################
def main():
    currMonth = 1
    endMonth = 12

    currDate, currYear = getStartDate()
    print("...excel spreadsheet creation in progress please wait...")
    
    while (currMonth <= endMonth):
        # create workbook (1st sheet at pos 0 created automatically)
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "1-15"
    
        # create 2nd sheet at pos 1
        ws2 = wb.create_sheet("16-End", 1)
        
        # Create date object in format mm-dd-yyyy from start date string
        dateTimeObj = getDatetimeObj(currDate)
        
        # get end date for the month
        endDate = calendar.monthrange(dateTimeObj.year, dateTimeObj.month)[1]
        
        # calculate holiday dates for the month
        holidayDates = []
        holidayDates = calcFedHolidays(dateTimeObj)
    
        # set row height for all rows
        setRowHeights(ws1)
        setRowHeights(ws2)
    
        # create cboc cell font/border/values for both sheets
        createCBOCCol(ws1)
        createCBOCCol(ws2)
        
        # create rest of cols (date cols) for both sheets
        dayDate = dateTimeObj.weekday()
        # (update day date to be last day date from first sheet before passing to second sheet)
        dayDate = createDateCols(ws1, MID_DATE, dateTimeObj, 1, dayDate, holidayDates)
        createDateCols(ws2, endDate - MID_DATE, dateTimeObj, 16, dayDate, holidayDates)
    
        # create rest of borders for blank areas that will get signatures/initials and times
        createSigBorders(ws1, MID_DATE * 2)
        createSigBorders(ws2, (endDate - MID_DATE) * 2)
        
        # def createHeader(ws, startRow, startCol, endRow, endCol, startDateObj):
        # create header for both sheets
        createHeader(ws1, HEADER_ROW, HEADER_AND_LABELS_COL, HEADER_ROW, (MID_DATE * 2) + 1, dateTimeObj)
        createHeader(ws2, HEADER_ROW, HEADER_AND_LABELS_COL, HEADER_ROW, ((endDate - MID_DATE) * 2) + 1, dateTimeObj)
        
        # if is jan, create folder for all the sheets. otherwise you're already in the folder
        if(currMonth == 1):
            # if directory doesn't exist already, create it. then change to that directory from base directory either way
            if(os.path.isdir(str(currYear) + "_cboc_signin_sheets") == False):
                os.mkdir(str(currYear) + "_cboc_signin_sheets")
            os.chdir(str(currYear) + "_cboc_signin_sheets")

        # save workbook to excel file and exit
        if(currMonth < 10):
            strMonth = "0" + str(currMonth)
        else:
            strMonth = str(currMonth)
        wb.save(strMonth + calendar.month_name[dateTimeObj.month] + "_" + str(dateTimeObj.year) + '_cboc_signin_sheet.xlsx')  

        #increment the month
        currMonth += 1
        currDate = str(currMonth) + "-01-" + str(currYear)
    
    print("\nExcel spreadsheets creation completed please see folder: " + str(currYear) + "_cboc_signin_sheets for your excel files")

if __name__ == "__main__":
    main()