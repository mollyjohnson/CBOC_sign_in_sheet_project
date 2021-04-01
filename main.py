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

####################################################################
### Function Title: main()
### Arguments:
### Returns:
### Description: 
####################################################################
def main():
    # import openpyxl, datetime, and calendar
    import openpyxl
    from openpyxl import Workbook
    from datetime import datetime
    import calendar
    from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
    
    # create workbook and 1st sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "1-15"
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
    midDate = 15
    cbocColWidth = 10.86
    
    # get start date of the month from user
    startDateStr = input("\nEnter month's start date in the format mm-dd-yyyy: ")
    # Create date object in format mm-dd-yyyy
    startDateObj = datetime.strptime(startDateStr, "%m-%d-%Y")
    
    #get day number from date
    print('Day of Month: ', startDateObj.day)
    #get year from date
    print('Year: ', startDateObj.year)
    #to get name of day (in number) from date
    print('Day of Week (number): ', startDateObj.weekday())
    # to get name of day from date
    print('Day of Week (name): ', calendar.day_abbr[startDateObj.weekday()])
    # to get name of month from date
    print('Month name: ', calendar.month_name[startDateObj.month])
    
    #to iterate to next date/day name
    #print('Next date (num) of week: ', (startDateObj.day + 1))
    #print('Next day of week (name): ', calendar.day_abbr[(startDateObj.weekday()) + 1])
    
    # cell border values
    thin = Side(border_style = "thin", color = "000000")
    double = Side(border_style = "double", color = "000000")
    thick = Side(border_style = "thick", color = "001C54")
    cbocNameBorder = Border(top = thick , left = thick, right = thick, bottom = thick) 
    headerBorderLeft = Border(top = thick , left = thick, right = None, bottom = thick) 
    headerBorderRight = Border(top = thick , left = None, right = thick, bottom = thick)  
    headerBorderMid = Border(top = thick, left = None, right = None, bottom = thick)
    
    # font values
    headerFont = Font(name = 'Times New Roman', size = 28, bold = True)
    cbocNameFont = Font(name = 'Times New Roman', size = 10, bold = True)
    
    # create day names and dates sub header rows
    ws['A2'].font = cbocNameFont
    ws['A2'].border = cbocNameBorder
    ws.column_dimensions['A'].width = cbocColWidth
    
    # create header and merge cells A1 through AE1
    ws['A1'] = ("Month/Year: " + str(calendar.month_name[startDateObj.month]) + " " + str(startDateObj.year))
    data = ws['A1'].value 
    ws.merge_cells('A1:AE1')
    ws['A1'] = data
    
    # set header alignment to center, font to Times New Roman and size to 28
    ws['A1'].alignment = Alignment(horizontal = 'center')
    ws['A1'].font = headerFont
    
    # set border at far left and far right of header merged cells
    ws['A1'].border = headerBorderLeft
    ws['AE1'].border = headerBorderRight
    
    # set border at middle header merged cells
    for row in ws.iter_rows(min_row = 1, max_row = 1, min_col = 2, max_col = 30):
        for cell in row:
            cell.border = headerBorderMid
    
    # save workbook to excel file and exit
    wb.save('cboc_signin_sheet.xlsx')   

if __name__ == "__main__":
    main()