#!/usr/bin/python3 

# Author: Molly Johnson
# Date: 3/24/21
# Description: Creates cboc signin excel file for
# checking in clinics for every month in one year. Will adjust 
# for the days of the month, weekends, and federal holidays

# function comment template
####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################

# import necessary libraries 
from openpyxl import Workbook
from datetime import datetime
import calendar
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import os
import federalHolidayCalculator
import CBOC
import readFile

# create "constants"


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
###################################################################
def main():
    # create lists for both SMP and non SMP CBOC names
    smpCBOCs = []
    noSMPCBOCs = []

    # get CBOCs from file
    smpCBOCs, noSMPCBOCs = readFile.getCBOClists(smpCBOCs, noSMPCBOCs)

    for cboc in noSMPCBOCs:
        print(cboc)

    for cboc in smpCBOCs:
        print(cboc)

if __name__ == "__main__":
    main()