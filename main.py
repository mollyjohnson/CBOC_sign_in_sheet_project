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
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
# create lists for both SMP and non SMP CBOC names
smpCBOCs = []
noSMPCBOCs = []

# get CBOCs from file
smpCBOCs, noSMPCBOCs = readFile.getCBOClists(smpCBOCs, noSMPCBOCs)
print ('SMP CBOCs:')
print(smpCBOCs)
print('no SMP CBOCs:')
print(noSMPCBOCs)