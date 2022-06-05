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

# create "constants"
# CBOC file name
CBOC = 'CBOCs.txt'

####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
# Check whether the CBOC needs an SMP check line or not


####################################################################
### Function Title:
### Arguments:
### Returns:
### Description: 
###################################################################
# open CBOC.txt file
file = open(CBOC,'r')

# use readlines to read all lines in the text file and
# return the file contents as a list of strings
lines = []
lines = file.readlines()

# count num CBOCs from the file
numCBOCs = 0

# go through each line and strip the added newline character and increment number of CBOCs
for line in lines:
    print(line.rstrip('\n'))
    numCBOCs += 1

print('\n' + str(numCBOCs))

# close .txt file
file.close()