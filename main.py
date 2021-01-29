#!/usr/bin/python3

# Author: Molly Johnson
# Date: 1/27/21
# Description:
# Accepts an excel file of the previous month's CBOC sign in sheet,
# extracts the year and ending day/date to adjust for the next calendar month,
# and creates a new sheet accordingly. Adjusts for no CBOCs on weekends
# and federal holidays.

# import modules
import openpyxl
from openpyxl import load_workbook
import xlsxwriter

#workbook = load_workbook(filename = "practiceCBOCsignInSheetExcel.xlsx")
#workbook.sheetnames
#sheet = workbook.active
#sheet
#sheet.title

#workbook = xlsxwriter.Workbook('hello.xlsx')
#worksheet = workbook.add_worksheet()
#worksheet.write('A1', 'Hello world')
#workbook.close()