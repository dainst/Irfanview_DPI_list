# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# See PyCharm help at https://www.jetbrains.com/help/pycharm/

# subprocess.run(["ls", "-l"])
# from https://stackoverflow.com/questions/89228/how-do-i-execute-a-program-or-call-a-system-command
# use os.path.expandvars to pass variables to subprocess

# On excel see https://realpython.com/openpyxl-excel-spreadsheets-python/

# ********** Setup
import subprocess
import os

# ********** Where we are
# Get directory, ask where to go

# ********* Run Irfanview
subprocess.run(["iview_64.exe /info=DPI_list_irfanviewOUT.txt", "text=true"])

# ********* Extract data from TXT file
# headers

# ********* Write data to new Excel file
from openpyxl import Workbook
filename = "DPI_list.xlsx"

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "world!"

workbook.save(filename=filename)

# ********* Add formulae to Excel file
# formula for size/DPI coefficient
# formula for 1/4 page
# formula for 1/2 page
# formula for 1/1 page
# color cells if lower than acceptable

# ********* Cleanup
#del TXT file