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
# TODO do I need this? import os
from openpyxl import Workbook

# ********** Where we are
# TODO Ask where to find pics
# TODO Find Irfanview - 64, 32 or PortableApps (c or user)

# ********* Run Irfanview
subprocess.run(["iview_64.exe /info=DPI_list_irfanviewOUT.txt", "text=true"])

# ********* Extract data from TXT file
# TODO get each set of info for an image, then move to Excel, place, then return

# ********* Write data to new Excel file
filename = "DPI_list.xlsx"

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "world!"

workbook.save(filename=filename)

# ********* Add formulae to Excel file
# TODO formula for size/DPI coefficient
# TODO formula for 1/4 page
# TODO formula for 1/2 page
# TODO formula for 1/1 page
# TODO color cells if lower than acceptable, Green above 600, yellow 400-600, red for under 400

# ********* Cleanup
# TODO del Irfanview TXT file
