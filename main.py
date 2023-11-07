#!/usr/bin/env python

__license__ = "GPL"

# ********** Setup
import subprocess
import os
import sys
from openpyxl import Workbook
from distutils.spawn import find_executable

# ********** Where we are
# TODO Ask where to find pics
Pic_Dir = '"c:\\Users\\fAB\\Desktop\\MCB_BMRB\\^photos\\ZD215 Piano\\*.*"'  # Testing REM

# ********* Find and Run Irfanview
IrfanProg_Name = "i_view64.exe"
IrfanProg_Cmd = find_executable(IrfanProg_Name)
if not IrfanProg_Cmd:
    IrfanProg_Cmd = (os.path.expanduser('~')) + "\\PortableApps\\IrfanViewPortable\\App\\IrfanView64\\i_view64.exe"
if not IrfanProg_Cmd:
    IrfanProg_Name = "i_view32.exe"
    IrfanProg_Cmd = find_executable(IrfanProg_Name)
if not IrfanProg_Cmd:
    IrfanProg_Cmd = (os.path.expanduser('~')) + "\\PortableApps\\IrfanViewPortable\\App\\IrfanView\\i_view32.exe"
if not IrfanProg_Cmd:
    sys.exit("Irfanview not installed, please install and run again")

IrfanInfo_Txt = "DPI_list_irfanviewOUT.txt"  # This is the temp file that IrfanView creates

IrfanProg_Cmd = IrfanProg_Cmd + " " + Pic_Dir + " /info=" + IrfanInfo_Txt  # This calls IrfanView and creates TXT file
subprocess.run(IrfanProg_Cmd)

# ********* Extract data from TXT file
# TODO get each set of info for an image, then move to Excel, place, then return

excel_filename = "DPI_list.xlsx"

excel_workbook = Workbook()
excel_sheet = excel_workbook.active
excel_row = 5

text = open(IrfanInfo_Txt)

with open(IrfanInfo_Txt) as IrfanInfo_Data:
    for line in IrfanInfo_Data:
        if not line:
            excel_row = excel_row + 1
            continue
        if not " = " in line: continue
        img_header, img_info = line.split(" = ")


excel_workbook.save(filename=excel_filename)

# find https://www.w3schools.com/python/ref_string_find.asp
# https://stackoverflow.com/questions/29836812/extract-text-after-specific-character

# ********* Write data to new Excel file

# excel_sheet["A1"] = "hello"
# excel_sheet["B1"] = "world!"

# ********* Add formulae to Excel file
# TODO formula for size/DPI coefficient
# TODO formula for 1/4 page
# TODO formula for 1/2 page
# TODO formula for 1/1 page
# TODO color cells if lower than acceptable, Green above 600, yellow 400-600, red for under 400

# ********* Cleanup
# TODO del Irfanview TXT file


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
# from https://stackoverflow.com/questions/89228/how-do-i-execute-a-program-or-call-a-system-command
# use os.path.expandvars to pass variables to subprocess
# On Excel see https://realpython.com/openpyxl-excel-spreadsheets-python/
