#!/usr/bin/env python

__license__ = 'GPL'
__version__ = '0.1'

# ********** Setup
import subprocess
import os
import sys
import datetime
import time
import requests
import colorama
from distutils.spawn import find_executable
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
import tkinter as tk
from tkinter import filedialog

# ********** General Variables

# generate IDEAL and MIN Image Coefficient
# Formula: Image Quality Coefficient = ((Image width in inches * DPI) * (Image height in inches * DPI)) / 1,000,000
img_coef_page = 28.8  # this is the equivalent of a (8in x 10in @ 600DPI / 1000000) for a 1/1 page IDEAL
img_coef_page_min = 7.2  # this is the equivalent of a (8in x 10in @ 300DPI / 1000000) for a 1/1 page MIN

# Terminal colors

colorama.init()
print(colorama.ansi.clear_screen())

# Filenames
irfan_info_txt = 'DPI_list_irfanviewOUT.txt'
excel_filename = '^DPI_list.xlsx'

root = tk.Tk()
root.withdraw()
pic_dir = filedialog.askdirectory(title="Select Picture Folder")
pic_dir = pic_dir + '/'
if not pic_dir:
    sys.exit('No folder selected, please run again')

# setup PROBLEM colors and bold for Excel
red_fill = PatternFill(start_color='FFFF0000',
                       end_color='FFFF0000',
                       fill_type='solid')
yellow_fill = PatternFill(start_color='FFFFFF00',
                          end_color='FFFFFF00',
                          fill_type='solid')
grey_fill = PatternFill(start_color='FF808080',
                        end_color='FF808080',
                        fill_type='solid')
bold_font = Font(bold=True)
italic_font = Font(italic=True)

# ********** Intro
print()
print(colorama.Fore.BLUE + '                 ******************************************************')
print(colorama.Fore.BLUE + '                 *                Irfanview DPI Spreadsheet           *')
print(colorama.Fore.BLUE + '                 ******************************************************')
print()
print(colorama.Fore.BLUE + 'All the info for this program comes from IrfanView')
print(colorama.Fore.BLUE + 'This is an executable, for the source code please look on GitHub for Irfanview_DPI_list')
print(colorama.Fore.BLUE + 'The program will take a while, please be patient...')

# ********** Check for updates
# Get latest version from web
url = 'https://fabfab1.github.io/Irfanview_DPI_list/i_dpi_list_ver.html'
try:
    resp = requests.get(url)
    resp.raise_for_status()
except requests.exceptions.RequestException as e:
    print(colorama.Fore.RED + 'Warning: Could not get latest version')
    print(e)
    print(colorama.Fore.RED + 'Press Enter to continue without checking version')
    input()

script_ver_ideal = resp.text
script_ver_actual = __version__

if script_ver_actual != script_ver_ideal:
    ver_text_print = 'You have version ' + script_ver_actual + ' and the latest version is ' + script_ver_ideal
    print(colorama.Fore.RED + ver_text_print)
    print(colorama.Fore.RED + 'Press Enter to continue, or update the program')
    input()
else:
    print()
    print(colorama.Fore.BLUE + 'Good, you have the latest version of the program.')

# ********* Find and Run Irfanview
irfan_prog_name = 'i_view64.exe'
irfan_prog_cmd = find_executable(irfan_prog_name)
if not irfan_prog_cmd:
    irfan_prog_cmd = '\\Program Files\\IrfanView\\i_view64.exe'
if not irfan_prog_cmd:
    irfan_prog_cmd = (os.path.expanduser('~')) + '\\PortableApps\\IrfanViewPortable\\App\\IrfanView64\\i_view64.exe'
if not irfan_prog_cmd:
    irfan_prog_name = 'i_view32.exe'
    irfan_prog_cmd = find_executable(irfan_prog_name)
if not irfan_prog_cmd:
    irfan_prog_cmd = (os.path.expanduser('~')) + '\\PortableApps\\IrfanViewPortable\\App\\IrfanView\\i_view32.exe'
if not irfan_prog_cmd:
    irfan_prog_cmd = 'Program Files (x86)\\IrfanView\\i_view32.exe'
if not irfan_prog_cmd:
    sys.exit('Irfanview not installed, please install and run again')

# This calls IrfanView and creates TXT file
irfan_info_txt = os.path.join(pic_dir, irfan_info_txt)
if os.path.exists(irfan_info_txt):  # Delete TXT file if it already exists
    os.remove(irfan_info_txt)
excel_filename = os.path.join(pic_dir, excel_filename)
try:
    if os.path.exists(excel_filename):  # Delete Excel file if it exists
        os.remove(excel_filename)
except PermissionError:
    print("\n ******** Excel File Open! Please close it and run again.")
    time.sleep(10)
    sys.exit()
irfan_prog_cmd = irfan_prog_cmd + ' ' + '"' + pic_dir + '*.*' + '"' + ' /info=' + '"' + irfan_info_txt + '"'
with open(os.devnull, 'w') as devnull:
    subprocess.check_call(irfan_prog_cmd, stderr=devnull)

# ********* Extract data from TXT file

# Setup Excel file
excel_workbook = Workbook()
excel_workbook.remove(excel_workbook['Sheet'])   # Remove default sheet
excel_sheet = excel_workbook.create_sheet("DPI list")
interactive_sheet = excel_workbook.create_sheet("Interactive")
spalten_sheet = excel_workbook.create_sheet("DAI-Zeitschrift Spalten")

# Setup Headers
header_to_col = {
    'File name': 'A',
    'IMG Type & Compression': 'B',
    'Resolution': 'C',
    'Image Dim. (pixels)': 'D',
    'Image Orient.': 'E',
    'Print Size (CM)': 'F',
    'Print Size (IN)': 'G',
    'Web Page': 'H',
    '1/4 Page': 'I',
    '1/2 Page': 'J',
    '1/1 Page': 'K',
    'Beilage 2.5x': 'L'
}

header_to_col2 = {
    'File name': 'A',
    'DPI': 'B',
    'Width CM': 'C',
    'Height CM': 'D',
    'Quality Coefficient': 'E',
    'Quality Enough?': 'F',
}
header_to_col3 = {
    'File name': 'A',
    'IMG Type & Compression': 'B',
    'Resolution': 'C',
    'Image Dim. (pixels)': 'D',
    'Image Orient.': 'E',
    'Print Size (CM)': 'F',
    'Print Size (IN)': 'G',
    '2 Sp. 4.03cm': 'H',
    '3 Sp. 6.28cm': 'I',
    '4 Sp. 8.52cm': 'J',
    '5 Sp. 10.76cm': 'K',
    '6 Sp. 13cm': 'L',
    '8 Sp. 17.5cm': 'M',
    'VolleS. 25.17cm': 'N',
}

# Write Headers for Excel DPI list sheet
excel_sheet["A1"] = 'DPI List -- Data from IrfanView -- ' + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
excel_row = 4
for header in header_to_col:
    col = header_to_col[header]
    excel_sheet[f'{col}{excel_row}'] = header

# Write Headers for Excel Interactive sheet
interactive_sheet["A1"] = 'DPI List -- Data from IrfanView -- ' + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
for header in header_to_col2:
    col = header_to_col2[header]
    interactive_sheet[f'{col}{excel_row}'] = header

# Write Headers for Excel Spalten sheet
spalten_sheet["A1"] = 'DPI List -- Data from IrfanView -- ' + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
for header in header_to_col3:
    col = header_to_col3[header]
    spalten_sheet[f'{col}{excel_row}'] = header
for col in range(1, 15):  # 1-14 corresponds to columns A-N
    spalten_sheet.cell(row=4, column=col).font = italic_font

excel_row = excel_row + 2

# Parse TXT file and write to Excel
numb_images = 0
img_pix_x = img_pix_y = 0
with open(irfan_info_txt, encoding='utf-16-le') as irfan_info_data:
    for line in irfan_info_data:

        if not line.strip():
            excel_row = excel_row + 1
            continue
        if not ' = ' in line:
            continue
        img_header, img_info = line.split(' = ')
        img_header = img_header.strip()
        img_info = img_info.strip()

        if img_header == 'File name':
            excel_sheet['A' + str(excel_row)] = img_info
            interactive_sheet['A' + str(excel_row)] = img_info
            spalten_sheet['A' + str(excel_row)] = img_info
            numb_images = numb_images + 1

        if img_header == 'Directory':
            excel_sheet['A2'] = img_info[:-1]
            interactive_sheet['A2'] = img_info[:-1]
            spalten_sheet['A2'] = img_info[:-1]

        if img_header == 'Compression':
            excel_sheet['B' + str(excel_row)] = img_info
            spalten_sheet['B' + str(excel_row)] = img_info

        if img_header == 'Resolution':
            excel_sheet['C' + str(excel_row)] = img_info
            spalten_sheet['C' + str(excel_row)] = img_info
            img_info = img_info.split(' DPI')[0]
            img_DPI_x, img_DPI_y = img_info.split(' x ')
            img_DPI_x = int(img_DPI_x)
            img_DPI_y = int(img_DPI_y)
            interactive_sheet['B' + str(excel_row)] = img_DPI_x
            if not img_DPI_x == img_DPI_y:
                excel_sheet['C' + str(excel_row)].fill = red_fill
            if img_info in ['0 x 0', '96 x 96']:
                excel_sheet['A' + str(excel_row)].fill = grey_fill
                interactive_sheet['A' + str(excel_row)].fill = grey_fill
                spalten_sheet['A' + str(excel_row)].fill = grey_fill
                numb_images = numb_images - 1
                img_info = img_header = img_pix = img_pix_x = img_pix_y = img_DPI_x = img_DPI_y = img_orient = img_landscape = img_coef = 0
                continue

        if img_header == 'Image dimensions':
            img_pix = img_info.split('  Pixels')[0]
            img_pix_x, img_pix_y = map(int, img_pix.split(' x '))
            excel_sheet['D' + str(excel_row)] = img_pix.strip()
            spalten_sheet['D' + str(excel_row)] = img_pix.strip()

        if img_header == 'Print size':
            img_cm, img_in = img_info.split('; ')
            excel_sheet['F' + str(excel_row)], spalten_sheet['F' + str(excel_row)] = img_cm, img_cm
            excel_sheet['G' + str(excel_row)], spalten_sheet['G' + str(excel_row)] = img_in, img_in
            img_in = img_in.split(' inches')[0]
            img_in_x, img_in_y = img_in.split(' x ')
            img_in_x = float(img_in_x)
            img_in_y = float(img_in_y)
            interactive_sheet['C' + str(excel_row)] = img_in_x
            interactive_sheet['D' + str(excel_row)] = img_in_y
            if img_in_x > img_in_y:
                img_orient = 'Landscape'
                img_landscape = True
            else:
                img_orient = 'Portrait'
                img_landscape = False

        if img_header == 'Color depth':
            color_depth = float(img_info.split()[0].replace(',', '.'))
            if color_depth < 3:
                excel_sheet['B' + str(excel_row)] = spalten_sheet['B' + str(excel_row)] = "BITMAP FILE"
                excel_sheet['B' + str(excel_row)].fill = spalten_sheet['B' + str(excel_row)].fill = grey_fill

            excel_sheet['E' + str(excel_row)] = img_orient
            spalten_sheet['E' + str(excel_row)] = img_orient
            img_coef = (((float(img_in_x) * float(img_DPI_x)) * (float(img_in_y) * float(img_DPI_x)))/1000000)
            interactive_sheet['E' + str(excel_row)] = img_coef
            formula = '=IF(E{0}<$F$3,CONCATENATE("False"),CONCATENATE("True"))'.format(excel_row)
            interactive_sheet['F' + str(excel_row)].value = formula
            # Set conditional format rule to make "False" red/bold
            rule = FormulaRule(formula=['F1="False"'], fill=red_fill, font=Font(bold=True))
            interactive_sheet.conditional_formatting.add('F1:F5000', rule)


            # EXCEL SHEET

            # WEB Page
            excel_sheet['H3'] = (0.10 * img_coef_page)
            if img_coef >= (0.10 * img_coef_page):
                excel_sheet['H' + str(excel_row)] = True
            else:
                excel_sheet['H' + str(excel_row)] = False
                excel_sheet['H' + str(excel_row)].fill = red_fill

            # 1/4 Page
            excel_sheet['I3'] = (0.25 * img_coef_page)
            if img_coef >= (0.25 * img_coef_page):
                excel_sheet['I' + str(excel_row)] = True
            else:
                excel_sheet['I' + str(excel_row)] = False
                excel_sheet['I' + str(excel_row)].fill = red_fill

            # 1/2 Page
            excel_sheet['J3'] = (0.50 * img_coef_page)
            if img_coef >= (0.50 * img_coef_page):
                excel_sheet['J' + str(excel_row)] = True
            else:
                excel_sheet['J' + str(excel_row)] = False
                excel_sheet['J' + str(excel_row)].fill = red_fill

            # 1/1 Page
            excel_sheet['k3'] = (1.0 * img_coef_page)
            if img_coef >= (1.0 * img_coef_page):
                excel_sheet['K' + str(excel_row)] = True
            else:
                excel_sheet['K' + str(excel_row)] = False
                excel_sheet['K' + str(excel_row)].fill = red_fill

            # Beilage Page
            excel_sheet['L3'] = (2.50 * img_coef_page)
            if img_coef >= (2.5 * img_coef_page):
                excel_sheet['L' + str(excel_row)] = True
            else:
                excel_sheet['L' + str(excel_row)] = False
                excel_sheet['L' + str(excel_row)].fill = red_fill

            # SPALTEN SHEET

            def calculate_dpi_spalten(img_pix_x, spalten_x):
                new_width_in = spalten_x / 2.54  # convert cm to inches
                img_DPI_x_spalten = round(img_pix_x / new_width_in)
                return img_DPI_x_spalten
            spalten_ideal_DPI = 600
            spalten_min_DPI = 300
            def set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column):
                if spalten_DPI < spalten_min_DPI:
                    spalten_sheet[curr_column + str(excel_row)].fill = red_fill
                elif spalten_min_DPI <= spalten_DPI < spalten_ideal_DPI:
                    spalten_sheet[curr_column + str(excel_row)].fill = yellow_fill

            # 2 Spalten 4.03cm
            spalten_x = 4.03
            curr_column = 'H'

            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['H' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # 3 Spalten 6.28cm
            spalten_x = 6.28
            curr_column = 'I'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['I' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # 4 Spalten 8.52cm
            spalten_x = 8.52
            curr_column = 'J'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['J' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # 5 Spalten 10.76cm
            spalten_x = 10.76
            curr_column = 'K'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['K' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # 6 Spalten 13cm
            spalten_x = 13
            curr_column = 'L'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['L' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # 8 Spalten 17.5cm
            spalten_x = 17.5
            curr_column = 'M'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['M' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

            # Volle Seite 25.17cm
            spalten_x = 25.17
            curr_column = 'N'
            spalten_DPI = calculate_dpi_spalten(img_pix_x, spalten_x)
            spalten_sheet['N' + str(excel_row)] = spalten_DPI
            set_fill_color(spalten_DPI, spalten_min_DPI, spalten_ideal_DPI, spalten_sheet, excel_row, curr_column)

# Comments to the Excel sheet
comment = """This is the minimum quality @600 DPI for this print size. It was generated using this formula: 
Image Quality Coefficient = ((Image width in inches * DPI) * (Image height in inches * DPI)) / 1,000,000"""
excel_sheet['H3'].comment = Comment(comment, 'FAB')
excel_sheet['I3'].comment = Comment(comment, 'FAB')
excel_sheet['J3'].comment = Comment(comment, 'FAB')
excel_sheet['K3'].comment = Comment(comment, 'FAB')
excel_sheet['L3'].comment = Comment(comment, 'FAB')

# Calculate the number of images
excel_row = excel_row + 1
spalten_sheet['A' + str(excel_row)] = "Total Images: " + str(numb_images)
spalten_sheet['A' + str(excel_row)].font = bold_font

# Now set up the interactive sheet
interactive_sheet['F3'] = 28.8
comment = """This is the minimum quality you are looking for. You can change this number 
and the info below changes. Generate this number using this formula: 
Image Quality Coefficient = ((Image width in inches * DPI) * (Image height in inches * DPI)) / 1,000,000"""
interactive_sheet['F3'].comment = Comment(comment, 'FAB')

# Set column widths and alignment
def adjust_column_width(worksheet):
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

adjust_column_width(excel_sheet)
adjust_column_width(interactive_sheet)
adjust_column_width(spalten_sheet)

excel_sheet.column_dimensions['A'].width, excel_sheet.column_dimensions['B'].width = 21, 21
interactive_sheet.column_dimensions['A'].width, interactive_sheet.column_dimensions['B'].width = 21, 21
spalten_sheet.column_dimensions['A'].width, spalten_sheet.column_dimensions['B'].width = 21, 21
spalten_sheet.merge_cells('H3:N3')
spalten_sheet['H3'] = "Under 300 DPI red, 300-600 DPI yellow. Grey are problem files; check Bitmap DPI by hand."
spalten_sheet['H3'].font = bold_font
spalten_sheet['H3'].alignment = Alignment(horizontal='center')
interactive_sheet.column_dimensions['F'].alignment = Alignment(horizontal='center')

# After creating all the sheets and before saving the workbook
excel_workbook.active = spalten_sheet

# TODO TESTING REMOVE 2 SHEETS
dpi_list_sheet = excel_workbook["DPI list"]
interactive_sheet = excel_workbook["Interactive"]
excel_workbook.remove(dpi_list_sheet)
excel_workbook.remove(interactive_sheet)



excel_workbook.save(filename=excel_filename)
excel_workbook.close()

# ********* Outtro
if os.path.exists(irfan_info_txt):  # Delete TXT file if it already exists
    os.remove(irfan_info_txt)
print()
print(colorama.Fore.BLUE + '*****************************************')
print()
print(colorama.Fore.BLUE + 'Done! Please check the Excel file.')
print(colorama.Fore.BLUE + 'Remember that the image info is only as good as the info from IrfanView...')
print(colorama.Fore.BLUE + 'so if authors "fudge" the image DPI then this program will be wrong!')
print(colorama.Fore.BLUE + 'The Excel file has two sheets, the first is the DPI list and the second is interactive.')
print(colorama.Fore.BLUE + 'hope this was helpful.')
time.sleep(10)

# note for me on how to use pyinstaller:   pyinstaller --onefile --clean Irfanview_DPI_list.py
