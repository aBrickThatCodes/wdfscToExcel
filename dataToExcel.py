#! python 3
# dataToExcel.py - Place all data in the dir in an excel file.
# usage: dataToExcel.py <path> - pakuje wszystkie pliki tekstowe w <path> do .xlsx w <path> 

import openpyxl
import os
import argparse
import sys

parser=argparse.ArgumentParser(description='Place all data in the dir in an excel file.')
parser.add_argument('path',type=str,help='Path to the directory with the files')

path=vars(parser.parse_args())['path']
name=''
try:
    name=path.split(os.sep)[-1]
    os.chdir(path)
except:
    print('Could not open the directory')
    sys.exit(1)


filenames=os.listdir()
filenames.sort()

workbook=openpyxl.Workbook()


for filename in filenames:
    if not filename.endswith('.txt'):
        continue
    if not filename[:-4] in workbook.sheetnames:
        workbook.create_sheet(title=filename[:-4])
    sheet=workbook[filename[:-4]]
    with open(filename,'rt') as file:
        row=1
        for line in file:
            column=1
            for item in line.split():
                sheet.cell(row=row,column=column).value=item
                column+=1
            row+=1

del workbook['Sheet']
workbook.save(name+'.xlsx')
input("Finished.")