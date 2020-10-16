#! python 3
# dataToExcel.py - Place all data in the dir in an excel file.
# usage: dataToExcel.py <ind> <path> - pakuje plik z danymi dla zespołu wygenerowanego z <ind> w <path> do .xlsx w <path>

import openpyxl
import os
import argparse
import sys

parser = argparse.ArgumentParser(description='Place all data in the dir in an excel file.')
parser.add_argument('ind',type=int,help='Your index')
parser.add_argument('path', type=str, help='Path to the directory with the files')


args=vars(parser.parse_args())
path = args['path']
ind = args['ind']
name = ''
try:
    name = path.split(os.sep)[-1]
    os.chdir(path)
except:
    print('Could not open the directory')
    sys.exit(1)

filenames = os.listdir()
filenames.sort()

workbook = openpyxl.Workbook()

for filename in filenames:
    if not filename.endswith('.txt'):
        continue
    if filename[-5]==str(ind%5+1): 
        sheet = workbook.active
        sheet.name = filename[:-4]
        with open(filename, 'rt') as file:
            row = 1
            for line in file:
                column = 1
                for item in line.split():
                    sheet.cell(row=row, column=column).value = item
                    column += 1
                row += 1

workbook.save(name+'.xlsx')
input("Finished.")
