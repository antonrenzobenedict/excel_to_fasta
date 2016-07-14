#!/usr/bin/python

import sys
from openpyxl import load_workbook

reference = {}
iupac = {"AA":"A", "GG":"G", "CC":"C", "TT":"T", "AG":"R", "TC":"Y", "AT":"W", "GC":"S", "TG":"K", "AC":"M", "--":"N"}

excel_file = load_workbook(sys.argv[1])
multi_fasta = open(sys.argv[2], 'w')

reference_file = open('commonSNPs_3K-6K_pos+chrOnly.txt')

for line in reference_file:
        reference[str(line.split()[2])] = True

sheet = excel_file[excel_file.get_sheet_names()[0]]

row_count = sheet.max_row
col_count = sheet.max_column



for i in range(7,row_count + 1):
        multi_fasta.write(">" + str(sheet.cell(row=i, column=4).value) + " " + str(sheet.cell(row=i, column=6).value) + "\n") 
        for j in range(14,col_count + 1):
                position = str(sheet.cell(row=3, column=j).value)
                if reference.get(position):
                        multi_fasta.write(iupac[str(sheet.cell(row=i, column=j).value)])
        multi_fasta.write("\n")
