#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Author: YALY
# @Date:   2017-01-13 15:11:17
# @Last Modified by:   YALY
# @Last Modified time: 2017-02-04 13:52:59
# Version 0.1

'''
Split the xlsx file.
The version 0.1 only support 2 columns file, you can custom this.
The number of rows per file is 100 by default, you can custom this too.
'''

from openpyxl import Workbook, load_workbook

# Load xlsx file
wb1 = load_workbook('wb.xlsx')

# read worksheet
ws1 = wb1.active

# count work sheet rows
ws1_rows = len(tuple(ws1.rows))

# set the number of rows per file
sep_rows = 100

def insert_data(file_num, rows=100):
	for row in range(1, rows+1):
		ws.cell(row = row, column=1).value, ws.cell(row = row, column=2).value = ws1.cell(row = row + file_num * 100, column = 1).value, ws1.cell(row = row + file_num * 100, column = 2).value


if ws1_rows % 100 == 0:
	file_nums = ws1_rows / sep_rows
else:
	file_nums =  ws1_rows / sep_rows + 1

for file_num in range(0, file_nums):
	wb = Workbook()
	ws = wb.active
	insert_data(file_num, sep_rows)
	filename =  'wb_' + str(file_num) + '.xlsx'
	wb.save(filename=filename)