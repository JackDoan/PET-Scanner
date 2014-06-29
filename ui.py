import sys, os
from xlrd import open_workbook
from xlwt import Workbook
import pet

pet.header()
using = pet.file_select()
pet.header()
print "Using " + using
print ""
book = open_workbook('census.xlsx')
sheet = pet.sheet_select(book)
pet.header()
print "\nUsing sheet " + str(book.sheet_names()[0]) + " from " + using + "\n"
print str(sheet.nrows) + ' rows by ' + str(sheet.ncols) + ' columns'
pet.header()
col_to_filter = pet.column_to_filter(sheet)
pet.header()
print 'Building filter parameters...'
param = pet.filter_params(sheet, col_to_filter)
pet.header()
print 'Selecting all where column \"' + str(sheet.cell(0,col_to_filter).value) + '\" is \"' + param + '\"'


#r = row
r = 0
rows_we_want = [0]
while r <= sheet.nrows:
	try:
		if str(sheet.cell(r, col_to_filter).value) == param:
 			rows_we_want.append(int(r))
 		r = r+1
 	except IndexError:
 		break
print '\nProcessed ' + str(len(rows_we_want)) + ' rows'



newbook = Workbook()
newsheet = newbook.add_sheet('Sheet 1')

newrow = 0
c = 0

for r in rows_we_want:
	for d in sheet.row_values(r):
		if c <= sheet.ncols:
			#print 'row: ' + str(r) + 'col: ' + str(c) + 'data: ' + str(d)
			newsheet.row(newrow).write(c,str(d))
			c = c+1
	newrow = newrow+1
	for d in sheet.row_values(r):c = 0

newbook.save('test.xls')


