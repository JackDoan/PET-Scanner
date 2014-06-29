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
col_to_filter = pet.column_to_filter(sheet, "Select a column to filter by: ")
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

pet.header()		
print 'Processed ' + str(len(rows_we_want)) + ' rows'
print "\nBreaking things down by state...\n"

state_column = pet.column_to_filter(sheet, "Select the column that denotes the state: ")



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

statebook = open_workbook('test.xls')
statesheet = statebook.sheet_by_index(0)
states = []
#v = value
for v in statesheet.col_values(state_column,1): #!!! make this configurable
	if v in states or len(v) == 0:
		pass
	else:
		states.append(v)
print states

state_rows = []
for state in states:
	templist = [0]
	r = 0
	while r <= statesheet.nrows:
		try:
			if str(statesheet.cell(r, state_column).value) == state:
	 			templist.append(int(r))
	 		r = r+1
	 	except IndexError:
	 		#print 'shit'
	 		break
	state_rows.append(templist)
print state_rows

state_index = 0
for rows in state_rows:

	newbook = Workbook()
	newsheet = newbook.add_sheet(states[state_index])

	newrow = 0
	c = 0

	for r in rows:
		for d in statesheet.row_values(r):
			if c <= statesheet.ncols:
				#print 'row: ' + str(r) + 'col: ' + str(c) + 'data: ' + str(d)
				newsheet.row(newrow).write(c,str(d))
				c = c+1
		newrow = newrow+1
		for d in statesheet.row_values(r):c = 0

	newbook.save(states[state_index] + '.xls')
	state_index = state_index+1