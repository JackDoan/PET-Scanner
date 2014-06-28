import sys
from xlrd import open_workbook
from xlwt import Workbook

book = open_workbook('census.xlsx')
print str(book.nsheets) + " sheets found:"
print '\t' + '\n\t'.join(book.sheet_names())
sheet = book.sheet_by_index(0)
print 'defaulting to sheet 0: ' + str(book.sheet_names()[0])
print str(sheet.nrows) + ' rows by ' + str(sheet.ncols) + ' columns'

print 'Determening column headers...'
print '\t' + '\n\t'.join(sheet.row_values(0))
print 'Choosing to work with column AR: \"' + sheet.cell(0,43).value + '\"'

print 'Building filter parameters...'
params = []
#v = value
for v in sheet.col_values(43,1): #!!! make this configurable
	if v in params or len(v) == 0:
		pass
	else:
		params.append(v)
print '\t' + '\n\t'.join(params)
print 'Choosing to filter by: \"Y\"'

#r = row
r = 0
rows_we_want = [0]
while r <= sheet.nrows:
	try:
		if str(sheet.cell(r, 43).value) == "Y": ###HACK
 			rows_we_want.append(int(r))
 			#print 'ding! ' + str(r)
 		r = r+1
 	except IndexError:
 		break
print 'found ' + str(len(rows_we_want)) + ' rows'



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
