import sys, os
from xlrd import open_workbook
from xlwt import Workbook

def clear():
	os.system('cls' if os.name == 'nt' else 'clear')

def header():
	os.system('cls' if os.name == 'nt' else 'clear')
	print "Personal Excel Translation Scanner"
	print "=================================="
	print ""

def file_select():
	donegood = 0
	while donegood == 0:
		files_to_use = []
		i = 0
		#shamelessly stolen from StackOverflow: http://stackoverflow.com/questions/11968976/python-list-files-in-the-current-directory-only
		files = [f for f in os.listdir('.') if os.path.isfile(f)]
		for f in files:
			if '~$' in f: pass
			elif '.xls' in f:
				i = i+1 
				print '\t' + str(i) + ") " + str(f)
				files_to_use.append(f)
			elif '.xlsx' in f:
				i = i+1 
				print '\t' + str(i) + ") " + str(f)
				files_to_use.append(f)
		print ""
		files_index_use = input("Choose an Excel file in this directory: ")
		try:
			print files_to_use[int(files_index_use-1)]
			donegood = 1
		except:
			donegood = 0
	return files_to_use[int(files_index_use-1)]

def sheet_select(book):
	print str(book.nsheets) + " sheets found:"
	i = 1
	for s in book.sheet_names():
		print '\t' + str(i) + ") " + s
		i = i+1
	sheet_index = input("Choose a sheet within this workbook: ")
	sheet = book.sheet_by_index(sheet_index-1)
	return sheet

def column_to_filter(sheet):
	i = 1
	if len(sheet.row_values(0)) >= 40:
		os.system("mode con: cols=80 lines=" + str(len(sheet.row_values(0))+9))
		header()
	print 'Determening column headers...'
	for n in sheet.row_values(0):
		print '\t' + str(i) + ") " + n
		i = i+1
	col_to_use = input("Select a column to filter by: ")
	return int(col_to_use)-1

def filter_params(sheet, col_to_use):
	params = []
	#v = value
	for v in sheet.col_values(col_to_use,1): #!!! make this configurable
		if v in params or len(v) == 0:
			pass
		else:
			params.append(v)
	i = 1
	for p in params:
		print '\t' + str(i) + ") " + p
		i = i+1
	param = input("Select a value to exclude by: ")
	return params[int(param)-1]
