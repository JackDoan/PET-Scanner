import sys, os
from xlrd import open_workbook
from xlwt import Workbook

def header():
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
				print str(i) + ") " + str(f)
				files_to_use.append(f)
			elif '.xlsx' in f:
				i = i+1 
				print str(i) + ") " + str(f)
				files_to_use.append(f)
		print ""
		files_index_use = input("Choose an Excel file in this directory [1]: ")
		try:
			print files_to_use[int(files_index_use-1)]
			donegood = 1
		except:
			donegood = 0
	return files_to_use[int(files_index_use-1)]
header()
using = file_select()
print ""
print "Using " + using