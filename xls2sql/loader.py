import os, codecs, sys
import path
try:
	import xlrd
except ImportError:
	print "Trying to Install module: xlrd"
	os.system('python -m pip install xlrd')
finally:
	import xlrd

print('Load file ' + path.file)
try:
	file_xlrd = xlrd.open_workbook(path.file)
except IOError:
	raise
sheet_names = file_xlrd.sheet_names()
sheet = file_xlrd.sheet_by_name(sheet_names[0])

'''function for write the querry'''
def writeQuery(sheet, row, col):
	for k in range(1, sheet.ncols):
		cell = sheet.cell(row, col)
		cell_next = sheet.cell(row, k)
		if cell.value =="" or cell_next.value =="":
			pass
		else:
			if cell.value == cell_next.value:
				pass
			else:
				try:
					print(""+ cell.value+"'" +""+cell_next.value+"';")
					print(""+ cell_next.value+"" +"'"+cell.value+"';")
				except IndexError:
					pass

'''write the querry for all rows'''
def querryForAllRows(sheet):
	for row in range(sheet.nrows):
		for col in range(sheet.ncols):
			writeQuery(sheet, row, col)

'''save querry to file .sql'''
def writeToScript(sheet):	
	try:						
		with codecs.open(path.path + path.filename, 'w',encoding='utf-8') as f:
			sys.stdout = f
			print(querryForAllRows(sheet))
	except IOError:
		print("\ninvalid save file .sql")

'''remove duplicates'''
def uniqueLines(lineslist):
	unique = {}
	result = []
	for item in lineslist:
		if item.strip() in unique: continue
		unique[item.strip()] = 1
		result.append(item)
	return result

def removeDuplicates():
    file1 = codecs.open(path.path + path.filename,'r+',encoding='utf8')
    filelines = file1.readlines()
    filelines.pop()
    file1.close()
    output = codecs.open(path.path + path.filename,"w",encoding='utf8')
    output.writelines(uniqueLines(filelines))
    output.close()