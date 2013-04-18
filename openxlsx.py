from mmap import mmap,ACCESS_READ
from xlrd import open_workbook,cellname
from os.path import exists
from shelve import open as shOpen

# add metadata, and columns that are all the same value to netcdf attributes
# turn it into a class

# plans take command line file name and variable

# push it to git hub


fn = 'xlsx/ArcticandNordicData.xlsx'
print 'opening:',fn
if not exists(fn):
	print fn, 'does not exist'
	
book = open_workbook(fn,)#on_demand=True )
print 'This workbook contains ',book.nsheets,' worksheets.'

print 'The worksheets are called: '
for sheet_name in book.sheet_names():print '\t- ',sheet_name

print 'Getting "data" sheet'
datasheet = book.sheet_by_name('data')

print datasheet.name, 'sheet size is ',datasheet.nrows,' x ', datasheet.ncols



locator,header,units,metadata = [],[],[],[]

for h in datasheet.row(11)[0:20]:	locator.append(h.value)
for h in datasheet.row(1):		header.append(h.value)
for h in datasheet.row(2):		units.append(h.value)
for h in datasheet.col(10)[0:12]:	metadata.append(h.value)

outLat,outLon,outZ,outT=[], [],[],[]
data,dq=[],[]

for n in xrange(datasheet.nrows()):
	
	



#the plan is to create a shelve

#for row_i in xrange(datasheet.nrows):
#  for col_i in xrange(1):
#  	print cellname(row_i,col_i),'-', datasheet.cell(row_i,col_i).value

























