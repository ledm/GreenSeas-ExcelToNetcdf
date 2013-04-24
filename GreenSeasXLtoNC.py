from mmap import mmap,ACCESS_READ
from xlrd import open_workbook,cellname,colname
from xlrd import XL_CELL_ERROR as xl_cellerror
from xlrd import XL_CELL_EMPTY as xl_cellempty
from xlrd import XL_CELL_BLANK as xl_cellblank
from os.path import exists
from os import makedirs
from shelve import open as shOpen
try:from netCDF4 import Dataset,default_fillvals,date2num
except: from netCDF4 import Dataset, _default_fillvals,date2num
from datetime import date,datetime
from getpass import getuser
from numpy import array, float64, int64,int32
from numpy.ma import array as marray, masked_where
from dateutil.parser import parse

# Things to do:
	# correct strings:
	#	not Really feasible unless you make it a Varialbe Length netcdf, which is a bad idea.
	
	# include data quality flags.
	
	# better logging

	# implement command line control
	
	# implement link back to columns and rows of excel.
	
	# why are so many values fill 
	

class GreenSeasXLtoNC:
  def __init__(self, fni, fno,datanames=[u'Temperature',], saveShelve=False,saveNC=True):
	self.fni = fni
	self.fno = fno
	self.datanames=datanames
	self.saveShelve = saveShelve
	self.saveNC = saveNC	
	self._run_()


  def _run_(self):
	if not self._load_():return
	self._getData_()
	if self.saveShelve:
		if self.fno[-7:]!='.shelve':self.outShelveName =self.fno+'.shelve'
		self._saveShelve_()
	if self.saveNC:
		if self.fno[-3:]!='.nc':self.fno =self.fno+'.nc'	
		self._saveNC_()
	
  def _load_(self):
	print 'GreenSeasXLtoNC:\tINFO:\topening:',self.fni
	if not exists(self.fni):
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fni, 'does not exist'
		return False
	
	#load excel file
	self.book = open_workbook(self.fni,on_demand=True )
	print 'This workbook contains ',self.book.nsheets,' worksheets.'
	
	#get 'data' sheet
	print 'The worksheets are called: '
	for sheet_name in self.book.sheet_names():print '\t- ',sheet_name

	if 'data' not in self.book.sheet_names():
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fni, 'does not contain a sheet called "data".'
		print 'GreenSeasXLtoNC:\tERROR:\tAre you sure that', self.fni, ' is a Greenseas excel file?'		
		return False
	print 'Getting "data" sheet'
	self.datasheet = self.book.sheet_by_name('data')

	print self.datasheet.name, 'sheet size is ',self.datasheet.nrows,' x ', self.datasheet.ncols
	return True



  def _getData_(self):
	#loading file metadata
	header   = [h.value for h in self.datasheet.row(1)]
	units    = [h.value for h in self.datasheet.row(2)]
	locator  = [h.value for h in self.datasheet.row(11)[0:20]]
	metadataTitles = [h.value for h in self.datasheet.col(12)[0:12]]
	time     = [h.value for h in self.datasheet.col(9)[20:]]
	attributes={}
	
	#create excel coordinates for netcdf.
	#index begins at 21 because excel rows start at 1.
	self.index = xrange(21, 21+len(time))
	colnames = {h: colname(h) for h,head in enumerate(self.datasheet.row(1))}


	# add location to netcdf
	saveCols={}
	lineTitles={}
	unitTitles={}
	for l,loc in enumerate(locator):	
		if loc in ['', None]:continue
		print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',l,'\t',loc, 'in locator'
		#if l not in saveCols:
		saveCols[l] = True
		lineTitles[l]=loc
		unitTitles[l]=''
		if loc.find('[') > 0:
		  unitTitles[l]=loc[loc.find('['):].replace(']','')
	
	attributes['Note'] = header[5]
	header[5]=''
	
	# flag for saving all columns:
	if 'all' in self.datanames:
		for head in header[20:]:
			if head == '': continue	
			self.datanames.append(head)
	
	# add data columns titles to output to netcdf.
	for h,head in enumerate(header):
		if head == '':continue	
		if h in saveCols.keys(): continue
		for d in self.datanames:
			if h in saveCols.keys(): continue
			if head.lower().find(d.lower()) > -1:
				print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',h,'\t',d, 'in ',head
				saveCols[h] = True
				lineTitles[h] = header[h]
				unitTitles[h] = units[h]				
	saveCols = sorted(saveCols.keys())	
	
	print 'GreenSeasXLtoNC:\tInfo:\tSaving data from columns:',saveCols
		

	#create data dictionary
	#data={}
	#for d in saveCols:
	#	data[d]= [h.value for h in self.datasheet.col(d)[20:]]
	#	arr = []
	#	isaString = self._isaString_(lineTitles[d])
	#	for a in data[d][:]:
	#	    if isaString:
	#		if a == empty_cell:arr.append(default_fillvals['S1'])
	#		else: arr.append(str(a))
	#	    else:
	#		if a == empty_cell:
	#		    arr.append(default_fillvals['f4'])
	#		else:
	#		    try:   arr.append(float(a))
	#		    except:arr.append(default_fillvals['f4'])
	#	arr = marray(arr)
	#	data[d] = arr
		
	#create data dictionary
	data={}
	bad_cells = [xl_cellerror,xl_cellempty,xl_cellblank]
	for d in saveCols:
		data[d]= self.datasheet.col(d)[20:]
		arr = []
		isaString = self._isaString_(lineTitles[d])
		for a in data[d][:]:
		    if isaString:
			if a.ctype in bad_cells:
				arr.append(default_fillvals['S1'])
			else: arr.append(str(a.value))
		    else:
			if a.ctype in bad_cells:
			    arr.append(default_fillvals['f4'])
			else:
			    try:   arr.append(float(a.value))
			    except:arr.append(default_fillvals['f4'])
		arr = marray(arr)
		data[d] = arr
				
				
	# count number of datapoints in each column:
	print 'GreenSeasXLtoNC:\tInfo:\tCount number of datapoints in each column...' # can be slow
	datacounts = {d:0 for d in saveCols}
	for d in saveCols:
		for i in data[d][:]:
	 		if i: datacounts[d]+=1
	print 'GreenSeasXLtoNC:\tInfo:\tNumber of entries in each datacolumn:', datacounts
	
	
	# count number of datapoints in each row:
	print 'GreenSeasXLtoNC:\tInfo:\tCount number of datapoints in each row...' # can be slow	
	rowcounts = {d:0 for d in xrange(len(time))}
	for r in sorted(rowcounts.keys()):
		for d in saveCols:
			if d<20:continue
			#if rowcounts[r]: break
			if data[d][r] in ['', None, default_fillvals['f4'],]: continue
			rowcounts[r] += 1
			#print r,d,data[d][r],rowcounts[r] 

	
	
	# list data columns with no data
	emptyColummns=[]
	for h in saveCols:
		if datacounts[h] == 0:
			print 'GreenSeasXLtoNC:\tInfo:\tNo data for column ',h,lineTitles[h],'[',unitTitles[h],']'
			emptyColummns.append(h)
	print 'GreenSeasXLtoNC:\tInfo:\tEmpty Columns of "data":', emptyColummns
		
	# list data columns with only one value	
	oneValueInColumn=[]

	for h in saveCols:
		if h in emptyColummns:continue
		col = sorted(data[h])
		if col[0] == col[-1]:
			print 'GreenSeasXLtoNC:\tInfo:\tonly one "data": ',lineTitles[h],'[',unitTitles[h],']','value:', col[0]		
			if col[0] == default_fillvals['f4']:
				'GreenSeasXLtoNC:\tInfo:\tIgnoring masked data' 
				emptyColummns.append(h)
				continue			
			oneValueInColumn.append(h)
			attributes[lineTitles[h]] = col[0]
	print 'GreenSeasXLtoNC:\tInfo:\tColumns with only one non-masked "data":', oneValueInColumn
	print 'GreenSeasXLtoNC:\tInfo:\tnew file attributes:', attributes	
	
	# Meta data for those columns with only one value:
	ncVarName={}
	allNames=[]
	for h in saveCols:
		if h in emptyColummns:continue
		if h in oneValueInColumn:continue
		#ncVarName[h] = self._getNCvarName_(lineTitles[h])+'_'+colnames[h]
		name = self._getNCvarName_(lineTitles[h])
		#ensure netcdf variable keys are unique:
		if name in allNames:name+='_'+colnames[h]
		allNames.append(name)
		ncVarName[h] = name
	
	
	# convert time data into Datetime:
	dt=[]
	timeIsBad = []#{d:0 for d in saveCols}
	tunit='seconds since 1900-00-00'	
	for n,t in enumerate(time):
		try: 
		    dt.append(date2num(parse(t),units=tunit) )
		    timeIsBad.append(False)	
		except:
		    dt.append(-1)
		    timeIsBad.append(True)		    
	data[9] = masked_where(dt==-1,dt)
	unitTitles[9] = tunit
	time = data[9]
		

	# get data type (ie float, int, etc...):
	# netcdf4 requries some strange names for datatypes: 
	#	ie f8 instead of numpy.float64
	dataTypes={}
	dataIsAString=[]
	for h in saveCols:
		if h in emptyColummns:continue
		if h in oneValueInColumn:continue
		dataTypes[h] = data[h].dtype
		print h,dataTypes[h],
		if dataTypes[h] == float64: dataTypes[h] = 'f8'
		elif dataTypes[h] == int32: dataTypes[h] = 'i4'
		elif dataTypes[h] == int64: dataTypes[h] = 'i8'	
		else:
			dataTypes[h] = 'S1'
			dataIsAString.append(h)

	#create metadata.
	metadata = {}
	for h in saveCols:
		if h in emptyColummns:continue
		if h in oneValueInColumn:continue
		if h in dataIsAString:continue
		colmeta = [a.value for a in self.datasheet.col(h)[0:12]]
		md='  '
		for mdt,mdc in zip(metadataTitles,colmeta ):
			if mdc in ['', None]:continue
			md +=str(mdt)+':\t'+str(mdc)+'\n  '
		metadata[h] = md
	
		

	# save all info as public variables, so that it can be accessed if netCDF creation fails:
	self.saveCols = saveCols
	self.emptyColummns = emptyColummns
	self.oneValueInColumn = oneValueInColumn
	self.rowcounts=rowcounts
	self.ncVarName = ncVarName
	self.dataTypes = dataTypes
	self.dataIsAString=dataIsAString
	self.timeIsBad = timeIsBad
	self.metadata=metadata
	self.colnames = colnames
	self.data = data
	self.lineTitles = lineTitles
	self.unitTitles = unitTitles
	self.attributes = attributes
	self.time = time

  def _getNCvarName_(self,locName): 
 	# users are welcome to expand this list to include other elements of the green seas database.
  	exceptions = {	'Depth of Sea [m]': 'Bathymetry', 
		'Depth of sample [m]': 'Depth', 
		'Date& Time (local)': 'Time', 
		'UTC offset': 'UTCoffset', 
		'Explanation/ reference of any conversion factors or aggregation used (if relevant)': 'conversions', 
		'measure type1': 'mType1', 
		'measure type2': 'mType2', 
		'duplicated (1=Y, 0=N)': 'duplicated', 
		'GS Originator / PI': 'gsOriginator', 
		'Originator / PI': 'originator', 
		'Research Group(s) if relevant':'researchGroup',}
	if locName in exceptions.keys(): return exceptions[locName]

	#need to remove a lot of the dodgy characters.
	a = str(locName.replace(' ',''))
	a = a.replace('Total', 'T')
	a = a.replace('Temperature' , 'Temp')
	a = a.replace('Chlorophyll', 'Chl')
	a = a.replace('Salinity', 'Sal')	 
	a = a.replace('MixedLayerDepth', 'MLD')
	a = a.replace('oncentrationof', 'onc')
	a = a.replace('oncentration', 'onc')	
	a = a.replace('Dissolved', 'Diss')
	a = a.replace('ofwaterbody', '')
	a = a.replace('+', '')
	a = a.replace('-', '')						
	a = a.replace('>', 'GT')		
	a = a.replace('<', 'LT')
	a = a.replace('/', '')
	a = a.replace('\\', '')
	a = a.replace('[', '_').replace(']', '')
	a = a.replace('%', 'pcent')	
	a = a.replace(':','').replace('.','')
	return a
	  
	  
  def _isaString_(self,locName): 
	if locName in [ 'Date& Time (local)', 'Explanation/ reference of any conversion factors or aggregation used (if relevant)',
			'measure type1', 'measure type2', u'duplicated (1=Y, 0=N)', u'GS Originator / PI', u'Originator / PI',
			 u'Institute', u'Research Group(s) if relevant']: return True
	return False



  def _saveShelve_(self):
	print 'Saving output Shelve'
	if exists(self.outShelveName):
		print 'GreenSeasXLtoNC:\tWARNING:\tOverwriting previous shelve',self.outShelveName
	s = shOpen(self.outShelveName)
	s['time'] = self.time
	s['saveCols'] = self.saveCols
	s['emptyColummns'] = self.emptyColummns
	s['oneValueInColumn'] = self.oneValueInColumn
	s['rowcounts'] =self.rowcounts
	s['ncVarName'] = self.ncVarName
	s['dataTypes'] = self.dataTypes
	s['dataIsAString'] = self.dataIsAString
	s['timeIsBad'] = self.timeIsBad
	s['data'] = self.data
	s['lineTitles'] = self.lineTitles
	s['unitTitles'] = self.unitTitles
	s['attributes'] = self.attributes	
	s.close()

  def _saveNC_(self):
	print 'GreenSeasXLtoNC:\tINFO:\tCreating a new dataset:\t', self.fno
	nco = Dataset(self.fno,'w')	
	nco.setncattr('CreatedDate','This netcdf was created on the '+str(date.today()) +' by '+getuser()+' using GreenSeasXLtoNC.py')
	nco.setncattr('Original File',self.fni)	
	for att in self.attributes.keys():
		print 'GreenSeasXLtoNC:\tInfo:\tAdding Attribute:', att, self.attributes[att]
		nco.setncattr(att, self.attributes[att])
	
	nco.createDimension('i', None)
	
	nco.createVariable('index', 'i4', 'i',zlib=True,complevel=5)
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue
		print 'GreenSeasXLtoNC:\tInfo:\tCreating var:',v,self.ncVarName[v], self.dataTypes[v]
		nco.createVariable(self.ncVarName[v], self.dataTypes[v], 'i',zlib=True,complevel=5)
	
	nco.variables['index'].long_name =  'Excel Row index'
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var long_name:',v,self.ncVarName[v], self.lineTitles[v]
		nco.variables[self.ncVarName[v]].long_name =  self.lineTitles[v]

	nco.variables['index'].units =  ''		
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var units:',v,self.ncVarName[v], self.unitTitles[v]	
		nco.variables[self.ncVarName[v]].units =  self.unitTitles[v].replace('[','').replace(']','')

	nco.variables['index'].metadata =  ''
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding meta data:',v, self.metadata[v]	
		nco.variables[self.ncVarName[v]].metadata =  self.metadata[v]

	nco.variables['index'].xl_column =  '-1'	
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue	
		print 'GreenSeasXLtoNC:\tInfo:\tAdding excell column name:',v,self.ncVarName[v],':', self.colnames[v]
		nco.variables[self.ncVarName[v]].xl_column =  self.colnames[v]
	
	arr=[]
	for a,val in enumerate(self.index):
	    if self.rowcounts[a]== 0:continue
	    if self.timeIsBad[a]: continue
	    arr.append(val)
	nco.variables['index'][:] = marray(arr)
	
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tSaving data:',v,self.ncVarName[v]
		arr =  []
		for a,val in enumerate(self.data[v]):
			if self.rowcounts[a]== 0:continue
			if self.timeIsBad[a]: continue
			arr.append(val)
		nco.variables[self.ncVarName[v]][:] = marray(arr)
				
	

	nco.close()

#--------------------------------------------------
# folder, lastWord
#--------------------------------------------------
def folder(name):
	if name[-1] != '/':
		name = name+'/'
	if exists(name) is False:
		makedirs(name)
		print 'makedirs ', name
	return name
	
def lastWord(a, separator='/',debug=False):
	#returns the final word of a string.
	while a[-1] == separator: a = a[:-1]
	ncount, count = 0,1
	while ncount != count:
		ncount = count
		count = a[count:].find('/')+1+count
		#if ncount == count : break
	if debug: print 'final word:', a[count:]
	return a[count:]
	
	















