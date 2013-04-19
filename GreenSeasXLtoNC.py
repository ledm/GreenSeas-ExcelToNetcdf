from mmap import mmap,ACCESS_READ
from xlrd import open_workbook,cellname
from os.path import exists
from os import makedirs
from shelve import open as shOpen
from netCDF4 import Dataset
from datetime import date,datetime
from getpass import getuser
from numpy import array
from numpy.ma import array as marray, masked_where

#try:	, default_fillvals
#except: from netCDF4 import Dataset, _default_fillvals

#import logging


# add metadata, and columns that are all the same value to netcdf attributes
# turn it into a class

# plans take command line file name and variable

# push it to git hub

class GreenSeasXLtoNC:
  def __init__(self, fni, fno, debug=True,saveShelve=False,saveNC=True):
	self.fni = fni
	self.fno = fno
	self.debug=debug
	self.datanames=[u'Temperature',]
	self.saveShelve = saveShelve
	self.saveNC = saveNC	
	self._run_()


  def _run_(self):
	self._load_()
	self._getData_()
	if self.saveShelve:
		if self.fno[-7:]!='.shelve':self.outShelveName =self.fno+'.shelve'
		self._saveShelve_()
	if self.saveNC:
		if self.fno[-3:]!='.nc':self.fno =self.fno+'.nc'	
		self._saveNC_()
	
  def _load_(self):
	if self.debug: print 'GreenSeasXLtoNC:\tINFO:\topening:',self.fni
	if not exists(self.fni):
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fni, 'does not exist'
		return
	
	#load excel file
	self.book = open_workbook(self.fni,on_demand=True )
	print 'This workbook contains ',self.book.nsheets,' worksheets.'
	
	#get 'data' sheet
	print 'The worksheets are called: '
	for sheet_name in self.book.sheet_names():print '\t- ',sheet_name

	if 'data' not in self.book.sheet_names():
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fni, 'does not contain a sheet called "data".'
		return
	print 'Getting "data" sheet'
	self.datasheet = self.book.sheet_by_name('data')

	print self.datasheet.name, 'sheet size is ',self.datasheet.nrows,' x ', self.datasheet.ncols
	return



  def _getData_(self):
	
	#loading file metadata
	header   = [h.value for h in self.datasheet.row(1)]
	units    = [h.value for h in self.datasheet.row(2)]
	locator  = [h.value for h in self.datasheet.row(11)[0:20]]
	metadata = [h.value for h in self.datasheet.col(10)[0:12]]

	print 'GreenSeasXLtoNC:\tInfo:\tlocators:',locator
	
	key ={}
	for n,l in enumerate(locator):
		if l.lower() in [ 'lat', 'latitude']: key['lat']=n
		if l.lower() in [ 'lon','long', 'longitude']: key['lon']=n
		if l in [ 'time','t', 'Date& Time (local)']: key['time']=n
		if l in [ 'Depth of sample [m]']: key['z']=n
		if l in [ 'Depth of Sea [m]',]: key['bathy']=n
		if l in [ 'UTC offset',]: key['tOffset']=n
		if l in ['Institute',]: key['Institute']=n

	# create location t,z,y,x data
	#lat  = [h.value for h in self.datasheet.col(key['lat'])[20:]]
	#lon  = [h.value for h in self.datasheet.col(key['lon'])[20:]]
	#time = [h.value for h in self.datasheet.col(key['time'])[20:]]
	#depth= [h.value for h in self.datasheet.col(key['z'])[20:]]
	
	
	# add location and DQ info to netcdf
	saveCols=[]
	lineTitles={}
	unitTitles={}
	for l,loc in enumerate(locator):	
		if loc in ['', None]:continue
		print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',l,'\t',loc, 'in locator'
		saveCols.append(l)
		lineTitles[l]=loc
		unitTitles[l]=''
		if loc.find('[') > 0:
		  unitTitles[l]=loc[loc.find('['):]
	# add data columns to output to netcdf.
	for h,head in enumerate(header):
		if head == '':continue	
		for d in self.datanames: 
			if head.lower().find(d.lower()) > -1:
				print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',h,'\t',d, 'in ',head
				saveCols.append(h)
				lineTitles[h] = header[h]
				unitTitles[h] = units[h]				

	saveCols = sorted(saveCols)	
	
	print 'GreenSeasXLtoNC:\tInfo:\tSaving data from columns:',saveCols
		

	#create data dictionary
	data={}
	for d in saveCols:
		data[d]= [h.value for h in self.datasheet.col(d)[20:]]
		data[d] = masked_where(data[d]=='', data[d])
		

	# count number of entries in each column:
	datacounts = {h: 0 for h in saveCols}
	for i in xrange(len(lat)):
		a = [data[d][i] for d in saveCols]
		if a.count('') == len(a):continue
		#print '\n\n',count,i,':\tt,z,y,x:',time[i].value,depth[i].value,lat[i].value,lon[i].value,
		#print 'data:',
		for d in saveCols: 
			if data[d][i]:
			 	datacounts[d]+=1
			#print header[d],data[d][i]
	print 'GreenSeasXLtoNC:\tInfo:\tNumber of entries in each datacolumn:', datacounts
	
	# look for data columns with no data
	emptyColummns=[]
	for h in saveCols:
		if datacounts[h] == 0:
			print 'GreenSeasXLtoNC:\tInfo:\tNo data for coumn ',h,lineTitles[h],'[',unitTitles[h],']'
			emptyColummns.append(h)
	print 'GreenSeasXLtoNC:\tInfo:\tEmpty Columns of "data":', emptyColummns
		
	# look for data columns with only one value	
	oneValueInColumn=[]
	for h in saveCols:
		if h in emptyColummns:continue
		col = sorted(data[h])
		#print h,lineTitles[h],'[',unitTitles[h],']','col:', col[0], col[-1]
		if col[0] == col[-1]:
			print 'GreenSeasXLtoNC:\tInfo:\tonly one "data": ',lineTitles[h],'[',unitTitles[h],']','value:', col[0]
			oneValueInColumn.append(h)
			continue
		
	print 'GreenSeasXLtoNC:\tInfo:\tColumns with only one "data":', oneValueInColumn
	
	
	
	# Meta data for those columns:
	ncVarName={}
	allNames=[]
	for h in saveCols:
		if h in emptyColummns:continue
		if h in oneValueInColumn:continue
		name = self._getNCvarName_(lineTitles[h])
		#ensure netcdf variable keys are unique:
		if name in allNames:name+='_'+str(h)
		allNames.append(name)
		ncVarName[h] = name
		

	# figure out data type:
	dataTypes={}
	for h in saveCols:
		if h in emptyColummns:continue
		if h in oneValueInColumn:continue
		col = sorted(data[h])
		if type(col[0]) == type(col[-1]):
			dataTypes[h] = type(col[0])
			
			print 'GreenSeasXLtoNC:\tInfo:\t',h, ncVarName[h], ' is type:',  dataTypes[h]
		else:
			print 'GreenSeasXLtoNC:\tWARNING:\tTWO KINDS OF DATA IN', h, ncVarName[h], [col[0],type(col[0]) ],[col[-1], type(col[-1])]
		
		
			
	self.saveCols = saveCols
	self.emptyColummns = emptyColummns
	self.oneValueInColumn = oneValueInColumn
	self.ncVarName = ncVarName
	self.dataTypes = dataTypes
	self.data = data
	self.lineTitles = lineTitles
	self.unitTitles = unitTitles
	#self.lat = lat
	#self.lon = lon
	#self.time = time
	#self.depth = depth

  def _getNCvarName_(self,locName): 	
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
	if locName in exceptions:
		#print 'GreenSeasXLtoNC:\tInfo:\tgetNCvarName ', locName,'->', exceptions[locName]
		return exceptions[locName]
	else:
		#print 'GreenSeasXLtoNC:\tInfo:\tgetNCvarName ', locName,'->', locName.replace(' ','')
		return locName.replace(' ','')
		
	
  def _saveShelve_(self):
	print 'Saving output Shelve'
	if exists(self.outShelveName):
		print 'GreenSeasXLtoNC:\tWARNING:\tOverwriting previous shelve',self.outShelveName
	s = shOpen(self.outShelveName)
	s['data'] = self.data
	s['lineTitles'] = self.lineTitles
	s['unitTitles'] = self.unitTitles
	s['lat'] = self.lat
	s['lon'] = self.lon
	s['time'] = self.time
	s['depth'] = self.depth
	s.close()

  def _saveNC_(self):
	if self.debug: print 'GreenSeasXLtoNC:\tINFO:\tCreating a new dataset:\t', self.fno
	nco = Dataset(self.fno,'w')	
	nco.setncattr('CreatedDate','This netcdf was created on the '+str(date.today()) +' by '+getuser()+' using GreenSeasXLtoNC.py')
	nco.setncattr('Original File',self.fni)	
	
	nco.createDimension('i', None)
	
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tCreating var:',v,self.ncVarName[v], self.dataTypes[v]
		nco.createVariable(self.ncVarName[v], self.dataTypes[v], 'i',zlib=True,complevel=5)
	
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var long_name:',v,self.ncVarName[v], self.lineTitles[v]
		nco.variables[self.ncVarName[v]].long_name =  self.lineTitles[v]
		
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var units:',v,self.ncVarName[v], self.unitTitles[v]	
		nco.variables[self.ncVarName[v]].units =  self.unitTitles[v]
		
	for v in self.saveCols:
		if v in self.emptyColummns:continue
		if v in self.oneValueInColumn:continue
		print 'GreenSeasXLtoNC:\tInfo:\tSaving data:',v,self.ncVarName[v]
		arr =  array(self.data[v])
		nco.variables[self.ncVarName[v]][:] = arr
				
	

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
	
	















