from mmap import mmap,ACCESS_READ
from xlrd import open_workbook,cellname
from os.path import exists
from os import makedirs
from shelve import open as shOpen
#import logging


# add metadata, and columns that are all the same value to netcdf attributes
# turn it into a class

# plans take command line file name and variable

# push it to git hub

class GreenSeasXLtoNC:
  def __init__(self, fn, debug=True):
	self.fn = fn
	self.debug=debug
	self.datanames=[u'Temperature',]
	self._run_()


  def _run_(self):
	self._load_()
	self._getData_()
	self._saveShelve_()
	
	
  def _load_(self):
	if self.debug: print 'GreenSeasXLtoNC:\tINFO:\topening:',self.fn
	if not exists(self.fn):
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fn, 'does not exist'
		return
	
	#load excel file
	self.book = open_workbook(self.fn,on_demand=True )
	print 'This workbook contains ',self.book.nsheets,' worksheets.'
	
	#get 'data' sheet
	print 'The worksheets are called: '
	for sheet_name in self.book.sheet_names():print '\t- ',sheet_name

	if 'data' not in self.book.sheet_names():
		print 'GreenSeasXLtoNC:\tERROR:\t', self.fn, 'does not contain a sheet called "data".'
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
	lat  = [h.value for h in self.datasheet.col(key['lat'])[20:]]
	lon  = [h.value for h in self.datasheet.col(key['lon'])[20:]]
	time = [h.value for h in self.datasheet.col(key['time'])[20:]]
	depth= [h.value for h in self.datasheet.col(key['z'])[20:]]
	
	#which columns are being output to netcdf?
	saveCols=[]
	for h,head in enumerate(header):
		for d in self.datanames: 
			if head.lower().find(d.lower()) > -1:
				print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',d, 'in ',head
				saveCols.append(h)
	print 'GreenSeasXLtoNC:\tInfo:\tSaving data from columns:',saveCols
	
	# what is the meta data for those columns:
	lineTitles = {h:header[h] for h in saveCols }
	unitTitles = {h:units[h]  for h in saveCols }
	count =0

	#create data dict.
	data={}
	for d in saveCols:
		data[d]= [h.value for h in self.datasheet.col(d)[20:]]


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
		count+=1
	print 'GreenSeasXLtoNC:\tInfo:\tNumber of entries in each datacolumn:', datacounts
	
	emptyColummns=[]
	for h in sorted(saveCols):
		if datacounts[h] == 0:
			print 'GreenSeasXLtoNC:\tInfo:\tNo data for coumn ',h,lineTitles[h],'[',unitTitles[h],']'
			emptyColummns.append(h)
			

	self.data = data
	self.lineTitles = lineTitles
	self.unitTitles = unitTitles
	self.lat = lat
	self.lon = lon
	self.time = time
	self.depth = depth



  def _saveShelve_(self):
	print 'Saving output Shelve'
	s = shOpen('outShelve.sh')
	s['data'] = self.data
	s['lineTitles'] = self.lineTitles
	s['unitTitles'] = self.unitTitles
	s['lat'] = self.lat
	s['lon'] = self.lon
	s['time'] = self.time
	s['depth'] = self.depth
	s.close()


	
	#next: 
		# start implementing the netcdf output.

	



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
	
	















