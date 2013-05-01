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
	self.saveNC = saveNC	
	self.saveShelve = saveShelve
	self._run_()


  def _run_(self):
	if not self._load_():return
	self._findHeader_()
	self._getData_()

	if self.fno[-3:]=='.nc':	fno = self.fno[:-3]
	if self.fno[-7:]=='.shelve': 	fno = self.fno[:-7]
	
	self.outShelveName = fno+'.shelve'
	self.fno 	   = fno+'.nc'
	
	
	if self.saveNC:	self._saveNC_()
	
	if self.saveShelve:self._saveShelve_()
			
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

  def _findHeader_(self):

	heads = set(['SSTTemp', '5mdepthTemp', '10mdepthTemp', 'Temp', 'MLDTemp', 'MLD', '1pcentlightlevel',
		'TchlorophyllconcChla', 'Sal', 'Wateropticalmeasure1', 'Wateropticalmeasure2', 'Conversionfactor',
		'DepthofDeepChlamax', 'ConcChlamax', 'nanoplankton1015microm', 'nanoplankton1520microm',
		'microplankton2030microm', 'microplankton3050microm', 'microplankton50100microm', 'flagellate_Size220um',
		'flagellate_Size2um', 'Silicoflagellida', 'Choanoflagellida', 'flagellate_Size5um', 'flagellate_Size15um',
		'Dinophyceae_SizeGT20um', 'Dinophyceae_SizeGT20umheterotrophic', 'Strombidiumspp', 'Strobilidiumspp',
		'Ciliatea', 'bacteria', 'cyanobacteria', 'bacteriaheterotrophic', 'picoeukaryotic', 'Synechococcusspp',
		'Prochlorococcusspp', 'Dinophyceae_SizeGT20umSubgroupheterotrophic', 'Strombidiumspp_204',
		'Tintinnidae', 'PhosphatePO4', 'DissNitrateDepthGT1uMconc', 'DissNitrateDepthLT1uMconc', 'NitriteNO2',
		'NitrateNitrite', 'DissAmmonium', 'DissSilicate', 'Silicate', 'Dissbioavailableiron', 'FEIron', 
		'concoxygen', 'saturationofoxygen', 'TCarbonPOC10um', 'TCarbonPOC200um', 'TCarbonPOC5um', 'TCarbonPOC2um',
		'TCarbonPOCGTGFF', 'DissoxygenatMLD', 'Dissoxygenatsurface', 'Dissoxygenat5m', 'Dissoxygenat10m',
		'Nitrogen10um', 'Nitrogen200um', 'Nitrogen5um', 'Nitrogen2um', 'DissInorganicCarbon', 'InorganicCarbon',
		'Alkalinity', 'SPMinorganicsuspendedparticulatematter', 'SPMorganicsuspendedparticulatematter'] )
		
	expectedUnits = [ 'Units','Index', 'CODE', 'umol/l', 'N mol/ L', 'micromol m-3', 'umol/l/d', 'Abundance per L ?', '(mmol O2 m-3)',
	'(umol N /l)', 'mmol m-3', 'f-ratio', 'sum', 'mmol N (or P) m-3 d-1', 'weight units? ', 'total N uptake Chl-normalised',
	'pH', 'mg m-3  < 5um', 'mg m-3  < 20um', 'h', 'number per L?', '%', 'm ', 'g C m-3 d-1', 'M/Sec', 'proportion', '# per ml',
	'UG/L', 'nmol/l', 'other', 'PH', 'Meters', 'dimensionless', 'Number in sample?', 'Kgl', 'Beaufort', 'ML/L',
	'nmol  l-1 h-1', 'mg C m-3 d-1', '(mmol N m-3)', 'mmol  m-2 d-1', 'mmhg', 'Quality score', 'M', 'mg m-3  < 2um',
	'mg m-3', ' # per ml', 'L', 'data quality', 'Weight ?    ', 'COMPASS', 'mg C / m3/h', 'MEQ/L', 'degC', 'mg C / m3/d',
	'Number per L?', 'l/kg', 'Species', '(mg C m-3)', 'umol /l', 'total N uptake nutrient', 'weight units?',
	'(mmol Fe m-3)', '[m]', 'm', 'Number', 'number per L? ', '(mmol Si m-3)', '# per m3', '(umol eq kg)',
	'total N uptake N-normalised', 'mg C m-3', 'sum', 'CODE10', '(mmol P m-3)', 'Quality', '(umol O2/ l)', 'PSU']

			
	expectedloc = ['summerschool','Location', 'Long', 'Lat', 'Long', 'Depth of Sea [m]', 'Depth of sample [m]',
		 'Date& Time (local)', 'Ti me', 'UTC offset', 'measure type1', 'measure type2', 'duplicated (1=Y, 0=N)',
		 'GS Originator / PI', 'Originator / PI', 'Institute', 'Research Group(s) if relevant',
		 'Data collection method(s) description']
	
	expectedMetadata = ['Meta Data', 'Data Title', 'Units', 'Subtitle 1', 'Subtitle 2', 'Subtitle 3',
			'Field description', 'Originator / PI', 'Institute', 'Research Group(s) if relevant',
			'Data collection method(s) description',
			'Explanation/ reference of any conversion factors or aggregation used (if relevant)',]
						
	metas = set(self._getNCvarName_(e) for e in expectedMetadata)
	locs = set(self._getNCvarName_(e) for e in expectedloc)
	uns = set(self._getNCvarName_(e) for e in expectedUnits)
	
	#column search
	metaCols ={}
	for c in xrange(100):
		potMD = [self._getNCvarName_(h.value) for h in self.datasheet.col(c)[0:50]]
		inBoth=metas.intersection(potMD)

		if len(inBoth)> 0:
			print 'Found a meta data column candidate:',c,':', len(inBoth)
			metaCols[c] = len(inBoth)
	
	self.metaC = keywithmaxval(metaCols)
	
	#row search
	headRows, lowRows, unitRows={},{},{}
	metaRow = {}
	for r in xrange(50):
		pot = [self._getNCvarName_(h.value) for h in self.datasheet.row(r)[:]]
		headBoth = heads.intersection(pot)
		locBoth  = locs.intersection(pot)
		unitBoth = uns.intersection(pot)
		metaBoth  = metas.intersection(pot)
		if len(metaBoth)>0: 
			print 'Found the meta row candidate:', r, len(metaBoth),pot[self.metaC], metaBoth
			metaRow[r] = len(metaBoth)
			
		if len(headBoth)>0:
			#print 'Found the Header row candidate:  ',r,':', len(headBoth)
			headRows[r] = len(headBoth)
		if len(locBoth)> 0:
			#print 'Found the location row candidate:',r,':', len(locBoth)	
			lowRows[r] = len(locBoth)
		if len(unitBoth)> 0:
			#print 'Found the unit row candidate:',r,':', len(unitBoth)	
			unitRows[r] = len(unitBoth)
						

	self.locR  = keywithmaxval(lowRows)
	self.headR = keywithmaxval(headRows)
	self.unitR = keywithmaxval(unitRows)
	self.maxMDR = max(metaRow.keys())
	
	print 'best Column for metadata is column: ',self.metaC, 'to row:',self.maxMDR
	print 'best Row for Header is row: ',self.headR
	print 'best row for locations is row: ',self.locR
	print 'best row for units is row: ',self.unitR


	
	    

  def _getData_(self):
	#loading file metadata
	header   = [h.value for h in self.datasheet.row(self.headR)]
	units    = [h.value for h in self.datasheet.row(self.unitR)]

	lastMetaColumn = 20
	locator  = [h.value for h in self.datasheet.row(self.locR)[:lastMetaColumn]]
	
	ckey={}
	for n,l in enumerate(locator):
	    if l in [ 'time','t', 'Date& Time (local)']: ckey['time']=n
	    if l.lower() in [ 'lat', 'latitude']: ckey['lat']=n
	    if l.lower() in [ 'lon','long', 'longitude']: ckey['lon']=n
	    if l in [ 'Depth of sample [m]']: ckey['z']=n
      	    if l in [ 'Depth of Sea [m]',]: ckey['bathy']=n
      	    if l in [ 'UTC offset',]: ckey['tOffset']=n
	    if l in ['Institute',]: ckey['Institute']=n
	
	bad_cells = [xl_cellerror,xl_cellempty,xl_cellblank]
	    
	metadataTitles = {r:h.value for r,h in enumerate(self.datasheet.col(self.metaC)[:self.maxMDR]) if h.ctype not in bad_cells}
	endofHeadRow=max(metadataTitles.keys())



	    
	#create excel coordinates for netcdf.
	colnames = {h: colname(h) for h,head in enumerate(self.datasheet.row(0))} # row number doesn't matter here

	# which columns are we saving?
	saveCols={}
	lineTitles={}
	unitTitles={}
	attributes={}	
	for l,loc in enumerate(locator):	
		if loc in ['', None]:continue
		print 'GreenSeasXLtoNC:\tInfo:\tFOUND:\t',l,'\t',loc, 'in locator'
		saveCols[l] = True
		lineTitles[l]=loc
		unitTitles[l]=''
		if loc.find('[') > 0:
		  unitTitles[l]=loc[loc.find('['):].replace(']','')
	
	if header[5].find('Note')>-1:
		attributes['Note'] = header[5]
		header[5]=''
	
	# flag for saving all columns:
	if 'all' in self.datanames:
		for head in header[lastMetaColumn:]:
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
	

	print 'GreenSeasXLtoNC:\tInfo:\tInterograting columns:',saveCols

	# Meta data for those columns with only one value:
	ncVarName={}
	allNames=[]
	for h in saveCols:
		name = self._getNCvarName_(lineTitles[h])
		#ensure netcdf variable keys are unique:
		if name in allNames:name+='_'+ucToStr(colnames[h])
		allNames.append(name)
		ncVarName[h] = name


	# make an index to link netcdf back to spreadsheet
	index = {}
	for r in xrange(len(self.datasheet.col(saveCols[0])[self.maxMDR:])):
		index[r] = r+self.maxMDR
			
	#create data dictionary
	data={}	
	tunit='seconds since 1900-00-00'
	unitTitles[ckey['time']] = tunit
	for d in saveCols:
		tmpdata= self.datasheet.col(d)[self.maxMDR:]
		arr = []
		if d == ckey['time']: # time
		    for a in tmpdata[:]:
			if a.ctype in bad_cells:
			    arr.append(default_fillvals['i8'])		   
			else:
			    try:  	arr.append(int64(date2num(parse(a.value),units=tunit)))
			    except:	
			    	try: 
			    		arr.append(int(a.value))
			    		print 'GreenSeasXLtoNC:\tWarning: Can not read time effecitvely:',int(a.value)
			    	except: arr.append(default_fillvals['i8'])
		    data[d] = marray(arr)			    
		    continue
		isaString = self._isaString_(lineTitles[d])
		if isaString: #strings
		   for a in tmpdata[:]:
			if a.ctype in bad_cells:
			    arr.append(default_fillvals['S1'])
			else: 
			    try:	arr.append(ucToStr(a.value))
			    except:	arr.append(default_fillvals['S1'])
		else: # data
		   for a in tmpdata[:]:		
			if a.ctype in bad_cells:
			    arr.append(default_fillvals['f4'])
			else:
			    try:   	arr.append(float(a.value))
			    except:	arr.append(default_fillvals['f4'])
		data[d] = marray(arr)
		
	fillvals = default_fillvals.values()							

	# count number of data in each column:
	print 'GreenSeasXLtoNC:\tInfo:\tCount number of data in each column...' # can be slow
	datacounts = {d:0 for d in saveCols}
	for d in saveCols:
		for i in data[d][:]:
			if i in ['', None, ]: continue
			if i in fillvals: continue			
	 		datacounts[d]+=1
	print 'GreenSeasXLtoNC:\tInfo:\tMax number of entries to in a column:', max(datacounts.values())
	
			
		
	# list data columns with no data or only one value	
	removeCol=[]
	for h in saveCols:
		if datacounts[h] == 0:
			print 'GreenSeasXLtoNC:\tInfo:\tNo data for column ',h,lineTitles[h],'[',unitTitles[h],']'
			removeCol.append(h)
			continue	
		col = sorted(data[h])
		if col[0] == col[-1]:
			if col[0] in fillvals:
				print 'GreenSeasXLtoNC:\tInfo:\tIgnoring masked column', h, lineTitles[h],'[',unitTitles[h],']'
				removeCol.append(h)				
				continue
			print 'GreenSeasXLtoNC:\tInfo:\tonly one "data": ',lineTitles[h],'[',unitTitles[h],']','value:', col[0]
			removeCol.append(h)
			attributes[makeStringSafe(ucToStr(lineTitles[h]))] = ucToStr(col[0])

	for r in removeCol:saveCols.remove(r)

	print 'GreenSeasXLtoNC:\tInfo:\tnew file attributes:', attributes	
	
	

		

	print 'GreenSeasXLtoNC:\tInfo:\tFigure out which rows should be saved...'
	saveRows  = {a: False for a in index.keys()} #index.keys() are rows in data. #index.values are rows in excel.
	rowcounts = {a: 0     for a in index.keys()}	
	
	for r in sorted(saveRows.keys()):
		if data[ckey['time']][r] in ['', None,]:
			print 'No time value:',r, data[ckey['time']][r]
			continue
		if data[ckey['time']][r] in fillvals:
			print 'No time value:',r, data[ckey['time']][r]		
			continue	
		for d in saveCols:
			if d<lastMetaColumn:continue
			if data[d][r] in ['', None, ]: continue
			if data[d][r] in fillvals: continue	
			rowcounts[r] += 1			
			saveRows[r] = True
	print 'GreenSeasXLtoNC:\tInfo:\tMaximum number of rows to save: ',max(rowcounts.values())  # 
	
	#rowcounts = {d:0 for d in saveRows.keys()}
	#for r in sorted(rowcounts.keys()):
	#	#if saveRows[r] == False: continue
	#	for d in saveCols:
	#		if d<20:continue
	#		if data[d][r] in ['', None, ]:continue
	#		if data[d][r] in fillvals: continue
	#		rowcounts[r] += 1
	
	
		

	# get data type (ie float, int, etc...):
	# netcdf4 requries some strange names for datatypes: 
	#	ie f8 instead of numpy.float64
	dataTypes={}
	dataIsAString=[]
	for h in saveCols:
		dataTypes[h] = marray(data[h]).dtype
		print 'GreenSeasXLtoNC:\tInfo:\ttype: ',ncVarName[h], h,'\t',dataTypes[h]
		if dataTypes[h] == float64: dataTypes[h] = 'f8'
		elif dataTypes[h] == int32: dataTypes[h] = 'i4'
		elif dataTypes[h] == int64: dataTypes[h] = 'i8'	
		else:
			dataTypes[h] = 'S1'
			dataIsAString.append(h)

		
	print 'GreenSeasXLtoNC:\tInfo:\tCreate MetaData...'	
	#create metadata.
	metadata = {}
	for h in saveCols:
		if h in dataIsAString:continue
		datacol = self.datasheet.col(h)[:]
		colmeta = {metadataTitles[mdk]: datacol[mdk] for mdk in metadataTitles.keys() if metadataTitles[mdk] not in ['', None]}
		
		md='  '
		if len(colmeta.keys())> 20:
			print 'Too many metadata'
			print 'GreenSeasXLtoNC:\tWarning:\tMetadata reading failed. Please Consult original excel file for more info.'			
			metadata[h] = 'Metadata reading failed. Please Consult original excel file for more info.'
			continue
		for mdt,mdc in zip(colmeta.keys(),colmeta.values() ):
			if mdc in ['', None]:continue
			md +=ucToStr(mdt)+':\t'+ucToStr(mdc)+'\n  '
			#print md
		metadata[h] = md
	
		

	# save all info as public variables, so that it can be accessed if netCDF creation fails:
	self.saveCols = saveCols
	self.saveRows = saveRows
	self.rowcounts=rowcounts
	self.ncVarName = ncVarName
	self.dataTypes = dataTypes
	self.dataIsAString=dataIsAString
	self.metadata=metadata
	self.colnames = colnames
	self.data = data
	self.lineTitles = lineTitles
	self.unitTitles = unitTitles
	self.attributes = attributes
	self.index = index



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
	if locName in exceptions.keys(): return ucToStr(exceptions[locName])

	# shorten and make ascii safe.
	a = ucToStr(locName)
	a = a.replace(' ','')
	a = a.replace('Total', 'T')
	a = a.replace('Temperature' , 'Temp')
	a = a.replace('Chlorophyll', 'Chl')
	a = a.replace('Salinity', 'Sal')	 
	a = a.replace('MixedLayerDepth', 'MLD')
	a = a.replace('oncentrationof', 'onc')
	a = a.replace('oncentration', 'onc')	
	a = a.replace('Dissolved', 'Diss')
	a = a.replace('ofwaterbody', '')
	a = makeStringSafe(a)
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
	s['saveCols'] = self.saveCols
	s['rowcounts'] =self.rowcounts
	s['ncVarName'] = self.ncVarName
	s['dataTypes'] = self.dataTypes
	s['dataIsAString'] = self.dataIsAString
	for h in self.saveCols:
		s[self.ncVarName[h]] = array(self.data[h])
	s['lineTitles'] = self.lineTitles
	s['unitTitles'] = self.unitTitles
	s['attributes'] = self.attributes	
	s.close()

  def _saveNC_(self):
	print 'GreenSeasXLtoNC:\tINFO:\tCreating a new dataset:\t', self.fno
	nco = Dataset(self.fno,'w')	
	nco.setncattr('CreatedDate','This netcdf was created on the '+ucToStr(date.today()) +' by '+getuser()+' using GreenSeasXLtoNC.py')
	nco.setncattr('Original File',self.fni)	
	for att in self.attributes.keys():
		print 'GreenSeasXLtoNC:\tInfo:\tAdding Attribute:', att, self.attributes[att]
		nco.setncattr(ucToStr(att), ucToStr(self.attributes[att]))
	
	nco.createDimension('i', None)
	
	nco.createVariable('index', 'i4', 'i',zlib=True,complevel=5)
	
	for v in self.saveCols:
		if v in self.dataIsAString:continue
		print 'GreenSeasXLtoNC:\tInfo:\tCreating var:',v,self.ncVarName[v], self.dataTypes[v]
		nco.createVariable(self.ncVarName[v], self.dataTypes[v], 'i',zlib=True,complevel=5)
	
	nco.variables['index'].long_name =  'Excel Row index'
	for v in self.saveCols:
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var long_name:',v,self.ncVarName[v], self.lineTitles[v]
		nco.variables[self.ncVarName[v]].long_name =  self.lineTitles[v]

	nco.variables['index'].units =  ''		
	for v in self.saveCols:
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding var units:',v,self.ncVarName[v], self.unitTitles[v]	
		nco.variables[self.ncVarName[v]].units =  self.unitTitles[v].replace('[','').replace(']','')

	nco.variables['index'].metadata =  ''
	for v in self.saveCols:
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tAdding meta data:',v#, self.metadata[v]	
		nco.variables[self.ncVarName[v]].metadata =  self.metadata[v]

	nco.variables['index'].xl_column =  '-1'	
	for v in self.saveCols:
		if v in self.dataIsAString:continue	
		print 'GreenSeasXLtoNC:\tInfo:\tAdding excell column name:',v,self.ncVarName[v],':', self.colnames[v]
		nco.variables[self.ncVarName[v]].xl_column =  self.colnames[v]
	
	arr=[]
	for a,val in enumerate(self.index.values()):    
	    if not self.saveRows[a]: continue
	    if self.rowcounts[a] ==0 :continue
	    arr.append(val+1) # accounting for python 0th position is excels = 1st row
	nco.variables['index'][:] = marray(arr)
	
	for v in self.saveCols:
		if v in self.dataIsAString:continue		
		print 'GreenSeasXLtoNC:\tInfo:\tSaving data:',v,self.ncVarName[v],
		arr =  []
		for a,val in enumerate(self.data[v]):
			if not self.saveRows[a]: continue
			if self.rowcounts[a]==0:continue			
			arr.append(val)
		nco.variables[self.ncVarName[v]][:] = marray(arr)
		print len(arr)
				
	
	print 'GreenSeasXLtoNC:\tInfo:\tCreated ',self.fno
	nco.close()

#--------------------------------------------------
# Extra tools.
#--------------------------------------------------
def folder(name):
	if name[-1] != '/':
		name = name+'/'
	if exists(name) is False:
		makedirs(name)
		print 'makedirs ', name
	return name

def keywithmaxval(d):
	""" Figure out the key of the maximum value in a dictionairy"""
	# returns first one
	v,k=list(d.values()),list(d.keys())
	return k[v.index(max(v))]	
	
def ucToStr(d):
	try: return str(d)
	except:return unicode(d).encode('ascii','ignore')


def makeStringSafe(a):
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










