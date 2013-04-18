

from GreenSeasXLtoNC import GreenSeasXLtoNC

fn ='xlsx/ArcticandNordicData.xlsx'
a = GreenSeasXLtoNC(fn)

#other stuff

datasheet = a.datasheet
locator=[]
for h in datasheet.row(11)[0:20]:	locator.append(h.value)
key={}
for n,l in enumerate(locator):
		if l.lower() in [ 'lat', 'latitude']: key['lat']=n
		if l.lower() in [ 'lon','long', 'longitude']: key['lon']=n
		if l in [ 'time','t', 'Date& Time (local)']: key['time']=n
		if l in [ 'Depth of sample [m]']: key['z']=n
		if l in [ 'Depth of Sea [m]',]: key['bathy']=n
		if l in [ 'UTC offset',]: key['tOffset']=n
		if l in ['Institute',]: key['Institute']=n

time = datasheet.col(key['time'])[20:]

header   = [h.value for h in datasheet.row(1)]
units    = [h.value for h in datasheet.row(2)]
locator  = [h.value for h in datasheet.row(11)[0:20]]
metadata = [h.value for h in datasheet.col(10)[0:12]]

datanames=[u'Temperature',]
saveCols=[]
for h,head in enumerate(header):
	for d in datanames: 
		if head.lower().find(d.lower()) > -1:
			print 'FOUND:\t',d, 'in ',head
			saveCols.append(h)

print saveCols
lineTitles = [header[h] for h in saveCols ]
unitTitles = [units[h]  for h in saveCols ]
count =0

lat  = datasheet.col(key['lat'])[20:]
lon  = datasheet.col(key['lon'])[20:]
time = datasheet.col(key['time'])[20:]
depth= datasheet.col(key['z'])[20:]

data={}
for d in saveCols:
	data[d]= datasheet.col(d)[20:]

count = 0
#while count < 10:
for i in xrange(5):
		print count,'t,z,y,x:',time[i].value,depth[i].value,lat[i].value,lon[i].value,
		print 'data:',
		for d in saveCols: print header[d],data[d][i]
		count+=1

#time is a bit odd
#t = xldate_as_tuple(datasheet.col(key['time'])[20:],a.book.datemode)

