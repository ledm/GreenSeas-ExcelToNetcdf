from GreenSeasXLtoNC import folder
from netCDF4 import num2date, Dataset
from matplotlib import pyplot, dates as mdates
from numpy.ma import masked_where,array as marray
from datetime import datetime


def makePlot(times,data,Title,filename):
	print "Making: ", Title
	fig = pyplot.figure()
	ax = fig.add_subplot(111)
	ax.set_xlim(times.min(), times.max())
	#p =  pyplot.plot_date(x=times, y = data) 
	p =  pyplot.plot(times, data,'o',color='g') 	
	years  = mdates.YearLocator(2)
	ax.xaxis.set_major_locator(years)

	pyplot.title( Title)
	print "Saving: ", filename
	pyplot.savefig(filename,dpi=100, bbox_inches='tight')
	pyplot.close()


def main():
	fn = 'AtlanticData_TCSPN.nc'
	print 'opening ',fn
	nc = Dataset (fn,'r')

	time = num2date( nc.variables['Time'][:], nc.variables['Time'].units)
	
	for name in nc.variables.keys():
		if name == 'Time':continue
		
		data = nc.variables[name][:]
		
		tle = nc.variables[name].long_name
		imagefn = folder('images')+name+'.png'
		makePlot(time,data,tle ,imagefn)
		
	nc.close()
	

if __name__=="__main__":
	main() 
	print 'The end.'
