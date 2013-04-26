

from GreenSeasXLtoNC import GreenSeasXLtoNC,folder

##fni is path to the input excel spreadsheet 
#fni = 'xlsx/AtlanticData.xlsx'
fni = 'xlsx/AtlanticData_short.xlsx'

##fno is the output netcdf file you are going to create
#fno = 'AtlanticData.nc'
fno=folder('output')+'AtlanticData_temp.nc'
#fno = 'AtlanticData_TCSPN.nc'
#fno = 'AtlanticData_all.nc'

##datanames are the column names that you want to save.
##	This performs a search through all column headers, so 'Temperature' will save for all non-empty columns with Temperature in the header.
##	lat,lon,depth, and time and the other metadata are included automatically, so don't need to be requested.
## 	datanames = ['all',] is a flag for saving all the columns of the excell file.
## 	although it still ignores columns with no data.)
dns=['Temperature']
#dns= ['Temperature','Chlorophyll','Salinity','Chl-a','Phosphate', 'Nitrate']
#dns = ['all',]










## saveShelve  and saveNC are boolean flags to save a python shelve or a netcdf.
## The shelve contains all values required to make the netcdf, they are useful for debugging.
## NetCDF is a default format for storing arrays of data.
#shv=False
#snc=False
shv=True
snc=True

a = GreenSeasXLtoNC(fni,fno,datanames=dns, saveShelve=shv,saveNC=snc)


