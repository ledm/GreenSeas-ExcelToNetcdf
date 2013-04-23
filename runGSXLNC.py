

from GreenSeasXLtoNC import GreenSeasXLtoNC

##fni is path to the input excel spreadsheet 
fni = 'xlsx/AtlanticData.xlsx'

##fno is the output netcdf file you are going to create
fno = 'AtlanticData.nc'

##datanames are the column names that you want to save.
##	This performs a search through all column headers, so 'Temperature' will save for all non-empty columns with Temperature in the header.
##	lat,lon,depth, and time and the other metadata are included automatically, so don't need to be requested.



## a minimal selection for fast testing:
dns=['Temperature']
fno='AtlanticData_temp.nc'

## a typical selection of data to convert to netcdf:
#dns= ['Temperature','Chlorophyll','Salinity','Chl-a','Phosphate', 'Nitrate']
#fno = 'AtlanticData_TCSPN.nc'

## Save the entire excell file to netcdf:
## (datanames = 'all' is a flag for saving all the columns of the excell file.
## 	although it still ignores columns with no data.)
#dns = ['all',]
#fno = 'AtlanticData_all.nc'


## saveShelve  and saveNC are boolean flags to save a python shelve or a netcdf.
## The shelve contains all values required to make the netcdf, they are useful for debugging.
## NetCDF is a default format for storing arrays of data.
#shv=False
#snc=False
shv=True
snc=True

a = GreenSeasXLtoNC(fni,fno,datanames=dns, saveShelve=shv,saveNC=snc)


