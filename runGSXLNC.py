from GreenSeasXLtoNC import GreenSeasXLtoNC




#fni is the input excel spreadsheet 
fni = 'xlsx/AtlanticData.xlsx'

#fno is the output netcdf file you are going to create
fno = 'AtlanticData_temp_chl.nc'

#datanames are the column names that you want to save.
#	This performs a search through all column headers, so 'Temperature' will save for all non-empty columns with Temperature in the header.
#	lat,lon,depth, and time and the other metadata are included automatically, so don't need to be requested.

# saveShelve  and saveNC are boolean flags to save a python shelve or a netcdf.
# (shelves are useful for debugging)
shv=True
snc=True

dns= ['Temperature','Chlorophyll',]
a = GreenSeasXLtoNC(fni,fno,datanames=dns, saveShelve=shv,saveNC=snc)


