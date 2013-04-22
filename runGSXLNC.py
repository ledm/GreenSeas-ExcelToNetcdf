from GreenSeasXLtoNC import GreenSeasXLtoNC




#fn ='xlsx/ArcticandNordicData_short.xlsx'
fni = 'xlsx/AtlanticData.xlsx'
fno = 'AtlanticData_temp_chl.nc'

a = GreenSeasXLtoNC(fni,fno,datanames=['Temperature','Chlorophyll',], saveShelve=False,saveNC=True)


