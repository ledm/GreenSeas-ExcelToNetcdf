from GreenSeasXLtoNC import GreenSeasXLtoNC

#fn ='xlsx/ArcticandNordicData_short.xlsx'
fni = 'xlsx/AtlanticData_short.xlsx'
fno = 'tmp.nc'


a = GreenSeasXLtoNC(fni,fno, saveShelve=False,saveNC=True)

#other stuff

datasheet = a.datasheet

