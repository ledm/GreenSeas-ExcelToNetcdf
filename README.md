GreenSeas-ExcelToNetcdf
=======================

A package to convert the GreenSeas Excel spreadsheets into Netcdf formats

To run this, you will also need the xlrd (excel read) package. It's available on github:

		git clone https://github.com/ledm/xlrd.git
		sudo python setup.py install

You'll also need the standard python and netcdf libraries installed.

Note that no GreenSeas excel files are provided with this package.



GreenSeas-ExcelToNetcdf contents:
---------------------------------

There are only three files in the repository, the main script, GreenSeasXLtoNC.py, a typical running script, runGSXLNC.py, and a README. You'll need to edit runGSXLNC.py, and change the file locations. 

You'll need to edit runGSXLNC.py, and change the file locations to your own files.


GreenSeasXLtoNC.py and runGSXLNC.py
-----------------------------------
GreenSeasXLtoNC requires two arguments, an input excel file and an output filename.

The "datanames" option is a list of column names that you want to save.

GreenSeasXLtoNC performs a search through all column headers, so 'Temperature' will save for all non-empty columns with Temperature in the header.
lat,lon,depth, and time and the other metadata are included automatically, so don't need to be requested.
The default dataname is 'Temperature'.

There is also the option to save all the columns of the excell file

		datanames = ['all',] 

although it still ignores columns with no data, or only one datapoint.


plotFromNC.py
-------------
This file uses matplotlib.pyplot to make a time series plot of the data added into the new netcdf.



Caveats:
--------

* GreenSeasXLtoNC is still in beta phase, so make sure you update regularly.

* GreenSeasXLtoNC does not support netcdf string varialbes.

* GreenSeasXLtoNC is not yet flexible for different excel input formats.

* GreenSeasXLtoNC does not yet include much of the GreenSeas metadata.

* GreenSeasXLtoNC is the early draft stage, so please email me (ledm@pml.ac.uk) if you see a bug, or if you need some help running it, or want to request features. 

* Some python packages are OS-dependent, and I've only tested this in Linux. 



