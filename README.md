GreenSeas-ExcelToNetcdf
=======================

A package to convert the GreenSeas Excel spreadsheets into Netcdf formats

To run this, you will need the xlrd (excel read) package. It's available on github, but needs an install:

		git clone https://github.com/ledm/xlrd.git
		sudo python setup.py install


You'll also need a bunch of the standard python and netcdf libraries installed.


There are only two files in the repository, one is the main script, GreenSeasXLtoNC.py, and the other is a typical running script, runGSXLNC.py. 

You'll need to edit runGSXLNC.py, and change the file locations to your own files.

Personally, I don't need any of the strings in the excel files, just the numbers, 

Obviously, this is the first draft, so please email me (ledm@pml.ac.uk) if you see a bug, or if you need some help running it, or want to request features. 


