# GMAT Automation
This repository contains python scripts to generate GMAT model batch files from an Excel workbook.

The top-level application is modelgen.py.  Configsheet.py is used to parse the Excel workbook.  

Example workbooks are included with names defined in compliance with the interface to configsheet.py.
  
Note how the GMAT sheet is defined. The table in this sheet is generated from an excel Query using the "optimalconfigs" tab as a source. Most of the dynamic model resources are defined on this sheet,

The Mission Params sheet contains named ranges for spacecraft ID ("SpCr") Starting Epoch and Inclination.  Both Starting Epoch and Inclination result in lists when read by configsheet.py.  The Inclination cells are defined according to the convention that the inclination is negative if decreasing and positive if increasing.  The value in the rightmost cells are costate values which are intimately coupled to the inclination change.
 
The Mission name is a named range at the top.

The python scripts use the Inclinations, and Starting Epoch to elaborate instances of the values onthe GMAT sheet.  The total number of mission models that will be created is equal to the number of rows on the GMAT sheet multiplied by the number of starting epochs multiplied by the number of rows of inclination change.
Code must be changed in module fromconfigsheet.py to use other named ranges.

The modelpov.py contains a dictionary that defines the correspondence between table headings on the GMAT tab and the actual GMAT resource names.  
Edit modelpov.py if the columns to be included in the model must change or the GMAT model parameters change names, for instance if the spacecraft name is changed.

Note that GMAT syntax is different than python syntax and thus if modelpov.py is changed, filter code may need also need to be added or changed in order to output a correct model file.  For instance, the epoch string is particularly sensitive to spaces and extraneous characters.

The modelgen.py utilizes three GMAT script templates which are included in the distribution, however these must be copied to the GMAT output path specified in gmat_startup_file.txt.

The environment variable %LOCALUSERAPP% must be defined and contain the path to the top level GMAT executable folder.  
The python scripts will search this path for gmat_startup_file.txt and use the OUTPUT_PATH defined in the most recent file.

See the GMAT Automation Software User Manual.docx for further detail.