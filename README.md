# GMAT Automation
This repository contains python scripts to generate GMAT model batch files from an Excel workbook.
Top level application is modelgen.py.

Two example workbooks are included.  
Note how the GMAT tab is defined.  
The table is generated from a Query using the "optimalconfigs" tab as a source.

Also note how the Mission Params tab contains named ranges for Starting Epoch and Inclination.  
The Mission name is a named range at the top.

The baseline scripts use the Inclinations, Starting Epoch and Mission named ranges.  
Code must be changed in module fromconfigsheet.py to use other named ranges.

The modelpov.py contains a dictionary that defines the correspondence between table headings on the GMAT tab and the actual GMAT resource names.  
Edit modelpov.py if the columns to be included in the model change.

Note that GMAT syntax is different than python syntax and thus if modelpov.py is changed, filter code in modelgen.py may need to be added or changed in order to output a correct model file.

Three GMAT scripts are included and must be relocated to the GMAT output path specified in gmat_startup_file.txt.

The environment variable %LOCALUSERAPP% must be defined and contain the path to the top level GMAT executable folder.  
The python scripts will search this path for gmat_startup_file.txt and use the OUTPUT_PATH defined in the most recent file.
