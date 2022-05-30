# GMAT Automation
The GMAT Automation package is a related set of Python scripts used to generate and run batches of GMAT model files. Various scripts are also provided to reduce the ReportFiles and Contact Locator files that GMAT execution produces.

Three modules are provided:
-modelcontrol: includes the gmat_batcher to execute batches of model scripts in a multithreaded mode.

-modelgen: extract GMAT model specifications from a spreadsheet and build model batch files 

-reportgen: Format and post process GMAT output reports into Excel file form,

See docs/GMATAutomation_SoftwareUserManual.docx for further detail.

The following external dependencies must be installed to Python/Lib/site-packages.

-datetime

-re

-numpy
	
-scipy
	
-subprocess
	
-pywin32
	
-xlwings
	
-xlsxwriter
	
-PyQt5
	