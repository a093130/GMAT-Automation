#! python
# -*- coding: utf-8 -*-
"""
    @file fromconfigsheet.py

    @brief: This module reads configurations from the model spec workbook 
    and returns a corresponding specification of various cases of GMAT model resources for inclusion in 
    GMAT batch script.

    @copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.
    XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
    https://docs.xlwings.org/en/stable/license.html
    
    @author  Colin Helms, colinhelms@outlook.com, [CCH]

    @version 0.4b0
 
    @details Interface Agreement: An Excel workbook exists which contains a 
    sheet named having a contiguous, table of variables starting in cell "A1".  
    The table row 0 contains parameter names and successive rows 1-n contain values
    that must be varied model by model in n successive batch runs of GMAT.
    
    The first line of table headings may not be the same as GMAT resource names.  
    The associated routine, "modelpov.py" defines a mapping of required GMAT
    resource names to worksheet table headings, which we refer to as parameter names.  
    Procedure modelgen.py will use this association to generate the correct 
    GMAT resource names.  Procedure fromconfigsheet.py will extract only the parameter 
    names defined in modelpov.py. Note that it is possible with this logic to extract NO
    data from the workbook, in this case the model.pov file may be edited to include
    the intended parameter names, or the workbook may be so edited.

    All use of Excel is isolated to this module such that Pandas or csv data sources
    may be accommodated by additional modules TBS.
    
    Inputs:
        fname - this is the path specification for the "Vehicle Optimizations
        Tables" workbook.  The QFileDialog() from PyQt may be used to browse for the
        workbook file.
        
    @remarks:
        Sat Oct 20 09:53:28 2018, [CCH] Created
        09 Feb 2019, [CCH]commit to GitHub repository GMAT-Automation, Integration Branch.
        30 Apr 2019, [CCH] Flow Costates and payload mass through to model from worksheet.
        Wed Apr 20 2022 [CCH] Reordered files and included in sdist preparing to build.
        Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.
        05 May 2022 - [CCH] Build version 0.3b3 for PyPI and upload as open source.
        16 Jun 2022, [CCH] https://github.com/a093130/GMAT-Automation/issues/1 (refactor modelspec)
        05 Dec 2023, [CCH] version 0.4b0, refactored to eliminate Alfano specializations

        
"""
import os
import sys
import logging
import pywintypes as pwin
import xlwings as xw
from pathlib import Path
from userexceptions import Ultima
from specializedmodelpov import CUserSpecializedPov
from PyQt5.QtWidgets import(QApplication, QFileDialog)

"""
@defgroup Globals
@brief Python global constants are persistent across calls to functions.  The ALLCAPS case
is reserved for identification of globals.
"""
# @{
""" Insert global definitions in group. """
ROOTPATH = Path(sys.argv[0]).parent.resolve()
STARTPATH = ROOTPATH.parents[1]
DATAPATH = STARTPATH.joinpath('data/')

# @} End of Globals


"""	@defgroup Classes
	@brief Class definitions are grouped for easy reference in documentation.
"""
# @{
""" Insert Class defs in group. """

#Tailored class strategic comment as follows:
"""
	@brief description
	@details detailed description
"""

# @} End of Class Group


"""	@defgroup Functions
	@brief Function Implementations are grouped for easy reference in documentation.
"""
# @{
""" Insert function defs in group. """

#Tailored function strategic comment as follows:

""" 
	@brief description
	
	@details detailed description
		
	@param <type> name [IN] - description
	@param <type> name [OUT] - description

	@return condition1 = value1
	@return condition2 = value2

""" 

def retrievespec(wingbk):
    """ @ brief Access the named Excel workbook and return a list of dictionaries 
    corresponding to the table structure.

    @description modelgen calls this function.

    @ raises pwin.com_error
    """
    pov = CUserSpecializedPov()

    sheetname = pov.sheetnames['GMAT']
    ssheet = wingbk.sheets(sheetname)

    """ configspec is a table of values for different GMAT model resources corresponding
    to the names in row 0, each row specifies each configuration of model to be produced.
    The specialized CModelPov instance should map the table names (row 0) to GMAT model
    parameters.  configspec may be quite large.
    Note that the Excel server delivers all numbers as floats all dates as datetime.  If
    a number should be delivered in a different type (int), either use the range options
    feature e.g., "sheet['A1'].options(numbers=int).value" or change the Excel column type
    for instant change the column type to "text" if a number is actually an ID.
    """
    configspec = (ssheet.range('A1').expand().value).copy()

    """ Develop an index of table column names to column number, in order to correctly
    assign values to GMAT model parameters.
    The column names in row 0 of the table are removed by the pop operation
    """
    tablenames = configspec.pop(0)
    logging.debug('The read table headings are:\n%s', str(tablenames))
    
    modelspec = pov.getspecindex(tablenames)
    fmtspec = pov.gmatreqfmt

    """ Generate a list of model inputs for the required GMAT batch runs.
    The list "cases" contains rows of dictionaries. 
    Each dictionary is a combination of configspec and modelspec formed
    by associating the data value from configspec to a key which is the 
    GMAT resource name from modelspec.        
    """        
    cases = []
    case = {}
    for row, data in enumerate(configspec):
        
        for resource, col in modelspec.items():
            """ Generate the case corresponding to the row of configspec
            using the resource name and column number in modelspec. 
            
            The table heading was replaced with its column number in modelspec above.
            """

            """ GMAT is particular about the format of an ID type, must be integer. The default number
            format is float formatted string. The pov contains the specific format requirement. 
            At this point the only GMAT compatibility problem knonw is an Integer ID formatted as float. 
            """
            if fmtspec[resource] == 'd':
                data[col] = int(float(data[col]))

            case[resource] = data[col]

        """ In the following, we're using the case dictionary to forward 
        the values of several additional GMAT parameters to the model writer
        as retrieved from the specialized instance of modelpov.
        """
        cases.append(case.copy())
        case.clear()     

    logging.info ('There are %d cases defined', len(cases))                     
    logging.debug('The cases are:\n %s', repr(cases))

    return (cases)


# @} End of Functions Group


if __name__ == "__main__":
    """
    Test case and example of use.
    """    
    logging.basicConfig(
            filename='./configsheet.log',
            level=logging.DEBUG,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info('Started.')
        
    app = QApplication([])

    
    msg = 'Get the Excel workbook containing the model spec.'
    
    fname = QFileDialog().getOpenFileName(None, 
                    msg, 
                    str(STARTPATH),
                    filter='text files(*.xlsx)') 
    
    filepath = Path(fname[0])
    filename = filepath.name

    msg = 'The selected model spec is: {0}'.format(filename)
    print(msg)
    logging.info(msg)
    
    try:
        excel = xw.App()
        excel.visible=False
    
        wingbk = excel.books.open(fname[0])
        
        msg = 'File open successful, reading data from file {0}.'.format(filename)
        print(msg)
        logging.info(msg)

        cases = retrievespec(wingbk)
        wingbk.close()

        msg = 'The number of cases in the model spec: {0}'.format(len(cases))
        print(msg)
        logging.info(msg)

    except OSError as ouch:
        logging.error('Open {0} failed. \nOS error: {1}.'.format(ouch.strerror, ouch.filename))

    except pwin.com_error as ouch:
        logging.error('Access to sheet raised Windows com error. {0}, {1}'.format(type(ouch), ouch.args[1]))

    except Ultima as u:
        logging.info('%s %s', u.source, u.message)
    
    finally:
        app.quit()
        excel.quit()
        logging.shutdown()

    
    
    
    
    
