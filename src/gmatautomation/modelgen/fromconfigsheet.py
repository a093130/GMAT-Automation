#! python
# -*- coding: utf-8 -*-
"""
    @file fromconfigsheet.py

    @brief: This module reads configurations from the "Vehicle Optimization Tables" workbook 
    and returns a corresponding specification of various cases of GMAT model resources. It
    is currently a point design in support of variations in the AlfanoXfer.script GMAT models.

    @copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.
    XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
    https://docs.xlwings.org/en/stable/license.html
    
    @author  Colin Helms, colinhelms@outlook.com, [CCH]
 
    @details Interface Agreement: An Excel workbook exists which contains a 
    sheet named "GMAT" having a contiguous table starting in cell "A1".  
    The table consists of a first line of parameter names and successive lines 
    of spacecraft properties and relevant hardware configuration.
    
    The first line of table headings may not be exactly the same as GMAT resource names.  
    The associated routine, "modelpov.py" defines a mapping of required GMAT
    resource names to worksheet table headings, which we refer to as parameter names.  
    Procedure modelgen.py will use this association to generate the correct 
    GMAT resource names.  Procedure fromconfigsheet.py will extract only the parameter 
    names defined in modelpov.py. Note that it is possible with this logic to extract NO
    data from the workbook, in this case the model.pov file may be edited to include
    the intended parameter names, or the workbook may be so edited.
    
    Variation of orbital elements is independent of hardware configuration.  Specifically,
    inclination cases may be multiple for the given "GMAT" table and are gleaned from
    a separate n x 1 table of values in named range, "Inclinations" contained in
    a sheet named "Mission Params".
    
    Similarly, cases of initial epoch to be executed are gleaned from n x 4 table of values
    in named range, "Starting Epoch" on a sheet named "Mission Params". Each row, n,
    contains a UTC formatted time and date value in column 1, e.g. 
    "20 Mar 2020 03:49:00.000 UTC".

    For display, a GMAT viewpoint vector consisting of x, y, and z components of
    rendering camera position (in the J2000 ECI coordinate system) are associated
    with each epoch value, and are contained in columns (n,2), (n,3), and (n,4) of
    the "Starting Epoch" named range.
    
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

        
"""
import os
import re
import logging
import numpy as np
from win32ctypes import pywintypes as pwin
import xlwings as xw
from PyQt5.QtWidgets import(QApplication, QFileDialog)
from gmatautomation import modelpov as pov

class Ultima(Exception):
    """ Enclosing exception to ensure that cleanup occurs. """
    def __init__(self, source='fromconfigsheet.py', message='Exception caught, exiting module.'):
        self.source = source
        self.message = message
        logging.warning(self.message)

def sheetvars(sht):
    """ This is the configspec. The size of this range is variable. """

    logging.debug('Function sheetvars() called.')
    
    if sht != None:       
        return sht.range('A1').expand().value
    else:
        return None    
                               
def mission(sht):
    """ Reads the mission name for use as root of output filenames """

    logging.debug('Function mission() called.')
      
    if sht != None:       
        return sht.range('Mission_Name').value    
    else:
        return None
   
def smavars(sht):
    """ Reads the initial and final SMA values from workbook. 

    """

    logging.debug('Function smavars() called.')
      
    if sht != None:       
#        return np.around(sht.range('Altitude').value, 4)
        smarange = np.around(sht.range('Altitude').value, 4)
        return np.array(smarange, ndmin=2)
        """
        A corner case exists where the model spec contains only one line of SMA.
        A minimum dimension of 2 is expected.
        """
    else:
        return None

def epochvars(sht):
    """ Reads list of starting epoch values from workbook.
    
    A corner case exists where the model spec contains only one line of inclination.
    A minimum dimension of 2 is expected.      
    """

    logging.debug('Function epochvars() called.')
      
    if sht != None:       
#        return sht.range('Starting_Epoch').value
        epochrange = sht.range('Starting_Epoch').value
        return np.array(epochrange, ndmin=2)
        """
        A corner case exists where the model spec contains only one line of Epoch.
        A minimum dimension of 2 is expected.
        """
    else:
        return None

def inclvars(sht):
    """ Reads and returns the inclination table from the workbook. 
    The inclination table is returned as a list of two columns and m rows, where each row
    identifies one model case for simulation of an inclination change.
    Each row contains a floating point inclination value (positive or negative)
    and a floating point costate value (negative).  
    
    A corner case exists where the model spec contains only one line of Inclination.
    A minimum dimension of 2 is expected.
    """

    logging.debug('Function inclvars() called.')
      
    if sht != None:       
#        return np.around(sht.range('Inclinations').value, 3)  
        inclrange = np.around(sht.range('Inclinations').value, 3)
        return np.array(inclrange, ndmin=2)
        """
        A corner case exists where the model spec contains only one line of inclination.
        A minimum dimension of 2 is expected.
        """
    else:
        return None

def modelspec(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Access the named Excel workbook and return a list of dictionaries 
    corresponding to the table structure in the GMAT tab. """

    logging.debug('Function modelspec() called.')

    try:
        excel = xw.App()
        excel.visible=False
        
        wingbk = excel.books.open(fname)
        msheet = wingbk.sheets['Mission Params']
        ssheet = wingbk.sheets('GMAT')
        
    except OSError as ouch:
        logging.error('Open {0} failed. \nOS error: {1}.'.format(ouch.strerror, ouch.filename))

    except pwin.com_error as ouch:
        logging.error('Access to sheet raised Windows com error. {0}, {1}'.format(type(ouch), ouch.args[1]))
    
    try:        
        """ Get the configspec table. """
        spectable = sheetvars(ssheet)
        
        """ Develop an index of table column names to column number.
        The column names in row 0 of the table are popped, e.g. removed. 
        """
        configspec = spectable.copy()
        """ List configspec is the array of specified values for different GMAT model resources:
        one row for each configuration specified in the GMAT tab of the workbook."""
        tablenames = configspec.pop(0)
        """ List tablenames is a list of the worksheet table headings. 
        Note that pop() removes this row from configspec. 
        """
        modelspec = pov.getvarnames()
        """ Dictionary modelspec contains the GMAT worksheet resource-to-heading 
        association. """
        
        logging.debug('Variables in model configuration spec:\n%s', str(tablenames))
        
        specindex = {}
        for col, name in enumerate(tablenames):
            """ Dictionary "specindex" associates tablename as key to 
            the value of the worksheet column number.  It is used in the loop that
            follows.
            """
            specindex[name] = col
                
        for resource, tablename in modelspec.items():
            """ Map the GMAT model resource name to the worksheet column number. """
            
            if tablename in specindex:
                """ Match the variable name specified by the modelpov module with the name
                from the specindex dictionary. If the variable is not found, then this 
                resource will not be included in the generated GMAT model file.
                """
                modelspec[resource] = specindex[tablename]
                """ Replace the modelspec tablename with the column value.  The specindex
                contains column numbers associated with tablenames as keys.
                Now modelspec contains a column number in association with the GMAT
                resource name.
                """
            else:
                logging.warn('Variable name %s not found in workbook.', str(tablename))
    
        """ Get the epoch list, the inclination list and the mission name from
        the mission parameters tab of the workbook. The number of cases will be the 
        number of rows of configspec x number of epochs x number of inclinations.
        """
        epochlist = epochvars(msheet)
        """ List epochlist contains possible multiple values for gmat starting epoch associated to 
        the corresponding viewpoint vector.
        """
        inclist = inclvars(msheet)
        """ List inclist contains the multiple values selected for modeling inclination
        change, each inclination value is associated with an Alfano inclination costate.
        """
        smalist = smavars(msheet)
        """ List smalist contains the initial and final values of semi-major axis. """
        
        rege_comma = re.compile(',+')
        rege_utc = re.compile(' UTC')
        rege_spc = re.compile(' +')
        
        cases = []
        case = {}
        for row, data in enumerate(configspec):
            """ Generate a list of model inputs for the required GMAT batch runs.
            The list "cases" contains rows of dictionaries. 
            Each dictionary is a combination of configspec and modelspec formed
            by associating the data value from configspec to a key which is the 
            GMAT resource name from modelspec.            
            """            
            for resource, col in modelspec.items():
                """ Generate the case corresponding to the row of configspec
                using the resource name and column number in modelspec. The table heading 
                was replaced with its column number in modelspec above.
                """
                case[resource] = data[col]
                               
            for epoch, xview, yview, zview in epochlist:
                """ Elaborate the list of cases based on mission_params. """

                view = [ np.float(xview), np.float(yview), np.float(zview) ]
                
                for isma, fsma in smalist:
                    for incl, costate in inclist:
                        """ Elaborate the list of cases, a new line for each inclination. """
                        
                        start_incl = np.abs(incl)
                        case['COSTATE'] = costate                        
                        case['EOTV.INC'] = start_incl
                        
                        if start_incl == 0:
                            
                            case['MORE'] = 1
                            """ If the initial inclination is 0, the assumption is that the inclination
                            change will be positive, i.e. a return trajectory from geosynchronous inclination.
                            """                            
                        else:                       
                            case['MORE'] = incl/start_incl
                            """ The convention is that a negative value of inclination indicates a recquirement
                            to decrease inclination and positive value a requirement to increase inclination.
                                                        
                            The amount of inclination change is determined by the costate value.
                            """
                                              
                        case['EOTV.Epoch'] = epoch
                        case['DefaultOrbitView.ViewPointVector'] = view
                        case['SMA_INIT'] = isma
                        case['SMA_END'] = fsma
                        
                        cases.append(case.copy())
                                    
        logging.debug('Output is: %s', repr(cases))
                
        for case in cases:
            """ Fix GMAT syntax incompatibilities and inconsistencies. """
            
            case['ReportFile1.Filename'] = str(case['ReportFile1.Filename'])
            case['EOTV.Epoch'] = str(rege_utc.sub('', case['EOTV.Epoch']))
            case['DefaultOrbitView.ViewPointVector'] = \
                rege_comma.sub('', repr(case['DefaultOrbitView.ViewPointVector']))
                
            case['EOTV.Id'] = rege_spc.sub('', str(case['ReportFile1.Filename']))
                
        return (cases)
                 
    except Ultima as u:
        logging.info('%Error termination.')
        
    finally:
        srcname = os.path.basename(fname)
        excel.books[srcname].close()
        excel.quit()
        
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
    
    fname = QFileDialog().getOpenFileName(None, 'Open Configuration Workbook', 
                       os.getenv('USERPROFILE'))
        
    logging.info('Configuration workbook is: %s', fname[0])
    
    try:
        cases = modelspec(fname[0])
        
        logging.info('Terminating: number of cases in modelspec: {0}'.format(len(cases)))
      
    except Ultima as u:
        logging.info('%s %s', u.source, u.message)
    
    finally:
        app.quit()
        logging.shutdown()

    
    
    
    
    
