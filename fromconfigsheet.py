# -*- coding: utf-8 -*-
"""
Created on Sat Oct 20 09:53:28 2018

@author: colinhelms@outlook.com

@Description:
    This module reads configurations from the "Vehicle Optimization Tables" workbook 
    and returns a corresponding specification of various cases of GMAT model resources.
    
    Interface Agreement:
    An Excel workbook exists which contains a sheet named "GMAT" having a
    contiguous table starting in cell "A1".  The table consists of a first line
    of parameter names and successive lines of spacecraft properties and 
    relevant hardware configuration.
    
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
    
@Revisions
    08 Jan 2019, initial baseline  

            
    
"""
import os
import logging
from PyQt5.QtWidgets import(QApplication, QFileDialog)
import pywintypes as pwin
import xlwings as xw
import modelpov as pov

class Ultima(Exception):
    """ Enclosing exception to ensure that cleanup occurs. """
    def __init__(self, source='fromconfigsheet.py', message='Exception caught, exiting module.'):
        self.source = source
        self.message = message
        logging.warning(self.message)

def sheetvars(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Reads configs from workbook at given file path returns table from top left cell. """

    logging.debug('Function sheetvars() called.')
    
    try:
        wb = xw.Book(fname)
    except pwin.com_error as com:
        logging.error('Open workbook %s failed in call to sheetvars(). %s %s',\
        fname, str(type(com)), str(com.args[1]))
        
        raise Ultima()
    except OSError as ouch:
        logging.error('Open workbook failed in call to sheetvars(). \nOS error: %s.\nfilename:',\
                      ouch.strerror, ouch.filename)        
        raise Ultima()

    try:    
        sht = wb.sheets('GMAT')
    except pwin.com_error as ouch:
        logging.error('Access to sheet \"GMAT\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        
        raise Ultima()

    return sht.range('A1').expand().value
            
def epochvars(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Reads list of starting epoch values from workbook. """

    logging.debug('Function epochvars() called.')
    
    try:
        wb = xw.Book(fname)
    except pwin.com_error as com:
        logging.error('Open workbook %s failed in call to epochvars(). %s %s',\
                      fname, str(type(com)), str(com.args[1]))
        
        raise Ultima()
    except OSError as ouch:
        logging.error('Open workbook failed in call to epochvars(). \nOS error: %s.\nfilename:',\
                      ouch.strerror, ouch.filename)        
        raise Ultima()
        
    try:
        sht = wb.sheets('Mission Params')
    except pwin.com_error as ouch:
        logging.error('Access to sheet \"Mission Params\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima()

    try:
        return sht.range('Starting_Epoch').value
    except pwin.com_error as ouch:
        logging.error('Access to range \"Starting_Epoch\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima()

def inclvars(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Reads inclination from workbook. """

    logging.debug('Function inclvars() called.')
    
    try:
        wb = xw.Book(fname)
        
    except pwin.com_error as com:
        logging.error('Open workbook %s failed in call to inclvars(). %s %s',\
                      fname, str(type(com)), str(com.args[1]))       
        raise Ultima()
    except OSError as ouch:
        logging.error('Open workbook failed in call to inclvars(). \nOS error: %s.\nfilename:',\
                      ouch.strerror, ouch.filename)        
        raise Ultima()
         
    try:    
        sht = wb.sheets('Mission Params')
        
    except pwin.com_error as ouch:
        logging.error('Access to sheet \"Mission Params\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima
       
    try:    
        return sht.range('Inclinations').value
    
    except pwin.com_error as ouch:
        logging.error('Access to range \"Inclinations\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima

def mission(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Reads the mission name for use as root of output filenames """

    logging.debug('Function mission() called.')
    
    try:
        wb = xw.Book(fname)
        
    except pwin.com_error as com:
        logging.error('Open workbook %s failed in call to mission(). %s %s',\
                      fname, str(type(com)), str(com.args[1]))       
        raise Ultima()
    except OSError as ouch:
        logging.error('Open workbook failed in call to mission(). \nOS error: %s.\nfilename:',\
                      ouch.strerror, ouch.filename)        
        raise Ultima()
        
    try:    
        sht = wb.sheets('Mission Params')
        
    except pwin.com_error as ouch:
        logging.error('Access to sheet \"Mission Params\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima
        
    try:
        return sht.range('Mission_Name').value
    
    except pwin.com_error as ouch:
        """ If no mission name defined, simply return a default string. """
        logging.error('Access to range \"Mission_Name\" raised com error. %s %s',\
                      str(type(ouch)), str(ouch.args[1]))
        wb.close()
        raise Ultima
    
def modelspec(fname=r'Vehicle Optimization Tables.xlsx'):
    """ Access the named Excel workbook and return a list of dictionaries 
    corresponding to the table structure in the GMAT tab. """

    logging.debug('Function modelspec() called.')
    
    try:        
        """ Get the configspec table. """
        spectable = sheetvars(fname)
        
        """ Develop an index of table column names to column number.
        The column names in row 0 of the table are popped, e.g. removed. 
        """
        configspec = spectable.copy()
        """ List configspec is the array of specified values for different GMAT model resources:
        one row for each configuration specified in the GMAT tab of the workbook."""
        tablenames = configspec.pop(0)
        """ List tablenames is a list of the worksheet table headings. Note that pop() 
        removes this row from configspec. """
        modelspec = pov.getvarnames()
        """ Dictionary modelspec contains the GMAT model resource to table heading 
        association. """
        
        logging.info('Variables in model configuration spec:\n%s', str(tablenames))
        
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
        epochlist = epochvars(fname)
        """ List epochlist contains possible multiple values for gmat starting epoch associated to 
        the corresponding viewpoint vector.
        """
        inclist = inclvars(fname)
        """ List inclist contains the possible multiple values for gmat inclination
        associated to the corresponding Alfano inclination costate.
        """
        
        case = {}
        cases = []
        epoch_elab = {}
        incl_elab = {}
        for row, data in enumerate(configspec):
            """ Generate a list of model inputs for the required GMAT batch runs.
            The list "cases" contains rows of dictionaries. 
            Each dictionary is a combination of configspec and modelspec formed
            by associating the data value from configspec to a key which is the 
            GMAT resource name from modelspec.            
            """
            
            case.clear()
            for resource, col in modelspec.items():
                """ Generate each case using the resource name and column number in
                modelspec.
                """
                case[resource] = data[col]
               
            epoch_elab.clear()
            for epoch in epochlist:
                """ Elaborate the list of cases, a new line for each epoch. """
                epoch_elab = case.copy()
                
                epoch_elab['EOTV.Epoch'] = epoch[0]
                epoch_elab['DefaultOrbitView.ViewPointVector'] = epoch[1:4]
                
                for incl in inclist:
                    """ Elaborate the list of cases, a new line for each inclination. """
                    incl_elab = epoch_elab.copy()
                    incl_elab['EOTV.INC'] = incl
                    cases.append(incl_elab)
                                    
        logging.debug('Output is: %s', repr(cases))
        logging.info('Nominal termination. Rows processed = %s', row+1)
        
        return (cases)
                 
    except Ultima as u:
        logging.debug('Output is: %s', repr(cases))
        logging.info('%Error termination.')
        
if __name__ == "__main__":
    """
    Test case and example of use.
    """    
    logging.basicConfig(
            filename='./appLog.log',
            level=logging.DEBUG,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info('Started.')
        
    app = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open Configuration Workbook', 
                       os.getenv('USERPROFILE'))
        
    logging.info('Configuration workbook is: %s', fname[0])
    
    try:
        m_name = mission(fname[0])
    except Ultima as u:
        logging.info('%s %s', u.source, u.message)

    try:
        cases = modelspec(fname[0])
    except Ultima as u:
        logging.info('%s %s', u.source, u.message)
    finally:
        logging.debug('For mission, %s cases are:\n%s', m_name, repr(cases))
        
    print('For mission', m_name, 'GMAT simulation cases are:\n', repr(cases))
    
    logging.shutdown()
    
    
    
    
    
