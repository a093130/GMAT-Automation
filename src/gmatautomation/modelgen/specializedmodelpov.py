#! Python
# -*- coding: utf-8 -*-

#Tailored Header and Implementation file strategic leading comment as follows:
#Tailored for doxygen.
""" 
	@file m4modelpov.py
	@brief specialized point of view class for AstroForge
	
	@copyright 2023 AstroForge-Incorporated
	@author  Colin, [CCH]
	@authoremail colinhelms@outlook.com, colin@astroforge.io

	@version 0.1a0

	@details Derived from CModelPov in order to characterize the model spec.
	
	@remark Change History
		05 Dec 2023: [CCH] File created

	@bug [<initials>] <backlog item>

"""

import os
import sys
import logging
import time
from pathlib import Path
from modelpov import CModelPov

from PyQt5.QtWidgets import(QApplication, QFileDialog)

"""
@defgroup Globals
@brief Python global constants are persistent across calls to functions.  The ALLCAPS case
is reserved for identification of globals.
"""
# @{
""" Insert global definitions in group. """

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

class CUserSpecializedPov(CModelPov):
    """
        @brief specializes for the Mission 4 mission sequence model spec

        @details This POV instance maps between headings returned from the AstroForge BigQuery
        mission-design project by the arrival_position_velocity_by_body query.

        The reportvars instance returns headings to be uploaded to the mission-design project.
    """

    def __init__(self, **args):
        super().__init__(**args)

        self.mission_name = 'AF_M4'

        self.model_name = 'M4_LowThrustSystemDyn'

        """ This is filled in by the getkernelname() method, below. """
        self.kernelfiledef = ''
        self.naifdef = ''
        self.kernel = ''
        self.body = ''

        """ This is filled in by the mknameroot() method below. """
        self.nameroot =''
        
        """ These is a dictionary of Excel sheet names using polymorphic key names. """
        self.sheetnames = {'GMAT': 'Sheet1' }

        """ Any user variables that need to be created on-the-fly in the form {varname : GMAT Create Var|String}.  """
        self.varset = {}
        
        """ These variables map to the table headings in the range returned from sheetnames['GMAT']. 
        They are used by fromconfigsheet.py to build up the case dictionary. """
        self.var2gmat = {'NEA.NAIFId': 'Target_Id', 
                         'body': 'Target_Id', 
                         'mjdarr': 'Arrival__MJD', 
                         'mjddep': 'Departure__MJD',
                         'Eitri.X': 'r20__km',
                         'Eitri.Y': 'r21__km', 
                         'Eitri.Z': 'r22__km',
                         'Eitri.VX': 'v_arr0__km_per_s', 
                         'Eitri.VY': 'v_arr1__km_per_s', 
                         'Eitri.VZ': 'v_arr2__km_per_s'}
        
        self.gmatreqfmt ={'NEA.NAIFId': 'd', 
                         'body': 'f', 
                         'mjdarr': 'f', 
                         'mjddep': 'f',
                         'Eitri.X': 'f',
                         'Eitri.Y': 'f', 
                         'Eitri.Z': 'f',
                         'Eitri.VX': 'f', 
                         'Eitri.VY': 'f', 
                         'Eitri.VZ': 'f'}
        
        """ Resources that are not included as keys in a case. Typically derived from values in a case e.g., ReportFile names. """  
        self.resources = {}

        """ These report variables will be included in the output reports, where the report is build on-the-fly.  Alternatively
        just place the report definition in the StaticDefinitions.include file. 
        Note the list is a tuple such that multiple report files are accommodated. 
        """
        self.reportvars = [[]] # not used in this scenario.
        
        """ varmultiplier are continuation variables read from the 'mission' Sheet. """
        self.varmultiplier = [] # not used in this scenario.

    
    def getspecindex(self, tablenames):
        """ 
            @brief given tablenames, associates a resource name from var2gmat, initializes resource dictionary
            
            @details double associations created, first tablenames to columns, then using initialized instance of 
            GMAT resources to table_names, associates columns with GMAT resources.  Used to bind values in modelspec
            to the appropriate GMAT resource.
                
            @param <List> tablenames [IN] - 1 x N list of table names from the modelspec.

            @return condition1 = a dictionary providing the association between GMAT resources and columns of tablenames
        """
       
        """ Dictionary modelspec contains the GMAT worksheet resource-to-heading 
        association. 
        """
        specindex = {}
        modelspec = {}
        for col, name in enumerate(tablenames):
            specindex[name] = col
    
        resource2tablename = self.var2gmat.copy()

        if len(tablenames) == len(resource2tablename):

            """ Map the list elements of configspec to the GMAT variables in modelspec using the relationship
            between GMAT resource names and configspec tablenames contained in var2gmat.
            """     
            for resource, tablename in resource2tablename.items():
                if tablename in specindex:

                    """ The reource key is associated with a list element number. """
                    modelspec[resource] = specindex[tablename]

                else:
                    msg = 'table name %s not found in specindexn dict for resource key %s', str(tablename), str(resource)
                    logging.error(msg)
                    raise IndexError(msg)
                
            return modelspec.copy()
            
        else:
            msg = 'Number of workbook column names does not match number of GMAT resources.'
            logging.error(msg)
            raise ValueError(msg)
            
    
    def getreportfile(self, case):
        """ 
            @brief create the specialized GMAT resource name for GMAT reportfile path
            
            @details Multiple report files can be defined here and added to the resources dictionary.
            Note that the use of single quotes and brackets in the definition string uses GMAT syntax
            which is subtle and the GMAT script will fail if not correct.   This syntax has proven very 
            difficult to recreate programmatically and accounts for the particular use of double quotes
            and single quotes in the strings generated below.
                
            @param <dict> case [IN] - from case, use body, mjddep and mjdarr values in order to create .
            a desciptive, timetagged string.

            @return string List of GMAT file name resource definitions.
        """   
        
        timetag = time.strftime('J%j_%H%M%S',time.gmtime())

        reportfilename = "Report_" + self.mission_name +\
            "_body_" + str(case['body']) + \
            "_mjddep_" + str(case['mjddep']) + \
            "_mjdarr_" + str(case['mjdarr']) + \
            "_at_" + timetag + ".csv"

        reportfile1 = "ReportFile1.Filename =" + \
            "'" + reportfilename + "'"
       
        self.resources['ReportFile1'] = reportfile1

        return 
    
    
    def getdebugreport(self, case):
        """ 
            @brief create the specialized GMAT resource name for GMAT reportfile path
            
            @details The specialized model pov knows which files go with what definition.  Note that the
            use of single quotes and brackets by GMAT is very subtle and differs between resource types.
            This pattern was very difficult to satisfy and accounts for the particular use of double quotes
            and single quotes in the strings generated below.
                
            @param <dict> case [IN] - from case, use body, mjddep and mjdarr values in order to create .
            a desciptive, timetagged string.

            @return string List of GMAT file name resource definitions.
        """

        timetag = time.strftime('J%j_%H%M%S',time.gmtime())

        self.debugreportfilename = "Debug_" + self.mission_name + \
            "_body_" + str(case['body']) + \
            "_mjddep_" + str(case['mjddep']) + \
            "_mjdarr_" + str(case['mjdarr']) + \
            "_at_" + timetag + ".txt"

        self.debugreportfile = "DebugReport.Filename =" + \
            "'" + self.debugreportfilename + "'"    

        self.resources['DebugReport'] = self.debugreportfile

        return 

    def getkernelname(self,case):
        """ 
            @brief create the specialized GMAT resource name for GMAT reportfile path
            
            @details The specialized model pov knows which files go with what definition.  Note that the
            use of single quotes and brackets by GMAT is very subtle and differs between resource types.
            This pattern was very difficult to satisfy and accounts for the particular use of double quotes
            and single quotes in the strings generated below.
                
            @param <dict> case [IN] - from case, use body, mjddep and mjdarr values in order to create .
            a desciptive, timetagged string.

            @return string List of GMAT file name resource definitions.
        """
        bodynumber = str(case['body'])
        self.body = bodynumber.split('.')[0]
        self.kernel = self.body + '.bsp'
        self.naifdef = 'NEA.NAIFId = ' + self.kernel

        kernelpath = Path('../data/nea_ephem') / self.kernel
        self.kernelfiledef = "NEA.OrbitSpiceKernelName =" + "{'" + str(kernelpath) + "'}"

        self.resources['OrbitSpiceKernelName'] = self.kernelfiledef

        return
    

    def mknameroot(self, case):
        """ 
            @brief fromconfigsheet.py uses this function to make a descriptive nameroot.
            
            @details uses specialized
                
            @param <dict> case [IN] - body, mjddep and mjdarr values read out in order to create .
            a desciptive, timetagged string.  This is meant to name the batch files created by modelspec.

            @return condition1 = <string> root of file names  associated with the case.

        """
        bodynumber = str(case['body'])
        kernel = bodynumber.split('.')[0]
        dep = str(case['mjddep'])
        mjddep = dep.split('.')[0]
        arr = str(case['mjdarr'])
        mjdarr = arr.split('.')[0]

        self.nameroot = self.mission_name + \
            '_body_' + kernel + \
            '_mjddep_' + mjddep + \
            '_mjdarr_' + mjdarr + \
            '_at_' + time.strftime('J%j_%H%M%S',time.gmtime())

        return self.nameroot
        
# @} End of Class Group


# Functions may be grouped in documentation

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

# @} End of Functions Group
    

if __name__ == "__main__":

    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code.
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
        
     This is the top-level entry point for the GMAT Model Generation. 
    """

    logging.basicConfig(
        filename='./specializedmodelpov.log',
        level=logging.INFO,
        format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', 
        datefmt='%d%B%Y_%H:%M:%S')

    logging.info('******************** Specializedmodelpov Test Cases Started ********************')

    """ Test case processes this representative one-line model spec. """
    configspec = [
    ['Target_Id','Departure__MJD','Arrival__MJD','r20__km','r21__km','r22__km','v_arr0__km_per_s','v_arr1__km_per_s','v_arr2__km_per_s','ve_magnitude__km_per_s'],
    ['3702319','30356','30696','-123589834.9','118242769.9','702440.2607','-24.4092748','-17.52189312','-0.381848534','0.510072548']]

    pov = CUserSpecializedPov()

    msg = 'mission name: {0}'.format(pov.mission_name)
    print(msg)
    logging.info(msg)

    msg = 'model name: {0}'.format(pov.model_name)
    print(msg)
    logging.info(msg)

    msg = 'sheet names: {0}'.format(pov.sheetnames)
    print(msg)
    logging.info(msg)

    tablenames = configspec.pop(0)
    data = configspec[0]

    try:
        modelspec = pov.getspecindex(tablenames)

        """ mimic configsheet.retrievespec() """
        case = {}
        for resource, col in modelspec.items():
            case[resource] = data[col]
            msg = "The case is resource: {0} : value {1}".format(resource, data[col])
            print(msg)
            logging.info(msg)

        nameroot = pov.mknameroot(case)

        msg = "the created nameroot is {0}".format(nameroot)
        print(msg)
        logging.info(msg)
        
        pov.getreportfile(case)
        pov.getdebugreport(case)
        pov.getkernelname(case)

        res = pov.resources
        for resource, name in res.items():
            msg = "The resource definition is {0}:{1}".format(resource, name)
            print(msg)
            logging.info(msg)

        body = pov.body
        msg = "The created body ID is {0}".format(body)
        print(msg)
        logging.info(msg)

        kernel = pov.kernel
        msg = "The created kernel is {0}".format(kernel)
        print(msg)
        logging.info(msg)

        naifId = pov.naifdef
        msg = "The created NAIF reference is {0}".format(naifId)
        print(msg)
        logging.info(msg)

        kernel = pov.kernelfiledef
        msg = "The created kernel file reference is {0}".format(kernel)
        print(msg)
        logging.info(msg)

    except OSError as err:
        logging.error("OS error: {0}".format(err))

    except:
        logging.error("Unexpected error:", sys.exc_info())

