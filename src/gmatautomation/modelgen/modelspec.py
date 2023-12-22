#! python
# -*- coding: utf-8 -*-
"""
    @file ModelSpec.py

    @brief: module container for class CModelSpec.  Refactored from modelgen.py.  

	@copyright 2023 Astroforge.incorporated
	@author  Colin Helms, [CCH]
	@authoremail colin@astroforge.io, colinhelms@outlook.com

    @version 0.4b0

    @details: This module provides GMAT scripts to implement batch processing of different 
    mission scenarios in which the spacecraft configuration and/or initial orbital elements 
    vary.
    
    The approach utilizes the GMAT 2018a #Include macro, which loads resources and 
    script snippets from external files.  The script creates Include files whose
    parameter and resource values vary in accordance with an Excel workbook.
    
    A top level GMAT script template is assumed to exist in the GMAT model directory.
    This script template will be copied and modified by modelgen.py.
    The script template must contain the GMAT create statements for GMAT resources.
      
    The script template must contain three #Include statements as follows:
    (1) a GMAT script file defining the static resources, those that don't change per run
    (2) An include file defining those variables written by CModelSpec.
    (3) An include file containing the GMAT Mission Sequence definition.
    
    The second #Include statement will be modified by CModelSpec to include a
    uniquely named filepath to the entire template copied to a "batch" directory
    as a unique filename.
    
    At completion a list of these filenames will be written out as a filename in
    the GMAT model output_file path.  The model name will be of the form:
        [Mission Name] + '__RunList_[Julian Date-time] + '.batch'
    Example: "AlfanoXfer__RunList_J009_0537.25.batch"
    
    Input:
    A dictionary is used to drive the GMAT resources and parameters written to
    Include macro file 2.  The dictionary is factored into a specialized class derived from
    CModelPov such that additional resources may be added or deleted without change to code.

    The Base classes of this module are crafted to support the Electric Orbit
    Transfer Vehicle study.  Classes must be derived in order to customize the
    GMAT Resources Written to Script.

    The Alfano trajectory is used in the base mission, and a user defined parameter,
    the costate (aka lambda) is updated in concert with the inclination to
    execute the Alfano-Edelbaum yaw control law.  This parameter is likely unique to
    this mission.
    The costate calculation is out of scope but is generated externally as part of the 
    gmatautomation controls module.  The costate is included in the base class as a mission
    parameter.   

    Include file 1 can be extracted from the manually created initial GMAT script, 
    the model design is the responsibility of the GMAT user.  The points of variation
    can be updated in a class derived from CModelPov.
    
    Methods of the class CConfigsheet are called to read the Excel worksheet which
    contains the Model Spec to update in lists called "cases".
              
    The top level GMAT script is intended to be called by the GMAT batch facility,
    therefore each variation of Include file 2 must be matched by a uniquely named
    top level GMAT script.
    
    Each model variation is executed for one or more user-defined Epoch dates, 
    therefore the number of top-level scripts to be generated is the number of 
    configurations in the configuration worksheet times the number of Epochs in the
    epoch list
    
    The model Include file name shall be unique for each different variation.
    The name shall be of the form:
        'case_''[num HET]'X'[power]'_'[payload mass]'_'[epoch]'_'[inclination]
    
    modelgen.py is coded to avoid overwriting an existing model file. The current 
    Julian day and time is suffixed to each filename.

    Module "modelgen.py" shall output each top level file as well as a 
    list of batch filenames.
    
    Input: 
        An Excel worksheet containing the points of variation
    as values for resources in named column ranges.
    
    Output: 
        A series of GMAT #Include files with resources and values, one for each line
        of the input workbook
        A series of uniquely named GMAT top-level model files, one for each line of
        the input workbook.
        A GMAT batch file listing the names of the above model files for execution by a 
        batcher utility.
        
    @remark Change History
        Fri Oct 19 14:35:48 2018, Created
        09 Jan 2019, [CCH] commit to GitHub repository GMAT-Automation, Integration Branch.
        10 Jan 2019, [CCH] Implement the GMAT batch command.
        08 Feb 2019, [CCH] Fix near line 443, include model filename fixed.
        09 Feb 2019, [CCH] Fix is "ReportFile" should be "ReportFile1".
        10 Apr 2019, [CCH] Flow Costates through to model from worksheet.
        16 Apr 2019, c[CCH] onfigspec value formatting moved to fromconfigsheet.py
        26 May 2019, [CCH] Factor out class CGMATParticulars to the gmatlocator module.
        20 Apr 2022, [CCH] reorganized and included in sdist
        26 Apr 2022, [CCH] Version 0.2a1, Buildable package, locally deployable.
        05 May 2022, [CCH] Build version 0.3b3 for PyPI and upload as open source.
        16 Jun 2022, [CCH] Copied from modelgen.py
        16 Jun 2022, [CCH] https://github.com/a093130/GMAT-Automation/issues/1 (refactor)
        05 Dec 2023, [CCH] @version 0.4b0, eliminate Alfano specializations

    @bug https://github.com/a093130/GMAT-Automation/issues
"""
import sys
import os
import re
import time
import logging
import pywintypes as pwin
import xlwings as xw
from shutil import copy as cp
from pathlib import Path
from userexceptions import Ultima
#from gmatautomation import fromconfigsheet as cfg
#from gmatautomation import CGmatParticulars
import fromconfigsheet as cfg
#from specializedmodelpov import CUserSpecializedPov as pov
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

"""
@defgroup Globals
@brief Python global constants are persistent across calls to functions.  The ALLCAPS case
is reserved for identification of globals.
"""
# @{
ROOTPATH = Path(sys.argv[0]).parent.resolve()
STARTPATH = ROOTPATH.parents[1]
DATAPATH = STARTPATH.joinpath('data/')
# @} End of Globals


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
class CModelSpec:
    """ 
        @brief  This class wraps operations on the configsheet to obtain the pov dictionary.
    """
    def __init__(self, path):

        self.model_template = 'ModelMissionTemplate.script'
        self.model_static_res = 'Include_StaticDefinitions.script'
        self.model_miss_def = 'Include_MissionDefinitions.script' 

        self.wbpath = Path(path)

        self.cases = []

    
    def settemplates(self, template:str, staticinclude:str, missioninclude:str):
        """ 
            @brief accessor to change the names of the template and include files.
                
            @param str template [IN] - name of model_template
            @param str staticinclude [IN] - name of model_static_res
            @param str missioninclude [IN] - name of model_miss_def
        """
        if isinstance(template, str):
            self.model_template = template

        if isinstance(staticinclude, str):
            self.model_static_res = staticinclude

        if isinstance(missioninclude, str):
            self.model_miss_def = missioninclude
    

    def get_cases(self):
        """ """
        """ 
            @brief accessor to get the configuration spec from Excel workbook
            
            @details Sets up xlwings and fromconfigsheet.modelspec()
            using the instance wbpath.
            The output of modelspec() is saved in the instance member self.cases[].
        """
        try:
            excel = xw.App()
            excel.visible=False
        
            wingbk = excel.books.open(str(self.wbpath))
            self.cases = cfg.retrievespec(wingbk)
            wingbk.close()

        except OSError as ouch:
            msg = 'Open {0} failed. \nOS error: {1}.'.format(ouch.strerror, ouch.filename)
            print(msg)
            logging.error(msg)

        except pwin.com_error as ouch:
            msg = 'Access to sheet raised Windows com error. {0}, {1}'.format(type(ouch), ouch.args[1])
            print(msg)
            logging.error(msg)
        
        finally:
            excel.quit()


if __name__ == "__main__":
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code.
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
        
     This is the top-level entry point for the GMAT Model Generation. 
    """

    logging.basicConfig(
        filename='./modelgen.log',
        level=logging.INFO,
        format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', 
        datefmt='%d%B%Y_%H:%M:%S')

    logging.info('******************** Model Spec Test Case Started ********************')
    
    qApp = QApplication([])
    
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
        spec = CModelSpec(fname[0])
        spec.get_cases()

        msg = 'The number of cases in the model spec: {0}'.format(len(spec.cases))
        print(msg)
        logging.info(msg)

    except OSError as err:
        logging.error("OS error: {0}".format(err))
    
    except Ultima as u:
        logging.info('%s %s', u.source, u.message)

    except:
        logging.error("Unexpected error:\n", sys.exc_info())

    finally:
        logging.shutdown()
        qApp.quit()
    

