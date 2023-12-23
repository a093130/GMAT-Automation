#! python
# -*- coding: utf-8 -*-
"""
    @file modelgen.py

    @brief: uses modelpov and modelspec to write out batch script files. 

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
        05 Dec 2023, [CCH] @version 0.4b0, refactored to eliminate Alfano specializations

    @bug https://github.com/a093130/GMAT-Automation/issues
"""
import sys
import os
import re
import time
import logging
from shutil import copy as cp
from pathlib import Path
from modelspec import CModelSpec
from gmatlocator import CGmatParticulars
from specializedmodelpov import CUserSpecializedPov
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

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
  
class CModelWriter:
    """ 
    @brief  This class wraps operations to generate the dynamic GMAT model include files.
    
    @details There should be one CModelWriter instance for each elaborated row in configspec (a "case").

    """ 
    def __init__(self, spec:dict, outpath:Path):
        """ 
            @brief instance initializer for CModelWriter
            
            @details detailed description
                
            @param <dict>  spec [IN] - One case instance from the cases list.
            @param <type>  outpath [IN] - path where the model template file and include files are written.

            @return condition1 = value1
            @return condition2 = value2

        """
        
        self.model_template = 'ModelMissionTemplate.script'
        self.model_static_res = 'StaticDefinitions.script'
        self.model_miss_def = 'MissionDefinitions.script'
        
        self.out_path = Path(outpath)

        """ This is a copy of the current case corresponding one row of the configsheet. """
        self.case = dict()
        self.case.update(spec)

        self.pov = CUserSpecializedPov()

        self.model = self.pov.model_name
        self.mission_name = self.pov.mission_name

        self.nameroot = self.pov.mknameroot(self.case)

        self.pov.getreportfile(self.case)
        self.pov.getdebugreport(self.case)
        self.pov.getkernelname(self.case)

        """ resourcedefs contains the GMAT Report and kernel filename definitions after the above three calls """
        self.resourcedefs = self.pov.resources

        self.body = self.pov.body
        self.naifdef = self.pov.naifdef

        """ Generate include filename """
        self.incl_path = self.out_path / 'Batch'
        self.model_path = self.incl_path

        self.inclname = self.nameroot + '.include'
        self.includefile = self.incl_path / self.inclname
        

    def xform_write(self):
        """ 
            @brief Extract each key:value pair, form GMAT syntax, write it to the outpath.
            
            @details xform_write creates an include file to contain user variables and values as read from the
            instance case.
                
            @param <type> name [IN] - description
            @param <type> name [OUT] - description

            @return condition1 = value1
            @return condition2 = value2

        """         
        writefilename = self.includefile.as_posix()
        
        """ User defined variables, special handling required. """     
        with open(writefilename,'w') as pth:  
            
            for key, value in self.pov.resources.items():
                line = 'GMAT ' + value + ';\n'
                pth.write(line)

            for key, value in self.case.items():
                
                """ If a GMAT user variable, the key in this case may or may not already 
                be created in the GMAT StaticDefinitions file.  If a dynamic variable 
                it should be definied in specializedmodelpov.py as an element of varset.
                """
                if len(self.pov.varset) > 0 and key in self.pov.varset:
                    lscreate = 'Create Variable ' + key + ';\n'
                    pth.write(lscreate)
                    
                line = 'GMAT ' + str(key) + ' = ' + str(value) + ';\n'
                pth.write(line)
                                
            logging.info('ModelWriter has written include file %s.', writefilename)
                
# @} End of Class Group 


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

    logging.info('******************** Automation Model Generation ********************')
    
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
    
    spec = CModelSpec(fname[0])

    """ Initialize CModelSpec """
    spec.get_cases()
    nrows = len(spec.cases)

    msg = 'The number of cases in the model spec: {0}'.format(nrows)
    print(msg)
    logging.info(msg)

    timetag = time.strftime('J%j_%H%M.%S', time.gmtime())
                            
    progress = QProgressDialog("Creating {0} Models ...".format(nrows), "Cancel", 0, nrows)
    progress.setWindowTitle('Model Generator')
    progress.setValue(0)
    progress.show()
    qApp.processEvents()
    
    gmat_paths = CGmatParticulars()
    o_path = gmat_paths.get_output_path()

    """ Temporary fix to remove 'OutpuPath=' """
    #o_path = str(gmat_paths.get_output_path())
    #path = o_path.split('=')[1] 
    #o_path = Path(path)

    """ Initialize an instance of writer for each case. """
    writer_list = []
    for case in spec.cases:
        mw = CModelWriter(case, o_path)
        writer_list.append(mw)

        mw.xform_write()
        """ Write out the include file """
    
    """ These lists will be written out to the batch files for runlist and reportlist. """
    batchlist = []
    reportlist = []
    
    """ File path for the ModelMissionTemplate.script is on the GMAT output path. """
    src = o_path / 'Batch' / spec.model_template

    regerpt = re.compile('Report')

    outrow = 0
    for mw in writer_list:
        """ Copy and rename the ModelMissionTemplate for each ModelWriter instance. """
        if progress.wasCanceled():
            break
        else:
            outrow += 1
            progress.setValue(outrow)
            
        dst = (mw.model_path / mw.nameroot).with_suffix('.script')
        
        static_include = o_path / 'Batch' / spec.model_static_res
        mission_include = o_path / 'Batch' / spec.model_miss_def
        
        """ Use shutils to copy source to destination files. """
        cp(src, dst)
        
        msg = 'Source model name: {0} copied to destination model name: {1}.'.format(src, dst)
        logging.info(msg)

        rege = re.compile('TBR')
        line = ["#Include 'TBR'\n", "#Include 'TBR'\n", "#Include 'TBR'\n"]
        regesquote = re.compile("'")

        res = mw.pov.resources
        rpt = []
        for resource, name in res.items():
            if regerpt.search(resource):
                filename = name.split('=')[1]

                """ Remove the GMAT path quote nightmare. """
                filename = regesquote.sub("",filename)
                filepath = o_path / filename

                """ Downstream processing only works on csv files. """
                if filepath.suffix == '.csv':
                    rpt.append(filepath) 
        
        numreports = len(rpt)

        try:                  
            with dst.open(mode='a+') as mmt:   
                """ Append the #Include macros to the destination filename. """
                line[0] = rege.sub(static_include.as_posix(),line[0])                      
                line[1] = rege.sub(mw.includefile.as_posix(), line[1])
                line[2] = rege.sub(mission_include.as_posix(), line[2])
                """ Order of these includes is important. """
                
                for edit in line:                         
                    mmt.write(edit)
                                                                                
                batchfile = str(dst) + '\n'
                batchlist.append(batchfile)
                """ GMAT will batch execute a list of the names of top-level models. """
                
                for rnum in range(0,numreports):
                    reportfile = str(rpt[rnum]) + '\n'
                    reportlist.append(reportfile)
                    """ Script reduce_report.py will summarize the contents of the named reports. """
                
        except OSError as err:
            logging.error("OS error: ", err.strerror)
            sys.exit(-1)
        except:
            logging.error("Unexpected error:\n", sys.exc_info())
            sys.exit(-1)


    """ Write out the batch file, containing the names of all the top level models. 
    These are executed by GMAT on the command line.
    """
    runfile = Path('RunList_' + timetag).with_suffix ('.batch')
    batchfilename = o_path / runfile
    
    """ Write out the batch file, containing the names of all the reports. 
    These are used in the batch report formatting from csv to Excel.
    """
    rptfile = Path('ReportList_' + timetag).with_suffix ('.batch')
    batchrptname = o_path / rptfile
    
    try:
        with batchfilename.open(mode='w') as bf:
            bf.writelines(batchlist)
            
        with batchrptname.open(mode='w') as rf:
            rf.writelines(reportlist)
        
    except OSError as err:
        logging.error("OS error: {0}".format(err))
    except:
        logging.error("Unexpected error:\n", sys.exc_info())
    finally:
        logging.info('GMAT batch file creation is completed.')
        logging.shutdown()
        qApp.quit()
    

