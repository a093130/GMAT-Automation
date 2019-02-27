# -*- coding: utf-8 -*-
"""
Created on Fri Oct 19 14:35:48 2018

@author: colinhelms@outlook.com

@Description:
    This script produces GMAT model Include files containing 
    variants of model resource values and parameters.  The module supports 
    batch processing of different mission scenarios in which the spacecraft
    configuration and/or initial orbital elements vary.
    
    The approach utilizes the GMAT 2018a #Include macro, which loads resources and 
    script snippets from external files.  The script creates Include files whose
    parameter and resource values vary in accordance with an Excel workbook.
    
    A top level GMAT script template must exist in the GMAT model directory.
    This script template will be copied and modified by modelgen.py.
    The script template must contain the GMAT create statements for GMAT resources.
      
    The script template must contain three #Include statements as follows:
    (1) a GMAT script file defining the static resources, those that don't change per run
    (2) An include file defining those variables written by "modelgen.py".
    (3) An include file containing the GMAT Mission Sequence definition.
    
    The second #Include statement will be modified by modelgen.py to include a
    uniquely named filepath and the entire template copied to a "batch" directory
    as a unique filename.
    
    At completion as list of these filenames will be written out as a filename in
    the GMAT model output_file path.  The model name will be of the form:
        [Mission Name] + '__RunList_[Julian Date-time] + '.batch'
    Example: "AlfanoXfer__RunList_J009_0537.25.batch"
    
    Input:
    A dictionary is used to drive the actual resources and parameters written to
    Include macro file 2.  The dictionary is factored into "modelpov.py" such
    that additional resources may be added or deleted without change to code.
    
    Include file 1 must be extracted from the initial model file written by GMAT, 
    the model design is the responsibility of the GMAT user.  The points of variation
    must also be updated in "modelpov.py" for the case of a new model.
    
    The external module "fromconfigsheet.py" is called to read excel worksheet
    to update the values of the dictionary.
 
    The Alfano trajectory is used in the current model mission, and a user defined
    parameter, the costate, must be updated in concert with the inclination to
    execute the Alfano-Edelbaum yaw control law.
    The costate calculation is out of scope as of this version [TODO]. So is 
    specified in the configsheet workbook.
        
    Notes:
       1. To model the return trip of the reusable vehicle, two include files
            must be generated, one with payload mass included, one without.
       2. Dry mass varies with the vehicle power and thrust.
       3. Efficiency, thrust and Isp vary with the selected thruster set-points.
       4. In order to cover the range of eclipse conditions the EOTV Epoch is
            varied for the four seasons:
                20 Mar 2020 03:49 UTC
                20 Jun 2020 21:43 UTC
                22 Sep 2020 13:30 UTC
                21 Dec 2020 10:02 UTC
                The epoch date is specified in the configsheet workbook.
       5. Propellant is given as an initial calculation, then the actual value
            from a run is substituted and the model rerun until convergence.
            This iterative process is executed by "modeliterate.py", however
            the code and data architecture must be factored for reuse.
       6. The output is a .csv file and it's name must not only be varied for each
            model, but also for each iteration of the model. An example is,
            'ReportFile_AlfanoXfer_20Jun2020_28.5i_64x5999_16T_Run3.csv'.
       7. The external module "modeliterate.py" must be able to identify and open
            the output file.
       8. The OrbitView viewpoint is superfluous in most cases, since model execution
            is intended to be in batch mode for this system.  However, if graphic
            output is desired, the convention is to vary the viewpoint with the 
            Starting Epoch as follows:
                20 Mar 2020 03:49 UTC, [ 80000 0 20000 ]
                20 Jun 2020 21:43 UTC, [ 0 80000 20000 ]
                22 Sep 2020 13:30 UTC, [ 0 -80000 20000 ]
                21 Dec 2020 10:02 UTC, [ -80000 0 20000 ]
                   
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
    as values for resources in columns with headings as above.
    
    Output: 
        A series of GMAT #Include files with resources and values, one for each line
        of the input workbook
        A series of uniquely named GMAT top-level model files, one for each line of
        the input workbook.
        A GMAT batch file listing the names of the above model files.
    
    TODO: the initial baseline depends on the exact spacecraft and hardware
    created in the top level template.  Four of these represent points of 
    variation in general:
        Create Spacecraft EOTV;
        Create ElectricThruster HET1;
        Create SolarPowerSystem EOTVSolarArrays;
        Create ElectricTank RAPTank1;
    Furthermore, there may be multiple ReportFile creates under various names.
    These instance dependencies can be avoided by reading and interpreting ModelMissionTemplate.script.
        
@Change Log:
    08 Jan 2019, initial baseline
    09 Jan 2019, Integration branch, ReportFiles are written to 'Report/' directory.
    10 Jan 2019, NASA did not implement the GMAT batch command.
    08 Feb 2019, Fix near line 443, include model filename fixed.
    09 Feb 2019, Fix near lines 249, 288, 300, 307 357: is "ReportFile" 
        should be "ReportFile1".
       
"""
import sys
import os
import re
import time
import logging
from shutil import copy as cp
from pathlib import Path
from PyQt5.QtWidgets import(QApplication, QFileDialog)
import fromconfigsheet as cfg

model_template = 'ModelMissionTemplate.script'
model_static_res = 'Include_StaticDefinitions.script'
model_miss_def = 'Include_MissionDefinitions.script'

class GMAT_Particulars:
    """ This class initializes its instance with the script output path taken from
    the gmat_startup_file.txt.
    
    TODO: make this class inherit from gmat_batcher.GMAT_Path
    """
    def __init__(self):
        logging.debug('Instance of class GMAT_Particulars constructed.')
        
        self.p_gmat = os.getenv('LOCALAPPDATA')+'\\GMAT'
        self.startup_file_path = ''
        self.output_path = ''
                
    def get_output_path(self):
        """ The path defined for all manner of output in gmat_startup_file.txt """
        logging.debug('Method get_output_path() called.')
        
        p = str(self.output_path)
        
        if p.count('\\') | p.count('/') <= 1:
            self.get_startup_file_path()
        
        rege = re.compile('^OUTPUT_PATH')
        
        try:
            with open(self.startup_file_path) as f:
                """ Extract path string text assigned to OUTPUT_PATH in file. """
                for line in f:
                    if rege.match(line):
                        self.output_path = line
                
        except OSError as err:
            logging.error("OS error: ", err.strerror)
            sys.exit(-1)
        except:
            logging.error("Unexpected error:\n", sys.exc_info())
            sys.exit(-1)

        rege = re.compile(r'^OUTPUT_PATH\s*= ')
        """ Clean-up the path string """
        self.output_path = rege.sub('', self.output_path)
        
        rege = re.compile('\n')
        """ Clean-up newline at the end of each line in a file. """
        self.output_path = rege.sub('', self.output_path)

        logging.info('The GMAT output path is %s.', self.output_path)
        
        return self.output_path
        
    def get_startup_file_path(self):
        """ Convenience function which searches for gmat_statup_file.txt. """
        logging.debug('Method get_startup_file_path() called.')
        
        p = Path(self.p_gmat)
        
        gmat_su_paths = list(p.glob('**/gmat_startup_file.txt'))
        
        self.startup_file = gmat_su_paths[0]
        """ Initialize startup_file path. """
        
        for pth in gmat_su_paths:
            """ Where multiple gmat_startup_file instances are found, use the last modified. """          
            old_p = Path(self.startup_file)
            old_mtime = old_p.stat().st_mtime
            
            p = Path(pth)
            mtime = p.stat().st_mtime

            if mtime - old_mtime > 0:
                self.startup_file_path = pth
            else:
                continue

        logging.info('The GMAT startup file is %s.', self.startup_file_path)
        
        return self.startup_file_path
    
class ModelSpec:
    """ This class wraps operations on the configsheet to obtain the pov dictionary."""
    def __init__(self, wbname):
        logging.debug('Instance of class ModelSpec constructed.')
        
        self.wbpath = wbname
        self.cases = []
           
    def get_cases(self):
        """ Access the initialized workbook to get the configuration spec """
        logging.debug('Method get_cases() called.')
        
        try:
            self.cases = cfg.modelspec(self.wbpath)

        except cfg.Ultima as u:
            logging.error('Call to modelspec failed. In %s, %s', u.source, u.message)
            
        rege_comma = re.compile(',+')
        rege_utc = re.compile(' UTC')
        
        for case in self.cases:
            """ Fix GMAT syntax incompatibilities and inconsistencies. 
            TODO: These edits here defeat the intent of factoring the GMAT
            resource dictionary into modelpov.py.  Perhaps the value formatting
            ought to be performed in fromconfigsheet.py using the workbook table headings.
            """
            case['ReportFile1.Filename'] = str(case['ReportFile1.Filename'])
            case['EOTV.Epoch'] = str(rege_utc.sub('', case['EOTV.Epoch']))
            case['DefaultOrbitView.ViewPointVector'] = \
                rege_comma.sub('', repr(case['DefaultOrbitView.ViewPointVector']))
        
        return self.cases
    
    def get_workbook(self):
        """ Get the instance workbook name. """
        logging.debug('Method get_workbook() called.')
        
        return self.wbpath
   
class ModelWriter:
    """ This class wraps operations to generate the GMAT model include files. """    
    def __init__(self, spec, outpath): 
        logging.debug('Instance of class ModelWriter constructed.')
        
        self.out_path = outpath
        self.case = {}
        self.nameroot = ''
        self.reportname = ''
        self.inclname = ''
        self.inclpath = ''
        self.model = ''
        
        self.case.update(spec)
        
        epoch = str(self.case['EOTV.Epoch'])
        epoch = epoch[0:12]
        """ Clean-up illegal filename charcter in EOTV.Epoch. """
        
        payload = self.case['EOTV.Id']
        """ Swap configuration Id (part of ReportFile name) into EOTV.Id.
        GMAT has no resource for payload mass, so EOTV.Id is used to carry it.
        TODO: use of the GMAT resource name here defeats the intent of factoring
        them into modelpov.py.  Using the ReportFile name to carry the spacecraft
        id is a kludge and should be corrected.  Suggest putting the sid into 
        the workbook on the Mission Parameters tab.
        """
        sid = self.case['ReportFile1.Filename']
        self.case['EOTV.Id'] = sid
        

        rege=re.compile(' +')
        """ Eliminate one or more blank characters. """
        
        self.nameroot = rege.sub('', self.case['ReportFile1.Filename'] +\
                                 '_' + str(payload) + 'kg' +\
                                 '_' + epoch +\
                                 '_' + str(self.case['EOTV.INC']) +\
                                 '_' + time.strftime('J%j_%H%M%S',time.gmtime()))
        """ Generate unique names for the model file output and the reportfile, 
        something like, '16HET8060W_2000.0kg_20Mar2020_28.5_J004_020337'.
        TODO: use of the literal 'ReportFile1' defeats the factorization of GMAT
        Resource names into modelpov.py.
        """       
        self.reportname = 'ReportFile1_'+ self.nameroot + '.csv'
        self.inclname = 'Include_' + self.nameroot + '.script'

        p = str(self.out_path)
        if p.count('\\') > 1:
            self.modelpath = self.inclpath = p + 'Batch\\'
        else:
            self.modelpath = self.inclpath = p + 'Batch/'
        
    def get_nameroot(self):
        """ Get the unique string at the root of all the generated filenames. """
        logging.debug('Method get_nameroot() called.')
        
        return self.nameroot
    
    def get_reportname(self):
        """ This is the 'ReportFile1.Filename' attribute"""
        logging.debug('Method get_reportname() called.')
        
        return self.reportname
    
    def get_inclname(self):
        """ Get the saved model name """
        logging.debug('Method get_inclname() called.')
        
        return self.inclname
           
    def get_inclpath(self):
        """ Get the path of the written inclfile for this instance. """
        logging.debug('Method get_inclfile() called.')
        
        return self.inclpath

    def set_modelfile(self, path):
        """ This is the top level model that includes the generated inclfile. """
        logging.debug('Method set_modelfile() called with path %s.', path)
        
        self.model = path
    
    def get_modelfile(self):
        """ This is the top level model that includes the generated inclfile. """
        logging.debug('Method get_modelfile() called.')
        
        return self.model

    def xform_write(self):
        """ Extract each key:value pair, form GMAT syntax, write it to the outpath. """
        logging.debug('Method xform_write() called.')
        
        self.reportname = self.out_path + 'Reports/' + self.reportname
        """ GMAT requires an absolute path for the ReportFile output. 
        TODO: Use of the literal 'ReportFile1' defeats the factorization of GMAT
        Resource names into modelpov.py.
        """        
        self.case['ReportFile1.Filename'] = "'" + str(self.reportname) + "'"
        self.case['EOTV.Id'] = "'" + str(self.case['EOTV.Id']) + "'"
        self.case['EOTV.Epoch'] = "'" + str(self.case['EOTV.Epoch']) + "'"
        """ GMAT requires the single quotes around character lines """
        
        writefilename = self.get_inclpath() + self.get_inclname()
        
        try:        
            with open(writefilename,'w') as pth:
                for key, value in self.case.items():
                    line = 'GMAT ' + str(key) + ' = ' + str(value) + ';\n'
                    pth.write(line)
                                    
                logging.info('ModelWriter has written include file %s.', writefilename)
                
        except OSError as err:
            logging.error("OS error: ", err.strerror)
            sys.exit(-1)
        except:
            logging.error("Unexpected error:\n", sys.exc_info())
            sys.exit(-1)

if __name__ == "__main__":
    """ This script is the top-level entry point for the GMAT Automation system. """
    logging.basicConfig(
        filename='./appLog.log',
        level=logging.INFO,
        format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', 
        datefmt='%d%B%Y_%H:%M:%S')

    logging.info('******************** Automation Started ********************')
    
    app = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open Configuration Workbook', 
                       os.getenv('USERPROFILE'))
        
    logging.info('Configuration workbook is %s', fname[0])
    
    spec = ModelSpec(fname[0])
    cases = spec.get_cases()
    
    gmat_paths = GMAT_Particulars()
    o_path = gmat_paths.get_output_path()
    
    src = o_path + model_template
    """ File path for the ModelMissionTemplate.script """
    
    writer_list = []
    """ Persist the output model attributes """
    
    for case in cases:
        """ Initialize an instance of writer for each line of configuration. """
        mw = ModelWriter(case, o_path)
        writer_list.append(mw)

        mw.xform_write()
        """ Write out the include file """
    
    batchlist = []
    """ This list will be written out to the batch file. """
    
    for mw in writer_list:
        """ Copy and rename the ModelMissionTemplate for each ModelWriter instance. """
        try:
            mission_name = cfg.mission(fname[0]) + '_'
            
        except Exception as e:
            """ Not a critical error. """
            logging.warn('Non-critical error accessing cfg.mission(): /n/t%s', e.__str__())
            mission_name = 'Batch'

        dst = mw.get_inclpath() + mission_name + mw.get_nameroot() + '.script'
        
        static_include = o_path + model_static_res
        mission_include = o_path + model_miss_def
        
        cp(src, dst)
        """ Use shutils to copy source to destination files. """
        
        logging.info('Source model name: %s copied to destination model name: %s.', src, dst )
        
        rege = re.compile('TBR')
        line = ["#Include 'TBR'\n", "#Include 'TBR'\n", "#Include 'TBR'\n"]

#FIX 02/08/2019       
        incl = mw.get_inclpath() + mw.get_inclname()
        
        try:                  
            with open(dst,'a+') as mmt:   
                """ Append the #Include macros to the destination filename. """
                line[0] = rege.sub(static_include, line[0])                        
#FIX 02/08/2019
#               line[1] = rege.sub(dst, line[1])
                line[1] = rege.sub(incl, line[1])
                line[2] = rege.sub(mission_include, line[2])
                """ Order of these includes is important. """
                
                for edit in line:
                    try:                           
                        mmt.write(edit)
                        
                        logging.info('Edit completed.')
                        
                    except OSError as err:
                        logging.error("OS error %s on writing %s.", err.strerror, edit)
                        sys.exit(-1)
                    except:
                        logging.error("Unexpected error:\n", sys.exc_info())
                        sys.exit(-1)
                                                                                
                batchfile = str(dst) + '\n'
                batchlist.append(batchfile)
                """ GMAT will batch execute a list of the names of top-level models. """
                            
        except OSError as err:
            logging.error("OS error: ", err.strerror)
            sys.exit(-1)
        except:
            logging.error("Unexpected error:\n", sys.exc_info())
            sys.exit(-1)

    batchfilename = \
    o_path + mission_name + '_RunList_' + time.strftime('J%j_%H%M.%S', time.gmtime()) + '.batch'
    """ Write out the batch file, containing the names of all the top level models. """
           
    try:
        with open(batchfilename,'w') as bf:
            bf.writelines(batchlist)
                
    except OSError as err:
        logging.error("OS error: {0}".format(err))
        sys.exit(-1)
    except ValueError:
        logging.error("Could not append time to a filename %s.", batchfilename)
        sys.exit(-1)
    except:
        logging.error("Unexpected error:\n", sys.exc_info())
        sys.exit(-1)
    
    logging.info('GMAT batch file creation is completed.')
    logging.shutdown()
    
    
