#! python
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 19 14:35:48 2018

@author: colin helms

@description:
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
    10 Apr 2019, Flow Costates through to model from worksheet.
    16 Apr 2019, configspec value formatting moved to fromconfigsheet.py
    26 May 2019, Factor out class CGMATParticulars to the gmatlocator module.
       
"""
#from modelgen import find_gmat

import sys
import os
import re
import time
import logging
import fromconfigsheet as cfg
from shutil import copy as cp
from gmatlocator import CGMATParticulars
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

model_template = 'ModelMissionTemplate.script'
model_static_res = 'Include_StaticDefinitions.script'
model_miss_def = 'Include_MissionDefinitions.script'
""" These names are design assumptions and should not change. """

            
class CModelSpec:
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
                   
        return self.cases
    
    def get_workbook(self):
        """ Get the instance workbook name. """
        logging.debug('Method get_workbook() called.')
        
        return self.wbpath
   
class CModelWriter:
    """ This class wraps operations to generate the GMAT model include files.
    
    There should be one CModelWriter instance for each elaborated row in configspec.
    """ 
    def __init__(self, spec, outpath): 
        logging.debug('Instance of class ModelWriter constructed.')
        
        self.out_path = outpath
        self.case = {}
        self.nameroot = ''
        self.reportname = ''
        self.inclname = ''
        self.inclpath = ''
        self.model = ''
        self.mission_name = ''
        
        self.case.update(spec)
        
        import modelpov as pov
        
        gmat_vars = pov.getvarnames()
        msn_vars = pov.getrecursives()
        
        if 'EOTV.Epoch' in msn_vars:
            epoch = str(self.case['EOTV.Epoch'])
            
            epochstr = epoch[0:11]               
            """ Clean-up illegal character ':' in nameroot. """
        else:
            epochstr = 'default_epoch'
            
        if 'EOTV.INC' in msn_vars:
            inclination = round(self.case['EOTV.INC'], 2)
        else:
            inclination = 0

        if 'PL_MASS' in gmat_vars:
            payload = round(self.case['PL_MASS'], 0)
        else:
            payload = 0

        if 'SMA_INIT' in msn_vars:
            initsma = round(self.case['SMA_INIT'], 0)
        else:
            initsma = 6578
        
        if 'SMA_END' in msn_vars:
            finalsma = round(self.case['SMA_END'], 0)
        else:
            finalsma = 65781
            
        if 'COSTATE' in msn_vars:
            lamb = round(self.case['COSTATE'], 3)

        if initsma > 0:
            orbit_ratio = round(finalsma/initsma, 2)
        else:
            orbit_ratio = 10
           
        rege=re.compile(' +')
        """ Eliminate one or more blank characters. """
        
        self.mission_name = self.case['ReportFile1.Filename']
        """ Save the mission name passed in ReportFile1.Filename. """
        
        self.nameroot = rege.sub('', self.case['ReportFile1.Filename'] +\
                                 '_' + str(payload) + 'kg' +\
                                 '_' + epochstr +\
                                 '_R' + str(orbit_ratio) +\
                                 '_' + str(inclination) + 'deg' +\
                                 '_' + str(lamb) +\
                                 '_' + time.strftime('J%j_%H%M%S',time.gmtime()))
        """ Generate unique names for the model file output and the reportfile, 
        something like, '16HET8060W_2000.0kg_20Mar2020_28.5_J004_020337'.
        """
            
        self.inclname = 'Include_' + self.nameroot + '.script'
        """ Generated script filename """

        self.reportname = 'Reports/'+ 'Report_' + self.nameroot + '.csv'
        """ GMAT concatenates the value of the OUTPUT_PATH configuration variable
        with the value of 'ReportFile1.Filename'. 
        """
                
        self.case['ReportFile1.Filename'] = "'" + str(self.reportname) + "'"
        self.case['EOTV.Id'] = "'" + str(self.case['EOTV.Id']) + "'"
        self.case['EOTV.Epoch'] = "'" + str(self.case['EOTV.Epoch']) + "'"
        """ GMAT requires single quotes around character lines. 
        This was very hard - beware that GMAT will not execute the model
        if you screw this up.
        """
        p = str(self.out_path)
        
        if p.count('\\') > 1:
            self.modelpath = self.inclpath = p + 'Batch\\'
            self.reportpath = p + 'Reports\\Report_' + self.nameroot + '.csv\n'
        else:
            self.modelpath = self.inclpath = p + 'Batch/'
            self.reportpath = p + 'Reports/Report_' + self.nameroot + '.csv\n'
        
    def get_mission(self):
        """ Get the unique string at the root of all the generated filenames. """
        logging.debug('Method get_mission() called.')
        
        return self.mission_name

    def get_nameroot(self):
        """ Get the unique string at the root of all the generated filenames. """
        logging.debug('Method get_nameroot() called.')
        
        return self.nameroot
    
    def get_reportname(self):
        """ This is the 'ReportFile1.Filename' attribute"""
        logging.debug('Method get_reportname() called.')
        
        return self.reportname
    
    def get_reportpath(self):
        """ This is the absolute path to the output reportfiles.
        Note that the path format may differ from the 'ReportFile1.Filename'
        attribute but will work with report_reduce.py
        """
        logging.debug('Method get_reportpath() called.')
        
        return self.reportpath

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
        
        writefilename = self.get_inclpath() + self.get_inclname()
        
        varset = {'SMA_INIT','SMA_END','MORE','COSTATE', 'PL_MASS'}
        """ User defined variables, special handling required. """
        try:        
            with open(writefilename,'w') as pth:
                for key, value in self.case.items():
                    if key in varset:
                        lcrea = 'Create Variable ' + key + ';\n'
                        pth.write(lcrea)
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
        filename='./modelgen.log',
        level=logging.INFO,
        format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', 
        datefmt='%d%B%Y_%H:%M:%S')

    logging.info('******************** Automation Started ********************')
    
    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open Configuration Workbook', 
                       os.getenv('USERPROFILE'))
    # TODO - a more feature rich GUI.
    
    logging.info('Configuration workbook is %s', fname[0])
    
    spec = CModelSpec(fname[0])
    cases = spec.get_cases()
    nrows = len(cases)

    progress = QProgressDialog("Creating {0} Models ...".format(nrows), "Cancel", 0, nrows)
    progress.setWindowTitle('Model Generator')
    progress.setValue(0)
    progress.show()

    qApp.processEvents()
    
    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    
    src = o_path + model_template
    """ File path for the ModelMissionTemplate.script """
    
    writer_list = []
    """ Persist the output model attributes """
    
    for case in cases:
        """ Initialize an instance of writer for each line of configuration. """
        mw = CModelWriter(case, o_path)
        writer_list.append(mw)

        mw.xform_write()
        """ Write out the include file """
    
    batchlist = []
    reportlist = []
    """ These lists will be written out to the runlist and reportlist batch files. """
    outrow = 0
    for mw in writer_list:
        """ Copy and rename the ModelMissionTemplate for each ModelWriter instance. """
        if progress.wasCanceled():
            break
        
        outrow += 1
        progress.setValue(outrow)
            
        dst = mw.get_inclpath() + 'Batch_' + mw.get_nameroot() + '.script'
        
        static_include = o_path + model_static_res
        mission_include = o_path + model_miss_def
        
        cp(src, dst)
        """ Use shutils to copy source to destination files. """
        
        logging.info('Source model name: %s copied to destination model name: %s.', src, dst )
        
        rege = re.compile('TBR')
        line = ["#Include 'TBR'\n", "#Include 'TBR'\n", "#Include 'TBR'\n"]

#FIX 02/08/2019       
        includp = mw.get_inclpath() + mw.get_inclname()

        rptfile = mw.get_reportpath()
        
        try:                  
            with open(dst,'a+') as mmt:   
                """ Append the #Include macros to the destination filename. """
                line[0] = rege.sub(static_include, line[0])                        
#FIX 02/08/2019
#               line[1] = rege.sub(dst, line[1])
                line[1] = rege.sub(includp, line[1])
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
                
                reportlist.append(rptfile)
                """ Script reduce_report.py will summarize the contents of the named reports. """
                
        except OSError as err:
            logging.error("OS error: ", err.strerror)
            sys.exit(-1)
        except:
            logging.error("Unexpected error:\n", sys.exc_info())
            sys.exit(-1)

    batchfilename = \
    o_path + 'RunList_' + time.strftime('J%j_%H%M.%S', time.gmtime()) + '.batch'
    """ Write out the batch file, containing the names of all the top level models. """
           
    batchrptname = \
    o_path + 'ReportList_' + time.strftime('J%j_%H%M.%S', time.gmtime()) + '.batch'
    """ Write out the batch file, containing the names of all the top level models. """

    try:
        with open(batchfilename,'w') as bf:
            bf.writelines(batchlist)
            
        with open(batchrptname,'w') as rf:
            rf.writelines(reportlist)
        
    except OSError as err:
        logging.error("OS error: {0}".format(err))
    except:
        logging.error("Unexpected error:\n", sys.exc_info())
    finally:
        logging.info('GMAT batch file creation is completed.')
        logging.shutdown()
        qApp.quit()
    

