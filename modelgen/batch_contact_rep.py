#! Python
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 22:36:55 2019

@author: colinhelms@outlook.com

@description:  Executes programs in batch mode to reduce data.

Do the minimum processing to write a clean Excel workbook for each csv file
in the batch.

@Copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@Change Log:
    08 Mar 2019, initial baseline
    18 MAR 2022, refactored __main__ into new procedure specifically for batches of GMAT Contact Reports.
"""
import os
import time
import re
import platform
import logging
import traceback
import getpass
import csv
import xlsxwriter as xwrt
import xlsxwriter.utility as xlut
from pathlib import Path
from pathlib import PurePath
from gmatlocator import CGMATParticulars
from reduce_report import CleanUpData
from reduce_report import CleanUpReports
from reduce_report import dtdict
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

class CombineReports(CleanUpData):
    """" Application specialization to form one combined report. """
    def __init__(self):
        self.filelist = {}

    def extend(self, rpt):
        """ Call parent class do_batch() and the CombineReports.extend function will be called.
            This version of extend will identify the report type and call the appropriate methods.
        """
        regetarget = re.compile('Target: ')
        regeobsrvr = re.compile('Observer: ')
        regenumevt = re.compile('Number of events: ')
        regeheading = re.compile('Start Time')
        regesatnum = re.compile('LEOsat')
        regetime = re.compile(dtdict['GMAT1'][2])
        """ Regular Expression Match patterns to identify files data items. """

        datadict = CleanUpData.extend(rpt)
        for k, v in datadict.items:
            """ Classify what type of file this is. """
            mtarg = regetarget.match(v[0])
            mobsv = regeobsrvr.match(v[0])
            mevts = regenumevt.match(v[0])
            satn = regesatnum.match(v[0])
            headg = regeheading.match(v[0])
            ttg = regetime.match(v[0])
            
            if mtarg:
                """ This is a SIGHT Report """
                
            elif mobsv:
                """ This is a SIGHT Report """

            elif mevts:
                """ This is a SIGHT Report """

            elif satn:
                """ This is a Link Report """

            elif headg:
                """ This is a Link Report"""

            elif ttg:
                """ This is a Link Report"""

            else:
                """ Blank line or unknown type"""
                pass

                
                
                

        



        
    def merge(self, batchfile):

        try:
            CleanUpData.do_batch(batchfile)

            return(None)

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])

if __name__ == "__main__":
    """ Retrieve the formatting batch file, open and format each .csv file listed """
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
    logging.basicConfig(
            filename='./batch_contact_rep.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! Batch Report Format Execution Started !!!!!!!!!!")
    
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
    cleanup = CleanUpReports

    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open report batch file.', 
                       o_path(),
                       filter='Batch files(*.batch)')

    logging.info('Report batch file is %s', fname[0])

    try:
        cleanup.do_batch(fname[0])
        """ Base class creation of Excel files from GMAT output txt and csv files. """

        """ Derived class combination of Link Report with SIGHT Report. """
                                                        
    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
    
    finally:
        logging.info ('!!!!! Batch Processing Completed !!!!!')
        print('Batch Processing Completed.')
        qApp.quit()
        exit()
    