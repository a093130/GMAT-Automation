#! Python
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 22:36:55 2019

@author: colinhelms@outlook.com

@description:  module container for class definition ContactReports.


@Copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@changes:
    08 Mar 2019, initial baseline
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
import CleanUpData
import CleanUpReports
from reduce_report import dtdict
import reduce_report as rr
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

class ContactReports(CleanUpData):
    """" Application specialization to form one combined report. """
    def __init__(self):
        self.filelist = {}

    def extend(self, rpt):
        """ Call parent class do_batch() and the CombineReports.extend function will be called.
            This specialization of extend will build a dictionar to identify the report type 
            and call specialized methods to build a contact report by combining GMAT data
            which supports link budget and camera Field of View (FOV) calculations in Excel.
        """
        regetarget = re.compile('Target: ')
        regeobsrvr = re.compile('Observer: ')
        regenumevt = re.compile('Number of events: ')
        regeheading = re.compile('Start Time')
        regesatnum = re.compile('LEOsat')
        regetime = re.compile(dtdict['GMAT1'][2])
        """ Regular Expression Match patterns to identify files data items. """

        nospc = rr.decimate_spaces(rpt)
        nospc = Path(nospc)

        reduced = rr.decimate_commas(nospc)
        if nospc.exists():
            nospc.unlink()

        fname = (csvfile.stem).split('+')[0]
    """Get rid of the 'nospc' and 'reduced' keywords."""
    xlfile = newfilename(csvfile.parents[0]/fname, '.xlsx')
    """Slice the .csv suffix, append .xlsx suffix, open a new workbook under this name."""

    wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
    """  
    It may seem inefficient to create a .xlsx copy of the .csv file, but the Excel copy may be used for
    analysis of data items not included in the summary, e.g. thrust and beta angle history.
    """
    cell_heading = wb.add_format({'bold': True})
    cell_heading.set_align('center')
    cell_heading.set_align('vcenter')
    cell_heading.set_text_wrap()

    cell_wrap = wb.add_format({'text_wrap': True})
    cell_wrap.set_align('vcenter')

    cell_4plnum = wb.add_format({'num_format': '0.0000'})
    cell_4plnum.set_align('vcenter')

    cell_datetime = wb.add_format({'num_format': dtdict['GMAT1'][1]})
    cell_datetime.set_align('vcenter')
    
    sheet = wb.add_worksheet('Report')

        try:
            with open(reduced, 'rt', newline='', encoding='utf8') as f:
                reader = csv.reader(f, quoting=csv.QUOTE_NONE)

                lengs = []

                for row, line in enumerate(reader):
                    for col, data in enumerate(line):
                        leng = len(data) + 1

                        if len(lengs) < col+1:
                            lengs.append(leng)
                        else:
                            lengs[col] = leng

                        if row == 0:
                            data = regedot.sub(' ', data)
                            """ GMAT uses a lengthy dot notation in headings. We want these to wrap gracefully. """
                            sheet.write(row, col, data, cell_wrap)
                        else:
                            """ Set the width of each column for widest data. """
                            if row >= 1: 
                                sheet.set_column(col, col, leng)
                            
                            """ Detect date-time string, a specific re format must be matched, uses dtdict. """
                            if re.search(dtdict['GMAT1'][2], data):

                                gmat_date = dt.datetime.strptime(data, dtdict['GMAT1'][3])

                                sheet.write(row, col, gmat_date, cell_datetime)
                            else:
                                """ Workbook is initialized to treat strings that look like numbers as numbers.
                                    Defer application specific number formatting.
                                """
                                sheet.write(row, col, data)

            #sheet.freeze_panes('A2') #This is too specific and should be deferred
            return str(xlfile)

        except OSError as e:
            logging.error("OS error in csv_to_xlsx(): %s for filename %s", e.strerror, e.filename)
            return None

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exceptionin csv_to_xlsx(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
            return None
        
    finally:
        wb.close()
            
            if mtarg:
                """ This is a SIGHT Report """
                

            elif satn:
                """ This is a Link Report """


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
    