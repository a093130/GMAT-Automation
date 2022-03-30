#! Python
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 22:36:55 2019

@author: colinhelms@outlook.com

@Description:  Sequentially opens .csv formatted GMAT ReportFiles, 
copies pertinent cells and rows into an .xlsx formatted summary file. 
Also incorporates a difference formula in last column, which is fuel residual (or shortage).

The .csv source files are formatted and saved as .xlsx files, which support subsequent
engineering use.

The file name is split upon the '_' separator and each element is written
to the Summary file as 'metadata'.  This permits arbitrary description or parameterization
to be encoded in the file name and carried forward to the summary file.

@Copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@Change Log:
    08 Mar 2019, initial baseline
    17 Mar 2022, Re-factor to support different report formats.
    18 Mar 2022, Make unit test generic.
"""
from genericpath import isfile
import os
import time
import re
import platform
import logging
import traceback
import getpass
import csv
from pathlib import Path
from pathlib import PurePath
import datetime as dt
import xlsxwriter as xwrt
import xlsxwriter.utility as xlut
from gmatlocator import CGMATParticulars
import CleanUpData
import CleanUpReports
import ContactReports
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

dtdict = {'GMAT1':[r'21 Mar 2024 04:52:31.467',
            'd mmm yyyy hh:mm:ss.000',
            r'^\d\d\s[A-S][a-z]+\s\d\d\d\d\s\d\d:\d\d:\d\d.\d\d\d',
            r'%d %b %Y %H:%M:%S.%f']}
""" Dictionary containing specific GMAT date time formats.
    Used for converting datetime strings written to Excel to UT1 dates and then displaying the numerical date in GMAT format using Excel.
    Element is List of date string, Excel cell number format, regular expression syntax, and datetime library format string parameter.
"""

def timetag():
    """ Snapshot a time tag string"""
    return(time.strftime('J%j_%H%M%S',time.gmtime()))
    
def newfilename(pathname, suffix='.txt', keyword=None):
    """ Utility function for a file path operation often done.
        pathname: string representing unix or windows path convention to a file.
        keyword: string to be appended to filename
        suffix.
    """
    filepath = PurePath(pathname)
    filename = filepath.stem

    if keyword:
        newfilename = filename + keyword + suffix
    else:
        newfilename = filename + suffix

    filepath = filepath.parents[0]

    return(filepath/newfilename)


def decimate_spaces(filename):
    """ Read a text file with multiple space delimiters, decimate the spaces and substitute commas.
        Do not replace single spaces, as these are in the time format.
    """
    logging.debug("Decimating spaces in {0}".format(filename))

    rege2spc = re.compile(' [ ]+')
    regeoddspc = re.compile(',[ ]+')
    regecr = re.compile('\s')

    fixedlns = []

    try:
        with open(filename, 'r') as fin:
            lines = list(fin)

            for r, line in enumerate(lines):
                if regecr.match(line) == None:
                    line = rege2spc.sub(',', line)
                    line = regeoddspc.sub('', line)

                    ''' It is better to make a new list than to insert into lines.'''
                    fixedlns.append(line)
                else:
                    """ Skip non-printable lines. """
                    continue
                        
    except OSError as e:
        logging.error("OS error #1in decimate_spaces(): %s reading filename %s", e.strerror, e.filename)
        
        return None
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("File read exception #1 in decimate_spaces(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        
        return None

    filename = newfilename(filename, '.txt', '+nospc')
    ''' Make new filename, don't overwrite the original file.
        The batch procedure splits filenames on '_' so we use '+' instead.
    '''

    try:
        """ Write cleaned up lines to new filename. """
        with open(filename, 'w+') as fout:
            for row, line in enumerate(fixedlns):
                fout.write(line)

    except OSError as e:
        logging.error("OS error #2 in decimate_spaces(): %s writing clean data %s", e.strerror, e.filename)
        
        return None
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Clean data write exception #2 in decimate_spaces(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        
        return None

    return(filename)

def decimate_commas(filename):
    """ Read a malformed csv file, which contains empty fields. Decimate commas. Write a clean file.
        Return the clean filename so that this function can be used as a parameter in 
        lines_from_csv(csvfile).
    """
    logging.debug("Decimating commas in {0}".format(filename))

    fixedlns = []

    regecom = re.compile('[,]+')
    regeolcom = re.compile(',$')

    try:
        with open(filename, 'r') as fin:
            lines = list(fin)

            for row, line in enumerate(lines):
                if line.isprintable:
                    line = regecom.sub(',', line)
                    line = regeolcom.sub('', line)

                    ''' It is better to make a new list than to insert into lines.'''
                    fixedlns.append(line)
                else:
                    """ Skip non-printable lines. """
                    continue

    except OSError as e:
        logging.error("OS error #1 in decimate_commas(): %s reading filename %s", e.strerror, e.filename)
        
        return None
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("File read exception #1 in decimate_commas(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        
        return None

    filename = newfilename(filename, '.csv', '+reduced')
    ''' Make new filename, don't overwrite the original file.
        The batch procedure splits filenames on '_' so we use '+' instead.
    '''

    try:
        with open(filename, 'w+') as fout:
            """ Write cleaned up lines to new filename. """
            for row, line in enumerate(fixedlns):
                fout.write(line)

        return(filename)

    except OSError as e:
        logging.error("OS error #2 in decimate_commas(): %s writing clean data to filename %s", e.strerror, e.filename)
        
        return None
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Clean data write exception #2 in decimate_commas(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        
        return None

def lines_from_csv(csvfile):
    """ Read a well-formed .csv file, or one which contains intentional empty fields. 
        Return a dictionary with row as key and list of lines as elements.
    """
    logging.debug("Extracting lines from report file {0}".format(csvfile))
    
    data = {}
    
    try:
        regecr = re.compile('\s')
        regesp = re.compile(' ')

        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
            lines = list(f)

            for row, line in enumerate(lines):
                line = regesp.sub('', line)
                line = regecr.sub('', line)
                rlist = line.split(',')
                
                data.update({row: rlist})
                
        return data
        
    except OSError as e:
        logging.error("OS error in lines_from_csv(): %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception in lines_from_csv(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
   
def csv_to_xlsx(csvfile):
    """ Read a .csv formatted file, write it to .xlsx formatted file of the same basename. 
        Return the writtenfilename.
        Reference Stack Overflow: 
        https://stackoverflow.com/questions/17684610/python-convert-csv-to-xlsx
        with important comments from:
        https://stackoverflow.com/users/235415/ethan
        https://stackoverflow.com/users/596841/pookie
    """
    logging.debug("Converting report file {0}".format(csvfile))

    fname = (csvfile.stem).split('+')[0]
    """Get rid of the 'nospc' and 'reduced' keywords."""
    xlfile = newfilename(csvfile.parents[0]/fname, '.xlsx')
    """Slice the .csv suffix, append .xlsx suffix, open a new workbook under this name."""

    wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})

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

    regedot = re.compile('\.')
    regecaps = re.compile('[A-Z]')
              
    try:
        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
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
                        """ GMAT uses a lengthy dot notation in headings. We want these to wrap gracefully."""
                        matchcap = regecaps.match(data)
                        sheet.write(row, col, data, cell_wrap)
                    else:
                        """ Set the width of each column for widest data. """
                        if row >= 1: 
                            sheet.set_column(col, col, leng)
                        
                        """ Detect date-time string, a specific re format must be matched, uses dtdict."""
                        if re.search(dtdict['GMAT1'][2], data):

                            gmat_date = dt.datetime.strptime(data, dtdict['GMAT1'][3])

                            sheet.write(row, col, gmat_date, cell_datetime)
                        else:
                            """ Workbook is initialized to treat strings that look like numbers as numbers.
                                Defer application specific number formatting.
                            """
                            sheet.write(row, col, data)

        sheet.freeze_panes('A2')
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

if __name__ == "__main__":
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
    logging.basicConfig(
            filename='./reduce_report.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! Reduce Report Execution Started !!!!!!!!!!")
    
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
 
    xlreport = CleanUpReports()
    dataonly = CleanUpData()
    batchrep = CleanUpReports()

    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open REPORT File. NOT BATCH!', 
                    o_path,
                    filter='text files(*.txt *.csv)')

    logging.info('Input report file is %s', fname[0])

    fbatch = QFileDialog().getOpenFileName(None, 'Open BATCH file.', 
                    o_path,
                    filter='Batch files(*.batch)')

    logging.info('Report batch file is %s', fbatch[0])

    try:
        """ Test Cases. Demonstrates a complete, three-step toolchain.
            Input should be a GMAT Tab delimited Contact File, but shouldn't fail in any case.
            Output files are located in the same path as the original.
        """ 
        """ Test Cases 1: Run through the toolchain to create an Excel File. """
        newfile = xlreport.extend(fname[0])

        logging.info('Test Case 1: cleaned Excel file is %s', newfile)
        print('Test Case 1: cleaned Excel file is %s', newfile)

        """ Test Case 2: Read the same csv file and return a data dictionary instead."""
        dataonly.extend(fname[0])

        logging.info("Test Case 2: First Row of data: \n\t{0}".format(dataonly.data[0]))
        print("Test Case 2: First Row of data: \n{0}".format(dataonly.data[0]))

        """ Test Case 3: Batch Processing of .txt report files. """
        batchrep.do_batch(fbatch[0])

        logging.info("Test Case 3: First file written: {0}".format(batchrep.filelist[0]))
        print("Test Case 3: First file written: {0}".format(batchrep.filelist[0]))

        print("All Test Cases completed successfully.")

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Test Case failed with exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
    
    finally:
        qApp.quit()
    