#! Python
# -*- coding: utf-8 -*-
"""
@Description:  Sequentially opens .csv formatted GMAT ReportFiles, removes whitespace 
copies pertinent rows and comma delimited fields into an .xlsx file.

@author: colinhelms@outlook.com

@Copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@Change Log:
    Created on Fri Mar 8 2019
    17 Mar 2022, Re-factor to support different report formats.
    29 Mar 2022, Modify Unit Tests to execute derived classes.
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
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

dtdict = {'GMAT1':[r'21 Mar 2024 04:52:31.467',
            'd mmm yyyy hh:mm:ss.000',
            r'^\d\d\s[A-S][a-z]+\s\d\d\d\d\s\d\d:\d\d:\d\d.\d\d\d',
            r'%d %b %Y %H:%M:%S']}
""" Dictionary containing specific GMAT date time formats.
    Used for converting datetime strings written to Excel to UT1 dates and then displaying the numerical date in GMAT format using Excel.
    Element is List of date string, Excel cell number format, regular expression syntax, and datetime library format string parameter.
"""

""" Regular expressions that are used in common. """
regedot = re.compile(r'\.')
regecr = re.compile(r'^\s')
regesp = re.compile(' ')
regecom = re.compile(r'[,]+')
regecamel = re.compile(r'(?<=[a-z])[A-Z]')
regedecimal = re.compile(r'-*[0-9]+\.[0-9]+')
regetime = re.compile(dtdict['GMAT1'][2])

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

def heading_row(data):
    """ Eliminate extraneous characters and break strings on caps.

        Obviously this function should be called for the heading row only,
        which is usually row 0.

        Returns the formatted data and max count of characters.
    """

    data = regedot.sub(' ', data)
    """ Eliminate dot notation. """
    try:

        miter = regecamel.finditer(data)
        """ Break on camel case. """
        if miter:
            mlist = list()
            for m in miter:
                """ Unfortunately miter is not reversible, so form a list. """
                mlist.append(m)
            
            for m in reversed(mlist):
                """ Find the caps position and insert a space. 
                    Uses built-in function reversed() because a space added at the left 
                    causes rightward positions to increment.
                """
                capspos = m.span()[0]
                data = data[0:capspos] + ' ' + data[capspos:len(data)]
            
            counts = list()
            for item in data.split(' '):
                counts.append(len(item))
            
        return data, max(counts)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s in heading_row():\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception', e.__doc__, ' heading_row():\n', lines[0], '\n', lines[1], '\n', lines[-1])


def decimate_spaces(filename):
    """ Read a text file with multiple space delimiters, decimate the spaces and substitute commas.
        Do not replace single spaces, as these are in the time format.
    """
    logging.debug("Decimating spaces in {0}".format(filename))

    regeoddspc = re.compile(r',[ ]+')
    rege2spc = re.compile(r' [ ]+')

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

        filename = newfilename(filename, '.txt', '+nospc')
        ''' Make new filename, don't overwrite the original file.
            The batch procedure splits filenames on '_' so we use '+' instead.
        '''

        """ Write cleaned up lines to new filename. """
        with open(filename, 'w+') as fout:
            for row, line in enumerate(fixedlns):
                fout.write(line)

        return(filename)

    except OSError as e:
        logging.error("OS error in decimate_spaces(): %s writing clean data %s", e.strerror, e.filename)
        print('OS error',  e.strerror, ' in decimate_spaces() for filename', e.filename)
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s in decimate_commas():\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception', e.__doc__, ' decimate_spaces():\n', lines[0], '\n', lines[1], '\n', lines[-1])


def decimate_commas(filename):
    """ Read a malformed csv file, which contains empty fields. Decimate commas. Write a clean file.
        Return the clean filename so that this function can be used as a parameter in 
        lines_from_csv(csvfile).
    """
    logging.debug("Decimating commas in {0}".format(filename))

    fixedlns = []

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

        filename = newfilename(filename, '.csv', '+reduced')
        ''' Make new filename, don't overwrite the original file.
            The batch procedure splits filenames on '_' so we use '+' instead.
        '''

        with open(filename, 'w+') as fout:
            """ Write cleaned up lines to new filename. """
            for row, line in enumerate(fixedlns):
                fout.write(line)

        return(filename)

    except OSError as e:
        logging.error("OS error in decimate_commas(): %s writing clean data to filename %s", e.strerror, e.filename)
        print('OS error',  e.strerror, 'decimate_commas() for filename', e.filename)
        
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s in decimate_commas():\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception', e.__doc__, ' in decimate_commas():\n', lines[0], '\n', lines[1], '\n', lines[-1])


def lines_from_csv(csvfile):
    """ Read a well-formed .csv file, or one which contains intentional empty fields. 
        Return a dictionary with row as key and list of lines as elements.
    """
    logging.debug("Extracting lines from report file {0}".format(csvfile))
    
    data = {}
    try:
        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
            lines = list(f)

            for row, line in enumerate(lines):
                #line = regesp.sub('', line)
                line = regecr.sub('', line)
                rlist = line.split(',')
                
                data.update({row: rlist})
                
        return data
        
    except OSError as e:
        logging.error("OS error in lines_from_csv(): %s for filename %s", e.strerror, e.filename)
        print('OS error',  e.strerror, 'lines_from_csv() for filename', e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s in lines_from_csv():\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__, ' in lines_from_csv():\n', lines[0], '\n', lines[1], '\n', lines[-1])
   

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

    cell_2plnum = wb.add_format({'num_format': '0.00'})
    cell_2plnum.set_align('vcenter')

    cell_datetime = wb.add_format({'num_format': dtdict['GMAT1'][1]})
    cell_datetime.set_align('vcenter')
    sheet = wb.add_worksheet('Report')
    """ The presence of the GMAT output report in a tab named 'Report' is a
        Mandatory Interface agreement. 
    """
    try:
        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
            reader = csv.reader(f, quoting=csv.QUOTE_NONE)

            lengs = list()
            for row, line in enumerate(reader):
                for col, data in enumerate(line):
                    if row == 0:
                        data, leng = heading_row(data) # New factorization, fixes squeezed column issue.

                        lengs.append(leng)
                        """ Row 0, lengs list is empty. """

                        sheet.set_column(col, col, leng)
                        """ Column width of row 0 is set to fit longest word in wrapped heading. """
                        sheet.write(row, col, data, cell_wrap)
                    else: # all subsequent rows
                        leng = len(data) + 1

                        if len(lengs) < col + 1:
                            """ There is no element of lengs corresponding to the (zero based) column. """
                            lengs.append(leng)
                        elif leng > lengs[col]:
                            """ Only update the column width if current data is longer than previous. """
                            lengs[col] = leng

                            sheet.set_column(col, col, leng)

                        if regetime.search(data):
                            """ Detect and convert date-time string to date-time value. """
                            gmat_date = dt.datetime.strptime(data, dtdict['GMAT1'][3])

                            sheet.write(row, col, gmat_date, cell_datetime)
                        elif regedecimal.search(data):
                            """ Detect a decimal number.  Note that string.isdecimal() does not work. """

                            if regedot.search(data):
                                modrem = data.split('.')
                                fracsz = len(modrem[1])

                            if fracsz < 4:
                                
                                sheet.write(row, col, data, cell_2plnum)
                            else:
                                sheet.write(row, col, data, cell_4plnum)
                            """ Prevent columns from being too narrow when number is reformatted"""

                        else:
                            sheet.write(row, col, data)
        return str(xlfile)

    except OSError as e:
        logging.error("OS error in csv_to_xlsx(): %s for filename %s", e.strerror, e.filename)
        print('OS error',  e.strerror, 'in csv_to_xlsx() for filename', e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s in csv_to_xlsx():\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__, ' in csv_to_xlsx():\n', lines[0], '\n', lines[1], '\n', lines[-1])
    
    finally:
        wb.close()

if __name__ == "__main__":
    """ Unit Tests for module. """
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
    

    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open REPORT File. NOT BATCH!', 
                    o_path,
                    filter='text files(*.txt *.csv)')

    logging.info('Input report file is %s', fname[0])

    try:
        """ Test Case 1: Run through the toolchain to create an Excel File. """
        nospc = decimate_spaces(fname[0])
        reduced = decimate_commas(nospc)
        xlfile = csv_to_xlsx(reduced)

        nospc = Path(nospc)
        if nospc.exists():
            nospc.unlink()

        reduced = Path(reduced)
        if reduced.exists():
            reduced.unlink()
    
        logging.info('Cleaned up file: %s.', str(xlfile))

    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)
        print('OS error',  e.strerror, 'for filename', e.filename)
    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__, ':\n', lines[0], '\n', lines[1], '\n', lines[-1])
 
    logging.info('Test Case 1: cleaned Excel file is %s', xlfile)
    print('Test Case 1: cleaned Excel file is %s', xlfile)

    try:
        """ Test Case 2: Read the same csv file and return a data dictionary instead."""
        nospc = decimate_spaces(fname[0])
        reduced = decimate_commas(nospc)
        data = lines_from_csv(reduced)
        
        nospc = Path(nospc)

        if nospc.exists():
            nospc.unlink()

        reduced = Path(reduced)
        if reduced.exists():
            reduced.unlink()

        logging.info("Test Case 2: First Row of data: \n\t{0}".format(data[0]))
        print("Test Case 2: First Row of data: \n{0}".format(data[0]))

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Test Case failed with exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
    
    finally:
        qApp.quit()
    