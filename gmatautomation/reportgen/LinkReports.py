#! Python
# -*- coding: utf-8 -*-
"""
@file: LinkReports.py

@brief: This module provides a container for class definition CLinkReports.

@copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

@author  Colin Helms, colinhelms@outlook.com, [CCH]

@details ContactReports provides Excel workbook formatting of Link Reports.
These files are GMAT ReportFiles that provide satellite geodetic Latitude,
Longitude, Altitude and topographic X/Y/Z coordinates relative to
a fixed ground station.  They fundamental to Link Budget and Earth Observation
Collection Planning calculations.
   
@remark Change History
    Created on Fri Mar 8 2019, GitHub repository GMAT-Automation.
    Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.

@bug https://github.com/a093130/GMAT-Automation/issues
"""
import re
import logging
import traceback
import csv
import pywintypes as pwin
import xlsxwriter as xwrt
from pathlib import Path
import datetime as dt
from gmatautomation import reduce_report as rr
from gmatautomation import CCleanUpReports
from gmatautomation import CGmatParticulars

class CLinkReports(CCleanUpReports):
    """ Specialization class to format a Link Report file. """
    def __init__(self, **args):
        super().__init__(**args)

        self.links = {}
        """ This is used by the instance to capture the Observer satellite and Target location."""
        return

    def extend(self, rpt):
        """ This specialization of extend() provides specialized methods to format 
            a Link Report. 
        """
        regegregor = re.compile('A1Gregorian')
        regesatnum = re.compile('LEOsat')
        regeaoi = re.compile(' X')
        regespch = re.compile(' [A-Z]+')
        """ Regular Expression Match patterns to identify files data items. """

        nospc = rr.decimate_spaces(rpt)
        nospc = Path(nospc)

        reduced = rr.decimate_commas(nospc)
        reduced = Path(reduced)
        """ Defer removal of reduced.csv to follow file read below. """

        if nospc.exists():
            nospc.unlink()

        fname = (rpt.stem).split('+')[0]
        """Get rid of the 'nospc' and 'reduced' keywords."""

        xlfile = rr.newfilename(rpt.parents[0]/fname, '.xlsx')
        """Slice the .csv suffix, append .xlsx suffix, open a new workbook under this name."""
        
        try:
            wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})

            sheet = wb.add_worksheet('Report')
            """ The presence of the GMAT output report in a tab named 'Report' is a
                Mandatory Interface agreement.
            """
        except OSError as e:
            lines = traceback.format_exc().splitlines()
            logging.error("OS error: %s in CLinkReports extend() for filename %s.\n%s\n%s\n%s", e.strerror, e.filename,\
                lines[0], lines[1], lines[-1])
            print('OS error: ', e.strerror,' in CLinkReports extend() for filename ', e.filename,\
                '\n', lines[0], '\n', lines[1], '\n', lines[-1])

            return # Let do_batch() try another file.      

        except pwin.com_error as ouch:
            lines = traceback.format_exc().splitlines()
            logging.error('Excel Workbook raised Windows com error in CLinkReports. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))
            print('Excel Workbook raised Windows com error in CLinkReports. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))
            
            return # Let do_batch() try another file.

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in CLinkReports extend(): %s\n%s\n%s\n%s", e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in CLinkReports extend(): ', e.__doc__, '\n', lines[0], '\n', lines[1],'\n', lines[-1])
        
            return # Let do_batch() try another file.

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

        cell_datetime = wb.add_format({'num_format': rr.dtdict['GMAT1'][1]})
        cell_datetime.set_align('vcenter')
    
        try:
            with open(reduced, 'rt', newline='', encoding='utf8') as f:
                reader = csv.reader(f, quoting=csv.QUOTE_NONE)
                
                lengs = []
                for row, line in enumerate(reader):
                    for col, data in enumerate(line):
                        if rr.regecr.match(data):
                            continue   

                        if len(data) > 0:
                            if row == 0:
                                """ Format the heading row.  Row 0 is always the heading row for LinkReports. """
                                data, leng = rr.heading_row(data)
                                
                                lengs.append(leng)
                                sheet.set_column(col, col, leng)

                                if col == 0:
                                    if regegregor.search(data):
                                        """ Find and record the satellite number. """
                                        match = regesatnum.match(data)
                                        if match:
                                            key1 = data[match.span()[0]:match.span()[1] + 1]
                                            key1 = rr.regesp.sub('',key1)
                                    else:
                                        """ Column 0 must match 'A1Gregorian' - incompatible report format. """
                                        raise  ValueError('{0}, row 0 col 0 is incompatible with expected heading.'.format(reduced.name()))

                                ematch = regeaoi.search(data)
                                """ The span of this match object will identify the end of the AOI string. """
                                bmatch = regespch.search(data)
                                """ The span of this match object will identify the beg + 1 of the AOI string. """

                                if ematch:
                                    key2 = data[(bmatch.span()[1] - 1):ematch.span()[0]]
                                    key2 = rr.regesp.sub('', key2)
                                    
                                    self.links.update({str(key1 +'@'+ key2):xlfile})
                                    """ Keep track of files written, by compound key of satellite and AOI. """
                                    
                                sheet.write(row, col, data, cell_heading)
                                
                            else: # all subsequent rows
                                """ Format report data rows. """
                                leng = len(data) + 1

                                if len(lengs) < col + 1:
                                    """ There is no element of lengs corresponding to the (zero based) column. """
                                    lengs.append(leng)
                                elif leng > lengs[col]:
                                   """ Only update the column width if current data is longer than previous. """
                                   lengs[col] = leng
                      
                                sheet.set_column(col, col, leng)
                            
                                if rr.regetime.search(data):
                                    """ Detect and convert date-time string to date-time value. """
                                    gmat_date = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                                    
                                    sheet.write(row, col, gmat_date, cell_datetime)
                                elif rr.regedecimal.search(data):
                                    """ Detect a decimal number.  Note that string.isdecimal() does not work. """

                                    if rr.regedot.search(data):
                                        modrem = data.split('.')
                                        fracsz = len(modrem[1])

                                    if fracsz < 4:
                                        
                                        sheet.write(row, col, data, cell_2plnum)
                                    else:
                                        sheet.write(row, col, data, cell_4plnum)
                                    """ Prevent columns from being too narrow when number is reformatted"""
                                else:
                                    sheet.write(row, col, data)

                sheet.freeze_panes('A2')
                """ Lock the first row, first column after formatting of all rows and columns is done. """

                logging.info('LinkReports extend() completed  for filename %s', xlfile.name)
                print ('LinkReports extend() completed  for filename:', xlfile.name)

            if reduced.exists():
                reduced.unlink()
            
            return

        except OSError as e:
            logging.error('%s in LinkReports extend(): %s in filename %s\n%s\n%s\n%s',\
                e.__doc__, e.strerror, e.filename, lines[0], lines[1], lines[-1])
            print(e.__doc__,' in LinkReports extend(): ', e.strerror,\
                '\nin filename ',e.filename,'\n',lines[0],'\n', lines[1],'\n', lines[-1])

        except ValueError as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Value Error. A1Gregorian heading not found. %s\n%s\n%s\n%s',\
                e.args[0], lines[0], lines[1], lines[-1])
            print('ValueError in Link Reports extend()', '\n',e.args[0],'\n', lines[0],'\n',lines[1],'\n',lines[1],'\n', lines[-1])

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Exception in LinkReports extend(): %s\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in LinkReports extend():' '\n', e.__doc__, '\n', lines[0],'\n', lines[1], '\n', lines[-1])
             
        finally:
            wb.close()
        

if __name__ == "__main__":
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code.
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
    import platform
    import getpass
    from PyQt5.QtWidgets import(QApplication, QFileDialog)

    logging.basicConfig(
            filename='./LinkReports.log',
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

    gmat_paths = CGmatParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open LinkReport batch file', 
                    str(o_path),
                    filter='text files(*.batch)')

    logging.info('Input batch file is %s', fname[0])

    reports = CLinkReports()
    batchfile = Path(fname[0])

    try:
        """ method do_batch() delegated to CCleanUpReports.
            method extend() from LinkReports specialization.
            Use LinkReport.batch as input.
        """
        reports.do_batch(batchfile)

        logging.info('Dictionary Entries follow:')
        print('Dictionary Entries follow:')
        
        for key, value in reports.links.items():
            print(key, value)

        logging.info('LinkReports complete.')
        print('LinkReports complete.')

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("LinkReports failed with exception: %s. %s\n%s", e.__doc__, lines[0], lines[-1])
        print('LinkReports failed with exception: ', e.__doc__, '\n', lines[0], '\n', lines[-1])
    
    finally:
        qApp.quit()
    