#! Python
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 22:36:55 2019

@author: colin helms

@Description:  Sequentially opens .csv formatted GMAT ReportFiles, 
copies pertinent cells and rows into an .xlsx formatted summary file. 
Also incorporates a difference formula in last column, which is fuel residual (or shortage).

The .csv source files are formatted and saved as .xlsx files, which support subsequent
engineering use.

The file name is split upon the '_' separator and each element is written
to the Summary file as 'metadata'.  This permits arbitrary description or parameterization
to be encoded in the file name and carried forward to the summary file.

@Copyright: Copyright (C) 2019 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@Change Log:
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
import xlwings as xwng
import xlsxwriter as xwrt
import xlsxwriter.utility as xlut
#from pathlib import Path
from gmatlocator import CGMATParticulars
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

def lines_from_csv(csvfile):
    """ Read a .csv formatted file, return a dictionary with row as key and list of lines as elements.
    """
    logging.debug("Extracting lines from report file {0}".format(csvfile))
    
    try:
        regecr = re.compile('\n')
        regesp = re.compile(' ')
    
        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
            lines = list(f)
            data = dict()
            
            for r, row in enumerate(lines):
                r = regesp.sub('', r)
                r = regecr.sub('', r)
                rlist = r.split(',')
                data = {row, rlist}
            return data
        
    except OSError as e:
        logging.error("OS error in csv_to_xlsx(): %s for filename %s", e.strerror, e.filename)
        return None

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exceptionin csv_to_xlsx(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        return None
    

def csv_to_xlsx(csvfile):
    """ Read a .csv formatted file, write it to .xlsx formatted file of the same basename. Return the written
    filename.
    
    Reference Stack Overflow: 
    https://stackoverflow.com/questions/17684610/python-convert-csv-to-xlsx
    with important comments from:
    https://stackoverflow.com/users/235415/ethan
    https://stackoverflow.com/users/596841/pookie
    
    Beware when data has embedded comma.  
    TODO: modify the function to make the csv delimiter variable.  The csv module provides for this.
    """
    logging.debug("Converting report file {0}".format(csvfile))
    
    xlfile = csvfile[:-4] + '.xlsx'
    
    wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
    """ Slice the .csv suffix, append .xlsx suffix, open a new workbook under this name. 
    It seems inefficient to create a .xlsx copy of the .csv file, but the Excel copy is used for
    analysis of data items not included in the summary, e.g. thrust and beta angle history.
    """
    
    sheet = wb.add_worksheet('Report')
    
    sheet.set_column('A:A', 14)
    sheet.set_column('B:B', 6)
    sheet.set_column('C:C', 14)
    sheet.set_column('D:D', 14)
    sheet.set_column('E:E', 14)
    sheet.set_column('F:F', 24)
    sheet.set_column('G:G', 24)
    sheet.set_column('H:H', 24)
    sheet.set_column('I:I', 24)
    sheet.set_column('J:J', 24)
    sheet.set_column('K:K', 24)
    sheet.set_column('L:L', 6)
    sheet.set_column('M:M', 14)
    sheet.set_column('N:N', 14)
    sheet.set_column('O:O', 14)
    sheet.set_column('P:P', 14)
        
    #sheet.set_selection('C2')
    
    sheet.split_panes('C2')

    try:
    
        with open(csvfile, 'rt', newline='', encoding='utf8') as f:
            reader = csv.reader(f, quoting=csv.QUOTE_NONE)
            
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    sheet.write(r, c, col)   

        return xlfile

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
    """ Retrieve the formatting batch file, open and format each .csv file listed """
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
    logging.basicConfig(
            filename='./reduce_report.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! Reduce Report Execution Started !!!!!!!!!!")
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open report batch file.', 
                       os.getenv('USERPROFILE'),
                       filter='Batch files(*.batch)')

    logging.info('Report batch file is %s', fname[0])
    
    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    """ Get the GMAT script path for the summary file.  Put it with the script. 
    The Reports directory is assumed to be in this base path also.
    """
    sfname = 'ReportFile_Summary_' + time.strftime('J%j_%H%M%S',time.gmtime()) + '.xlsx'
    
    summaryfile = os.path.join(o_path, sfname)
    logging.info("Output summary file %s", summaryfile)

    try:
        excel = xwng.App()
        excel.visible=False
    
        xout = xwrt.Workbook(summaryfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
        """ Write summary to this file using XlWriter """
        
        cell_heading = xout.add_format({'bold': True})
        cell_heading.set_align('center')
        cell_heading.set_align('vcenter')
         
        format_2decplace = xout.add_format()
        format_2decplace.set_num_format('0.00')
        
        sumsheet = xout.add_worksheet('Data')
        sumsheet.set_row(0, 15, cell_heading)
                
        sumsheet.activate()
        
        """ Above set outrow 0 with headings """
        
        with open(fname[0]) as f:
            """ Open the master batch file selected in QtFileDialog. 
            This file should contain a line by line list of file paths to
            .csv report files generated by GMAT.
            
            It is necessary to convert .csv files to .xlsx files, with care to ensure
            strings read are written as numbers.  Use the csv module for this. 
            """
            pos = f.tell()
            trialpath = f.readline()
            f.seek(pos)            
            """ Need the first filename in order to determine number of columns for metadata. """
            
            trialname = os.path.basename(trialpath)
            """Strip the path prefix. Do this again below for each file. """

            trialdata = (os.path.splitext(trialname)[0]).split('_')
            """ Get rid of the extension and split the basename into a list. 
            Number of elements in list gives the number of extra columns needed.
            """
            skipcols = len(trialdata)
            """ Number of elements """
                            
            sumsheet.write('A1', 'Report File Name')
            sumsheet.set_column(0, 0, len(trialname) + 4)
            
            col = 0
            """ Offset 1 column """
            for data in trialdata:
                col += 1
                sumsheet.write(0, col, 'Metadata')
                sumsheet.set_column(col, col, len(data) + 4)
                
            sumsheet.set_column(1 + skipcols, 1 + skipcols, 14, format_2decplace)
            sumsheet.write(0, 1 + skipcols, 'Elapsed Days', cell_heading)
            
            sumsheet.set_column(2 + skipcols, 2 + skipcols, 6)
            sumsheet.write(0, 2 + skipcols, 'Revs' )    
            
            sumsheet.set_column(3 + skipcols, 3 + skipcols, 14, format_2decplace)
            sumsheet.write(0, 3 + skipcols, 'Rem. Fuel (kg)')

            sumsheet.set_column(4 + skipcols, 4 + skipcols, 14)
            sumsheet.write(0, 4 + skipcols, 'SMA (km)')

            sumsheet.set_column(5 + skipcols, 5 + skipcols, 10, format_2decplace)
            sumsheet.write(0, 5 + skipcols, 'INC (deg)', cell_heading)

            sumsheet.set_column(6 + skipcols, 6 + skipcols, 10, format_2decplace)
            sumsheet.write(0, 6 + skipcols, 'ECC', cell_heading)

            sumsheet.set_column(7 + skipcols, 7 + skipcols, 14)
            sumsheet.write(0, 7 + skipcols, 'Initial Fuel (kg)')

            sumsheet.set_column(8 + skipcols, 8 + skipcols, 14)
            sumsheet.write(0, 8 + skipcols, 'Fuel Used (kg)')
            
            nrows = len(list(f))
            f.seek(pos)
            
            progress = QProgressDialog("Summary Report", "Cancel", 0, nrows)
            progress.show()
            progress.setValue(0)
            qApp.processEvents()
            
            outrow = 0            
            for filepath in f:                
                """ Iterate through report files named in batch, outrow will start with the 
                first row under the heading and count to the final row in sumsheet.
                
                Convert the filepath from .csv to .xlsx format.
                
                Use the xlwings module to read the xlsx source file.
                
                Use XlWriter to write output in .xlsx format.  XLWriter offers significant
                control over the Excel column and row formats.
                """
                
                if progress.wasCanceled():
                    break
                
                outrow += 1
                progress.setValue(outrow)
                
                rpt = os.path.normpath(filepath)
                rege = re.compile('\n')
                filepath = rege.sub('', rpt)
                """ Cleanup filepath """

                xlsxpath = csv_to_xlsx(str(filepath))

                try:
                    wingbk = excel.books.open(xlsxpath)
                    """ Open xlxs workbook for reading using XlWings.
                    
                    There is a dependency on the source file, as implemented in accordance with 
                    the Include_StaticDefinitions.script model template file.
                                    
                    Source report file Col A, 'ElapsedDays', end range: write to column B of summary.
                    Source report file Col B, 'REV', end range: write to column C of summary.
                    Source report file Col C, 'FuelMass, end range': write to column D of summary.
                    Source report file Col D, 'SMA', end range: write to column E of summary.
                    Source report file Col E, 'INC', end range: write to column F of summary.
                    Source report file Col F, 'ECC', end range: write to column G of summary.
                    Source report file Cell C2, 'FuelMass': write to column H of summary.
                    (Column I will be written with a formula.)
                    """

                    logging.debug("Opened Excel format report file {0}".format(xlsxpath))
                    
                except OSError as e:
                    logging.error("OS error {0}, unable to open file {1}".format(e.strerror, e.filename))
                    raise e
            
                except Exception as e:
                    lines = traceback.format_exc().splitlines()
                    logging.error("Exception opening workbook: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
                    raise e

                insheet = wingbk.sheets['Report']
                """ Name of sheet should be same as set in csv_to_xlsx(). """
                                                                                   
                days = insheet.range('A1').end('down').value
                revs = insheet.range('B1').end('down').value
                remfuel = insheet.range('C1').end('down').value
                sma = insheet.range('D1').end('down').value
                inc = insheet.range('E1').end('down').value
                ecc = insheet.range('F1').end('down').value
                
                inifuel = insheet.range('C2').value
                               
                datalist = [days, revs, remfuel, sma, inc, ecc, inifuel]
                
                filename = os.path.basename(filepath)
                """Strip the path prefix. """

                basename = os.path.splitext(filename)[0]
                metadata = basename.split('_')
                """ Split the basename into a list of meta data. 
                Write the metadata to the metasheet.
                """
                
                sumsheet.write(outrow, 0, filename)
                
                outcol = 1               
                for data in metadata:                    
                    sumsheet.write(outrow, outcol, data)
                    outcol += 1
                
                for data in datalist:
                    sumsheet.write(outrow, outcol, data)
                    outcol += 1
                
                srcname = os.path.basename(xlsxpath)
                excel.books[srcname].close()
                
                fueldiff = \
                '=' + xlut.xl_rowcol_to_cell(outrow, 7 + skipcols) + \
                '-' + xlut.xl_rowcol_to_cell(outrow, 3 + skipcols)
                                
                sumsheet.write_formula(outrow, outcol, fueldiff)
                
                logging.info("Completed extract from Excel report file {0}".format(filename))
                
            progress.setValue(nrows)

    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
    
    finally:
        xout.close()
        excel.quit()
    