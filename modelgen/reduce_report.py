#! Python
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  8 22:36:55 2019

@author: colin helms

@Description:
    Sequentially opens .csv formatted GMAT ReportFiles, copies the top and bottom rows into a .xlsx file,
    finds difference in column C, first and last rows, which is Fuel residual (or shortage). 
    
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
from pathlib import Path
from gmatlocator import CGMATParticulars
from PyQt5.QtWidgets import(QApplication, QFileDialog)

def csv_to_xlsx(self, csvfile):
    """ Read a .csv formatted file, write it to .xlsx formatted file of the same basename.
    Reference Stack Overflow: 
    https://stackoverflow.com/questions/17684610/python-convert-csv-to-xlsx
    with important comments from:
    https://stackoverflow.com/users/235415/ethan
    https://stackoverflow.com/users/596841/pookie
    
    Beware when data has embedded comma.  
    TODO: modify the function to make the csv delimiter variable.  The csv module provides for this.
    """
    xlfile = csvfile[:-4] + '.xlsx'
    wb = xwrt.workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
    """ Slice off the .csv suffix, append .xlsx suffix, open a new workbook under this name. """
    
    sheet = wb.add_worksheet('Report')
    
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
        wb.close

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
            level=logging.DEBUG,
            format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! Reduce Report Execution Started !!!!!!!!!!")
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
    app = QApplication([])
    
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
        xout = xwrt.Workbook(summaryfile)
        """ Write summary to this file using XlWriter """
        
        cell_heading = xout.add_format({'bold': True})
        cell_heading.set_align('center')
        cell_heading.set_align('vcenter')
        
        sumsheet = xout.add_worksheet('Data')
        sumsheet.set_row(0, 15, cell_heading)
        
        metasheet = xout.add_worksheet('Metadata')
        metasheet.set_row(0, 15, cell_heading)
        
        sumsheet.activate()
        
        sumsheet.set_column('A:A', 40)
        sumsheet.set_column('B:B', 14)
        sumsheet.set_column('C:C', 6)
        sumsheet.set_column('D:D', 24)
        sumsheet.set_column('E:E', 14)
        sumsheet.set_column('F:F', 14 )
        sumsheet.set_column('G:G', 14)
        sumsheet.set_column('H:H', 24)
        sumsheet.set_column('I:I', 24)
            
        sumsheet.write('A1', 'Report File Name')
        sumsheet.write('B1', 'Elapsed Days')
        sumsheet.write('C1', 'Revs')    
        sumsheet.write('D1', 'Remaining Fuel (kg)')
        sumsheet.write('E1', 'SMA (km)')
        sumsheet.write('F1', 'INC (deg)')
        sumsheet.write('G1', 'ECC')
        sumsheet.write('H1', 'Initial Fuel (kg)')
        sumsheet.write('I1', 'Fuel Used (kg)')    
        
        sumsheet.split_panes(1, 1)
        
        outrow = 0
        """ Above set outrow 0 with headings """
        
        with open(fname[0]) as f:
            """ Open the master batch file selected in QtFileDialog. 
            This file should contain a line by line list of file paths to
            .csv report files generated by GMAT.
            
            It is necessary to convert .csv files to .xlsx files, with care to ensure
            strings read are written as numbers.  Use the csv module for this. 
            """
            for filepath in f:
                """ Iterate through report files named in batch, outrow will start with the 
                first row under the heading and count to the final row in sumsheet.
                
                Use the csv module to convert to Excel format.  Use the xlwings module 
                to read the source file xlsx format.
                
                Use XlWriter to write output in .xlsx format.  XLWriter also offers significant
                control over the Excel column and row formats.
                """
                outrow += 1
                
                rpt = os.path.normpath(filepath)
                rege = re.compile('\n')
                filepath = rege.sub('', rpt)
                """ Cleanup filepath """

                logging.debug("Reading report file %s", filepath)
                                
                wb = xwng.Book(csv_to_xlsx(filepath))
                """ There is a dependency on the source file, as implemented in accordance with 
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
                
                insheet = wb.sheets['Report']
                """ Name of sheet is set in csv_to_xlsx(). """
                
                if insheet.range('B1') != 'REV':
                    """ Cell 'B1' is used as a check should contain string "REV". """
                    raise ValueError('Workbook {0} is invalid, Sheet 1 check value fails.'.format(filepath))
                                
                days = insheet.range('A1').end('down').value()
                revs = insheet.range('B1').end('down').value()
                remfuel = insheet.range('C1').end('down').value()
                sma = insheet.range('D1').end('down').value()
                inc = insheet.range('E1').end('down').value()
                ecc = insheet.range('F1').end('down').value()
                
                inifuel = insheet.range('C2').value()
                
                filename = os.path.basename(filepath)
                """Strip the path prefix. """
                
                sumsheet.write(0, outrow, filename)
                sumsheet.write(1, outrow, days)
                sumsheet.write(2, outrow, revs)
                sumsheet.write(3, outrow, remfuel)
                sumsheet.write(4, outrow, sma)
                sumsheet.write(5, outrow, inc)
                sumsheet.write(6, outrow, ecc)
                sumsheet.write(7, outrow, inifuel)

                metasheet.write(outrow, 0, filename)
            
                metadata = (os.path.splitext(filename)[0]).split('_')
                """ Split the basename into a list of meta data. 
                Write the metadata to the metasheet.
                """
                outcol = 0
                for data in metadata:
                    metasheet.write(outrow, outcol, filename)
                    outcol += 1
                                                
            sumsheet.add_table(1, 7, outrow, 7)
            """ 'Table1': 'Initial Fuel Mass' """
            sumsheet.add_table(1, 3, outrow, 3)
            """ 'Table2': 'Remaining Fuel Mass' """
            sumsheet.write_formula('I1', '=Table1[] - Table2[]')

    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
    
    finally:
        xout.close()
    