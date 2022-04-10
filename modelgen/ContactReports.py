#! Python
# -*- coding: utf-8 -*-
"""
@description:  module container for class definition ContactReports.

@author: colinhelms@outlook.com

@copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@changes:
    Created on Fri Mar 8 2019
"""
import re
import platform
import logging
import traceback
import getpass
import csv
import xlsxwriter as xwrt
import datetime as dt
import reduce_report as rr
from pathlib import Path
from LinkReports import CLinkReports
from CleanUpReports import CCleanUpReports
from gmatlocator import CGMATParticulars
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

class ContactReports(CCleanUpReports):
    """ Specialization class to format a SIGHT Locator file. Is extended to decimate 
        Link Report files by time span of contacts in SIGHT Locator. Once saved the 
        new file is a Contact Report suitable for image collection and link budget
        calculations. 
    """

    def __init__(self, **args):
        self.aoi = {'':{}}
        """ This is used by the instance to capture the Observer satellite and Target location."""

        self.reports = CLinkReports()
        """ Associated class def is used to process Link Reports. """

        super().__init__(**args)


    def do_batch(self, contbatch, linkbatch): # This argument signature is unsatisfactory, should be dictionary form
        """ Call the parent instance of extend() for the linkbatch files, call the self instance of extend()
            for the contbatch files. 
        """

        super().do_batch(contbatch, None)
        """ Delegate up the MRO chain.
            See: https://stackoverflow.com/questions/32014260
            See: https://rhettinger.wordpress.com/2011/05/26/super-considered-super
        """
        
        self.reports.do_batch(linkbatch)

        regecr = re.compile('\s')
        try:    
            with open(contbatch) as f:
                for filepath in f:                
                    """ Iterate through report files named in batch. """
                    logging.info('Reducing file: {0}'.format(filepath))

                    filepath = regecr.sub('',filepath)
                    """Get rid of newline following path string."""

                    rpt = Path(filepath)

                    if rpt.exists() & rpt.is_file():
                        
                        self.extend(rpt)
                        """ Make sure that the derived instance of extend() is called. """

                        self.filelist.append(rpt.name)

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])    

    def extend(self, rpt):
        """ Call parent class do_batch(contfiles) and the ContactReports.extend function will be called.
            This specialization of extend() reads SIGHT Reports and writes out a formatted 
            Excel copy.

            Parameters:
                rpt - a single SIGHTLocator_[Sat#].txt file

            During the batch processing a two level dictionary is also compiled as a 
            map for creation of Contact Reports by the merge function. 
        """
        regetarget = re.compile('Target: ')
        regeobsrvr = re.compile('Observer: ')
        regeutchead = re.compile('UTC')
        regedurhead = re.compile('Duration')
        regecr = re.compile('\s')
        regetime = re.compile(rr.dtdict['GMAT1'][2])
        """ Regular Expression Match patterns to identify files data items. """

        fname = (rpt.stem).split('+')[0]
        """Get rid of the 'nospc' and 'reduced' keywords."""

        xlfile = rr.newfilename(rpt.parents[0]/fname, '.xlsx')
        """Slice the .csv suffix, append .xlsx suffix, open a new workbook under this name."""

        try:
            wb = xwrt.Workbook(xlfile, 
                {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
            
        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in extend(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
            return None

        cell_heading = wb.add_format({'bold': True})
        cell_heading.set_align('center')
        cell_heading.set_align('vcenter')
        cell_heading.set_text_wrap()

        cell_wrap = wb.add_format({'text_wrap': True})
        cell_wrap.set_align('vcenter')

        cell_3plnum = wb.add_format({'num_format': '0.000'})
        cell_3plnum.set_align('vcenter')

        cell_datetime = wb.add_format({'num_format': rr.dtdict['GMAT1'][1]})
        cell_datetime.set_align('vcenter')

        nospc = rr.decimate_space(rpt)
        nospc = Path(nospc)

        reduced = rr.decimate_commas(nospc)
        reduced = Path(reduced)
        
        if nospc.exists():
            nospc.unlink()
            """ Remove temporary file. """

        try:
            with open(reduced, 'rt', newline='', encoding='utf8') as f:
                reader = csv.reader(f, quoting=csv.QUOTE_NONE)

                lengs = list()
                for row, line in enumerate(reader):
                    for col, data in enumerate(line):

                        if regecr.match(data):
                            continue
                       
                        match = regetarget.match(data)
                        if match:
                            key1 = data[match.span()[1]:len(data)]
                            continue

                        match = regeobsrvr.match(data)
                        if match:
                            loc = data[match.span()[1]:len(data)]
                            key2 = loc.split('_')[1]
                            self.aoi.update({str(key1 +'@' + key2):xlfile})
                            """ Keep track of files written, by compound key of satellite and AOI. """

                            sheet = wb.add_worksheet(key2)
                            """ Start a new sheet for each Observer. """
                            lengs.clear()
                            continue

                        if regeutchead.search(data):
                            sheet.write(row, col, data, cell_heading)

                        elif regedurhead.search(data):
                            sheet.write(row, col, data, cell_heading)
                            col_duration = col
                        
                        elif regetime.match(data):
                            gmat_date = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                            sheet.write(row, col, gmat_date, cell_datetime)
                                
                        elif col == col_duration:
                            sheet.write(row, col, data, cell_3plnum)

                        else:
                            sheet.write(row, col, data)

                        leng = len(data) + 1

                        if len(lengs) < col+1:
                            lengs.append(leng)
                        else:
                            lengs[col] = leng
 
                        sheet.set_column(col, col, leng)
                        """ Set the width of each column for widest data. """

            if reduced.exists():
                reduced.unlink()
                """ Remove temporary file """

        except OSError as e:
            logging.error("OS error in csv_to_xlsx(): %s for filename %s", e.strerror, e.filename)
            return None

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exceptionin csv_to_xlsx(): %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
            return None
                
        finally:
            wb.close()
        
        
    def merge(self):
        """ This function is an extension that builds a contact report by combining Link Reports
            with SIGHT reports. It operates like the LinkReports class, but includes only time spans
            which are listed in the SIGHT Report.
        
            The instance aoi dictionary provides a means to open the related SIGHT Report by keys
            AOI and Satellite.

            ContactReports is also an instance of LinkReports.  If the parent extend() is called
            before merge() in ContactReports, the links dictionary will be updated and provide  
            the means to open the Link Report corresponding to the same satellite and AOI keys.

            The workflow is (1) create an instance of ContactReports, (2) start do_batch()

            The SIGHT Report contact time spans are used to decimate each such related
            Link Report. The saved ContactReport is the simply the decimated Link Report for 
            each Satellite and AOI in the dictionary.

            @TODO include local times of sunrise and sunset.
        """
        regecr = re.compile('\s')

        try:
            """ self.aoi = {AOI:{satellite:ContactReports}}
                super().links = {AOI:{satellite:LinkReports}}

                For each ContactReport use time spans to decimate the LinkReport for
                the same AOI and satellite
            """
            for loc, cfiles in self.aoi:
                


            with open(contfiles) as f:
                for filepath in f:                
                    """ Iterate through report files named in batch. """
                    logging.info('Merging Contacts in: {0}'.format(filepath))

                    filepath = regecr.sub('', filepath)
                    """Get rid of newline following path string."""

                    rpt = Path(filepath)

                    if rpt.exists() & rpt.is_file():
                        self.filelist.append(rpt.name)

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
    
    cleanup = ContactReports

    gmat_paths = CGMATParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open SIGHT Report batch file.', 
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
        logging.info ('SIGHT Report batch completed.')
        print('SIGHT Report batch completed.')

    fname = QFileDialog().getOpenFileName(None, 'Open Link report batch file.', 
                       o_path(),
                       filter='Batch files(*.batch)')

    logging.info('Link Report batch file is %s', fname[0])
    
    try:
        cleanup.merge(fname[0])
        """ Derived class combination of Link Report with SIGHT Report. """

    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
        
    finally:
        logging.info ('Merge Batch completed.')
        print('Merge Batch completed.')
        qApp.quit()
        
    