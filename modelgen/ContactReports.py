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
import logging
from runpy import run_path
import traceback
import sys
import pywintypes as pwin
import xlwings as xw
import xlsxwriter as xwrt
import datetime as dt
import reduce_report as rr
from pathlib import Path
from LinkReports import CLinkReports
from CleanUpReports import CCleanUpReports
from LinkReports import CLinkReports
from gmatlocator import CGMATParticulars

class CContactReports(CCleanUpReports):
    """ Specialization class to generate Contact Reports.

        Contact Reports of Link Report data rows which are selected between the 
        start time and stop time for each visibility as tabulated in a SIGHT Report.
        Contact reports are generated, one for each defined spacecraft and contain 
        worksheets for each Area of Interest (AOI) or Ground Station.

        Link Reports are GMAT ReportFiles consisting of rows containing
        geodetic latitude, longitude and altitude of a spacecraft nadir point 
        as well as Cartesian coordinates of the spacecraft relative to a Ground Station
        or AOI. There is one Link Report produced for each AOI and spacecraft
        combination.
        
        SIGHT Reports are ContactLocator Reports output by the GMAT Event Locator
        following an ephemeris simulation.  These reports are output, one for each
        defined spacecraft.  They give the start time, stop time and duration for each
        visibility between the spacecraft and AOI or Ground Station.
    """

    def __init__(self, **args):
        super().__init__(**args)

        self.links = dict()
        """ Dictionary initialized to capture the link files for each satellite and AOI. """
        self.aoi = dict()
        """ Dictionary used by the instance to capture the contact files for each satellite and AOI. """
        return
    
    def setlinks(self, lrfiles:dict):
        """ links instance variable must be set from external source. """
        try:
            if isinstance(lrfiles,dict):
                self.links = lrfiles.copy()
                
            else:
                raise ValueError('setlinks() input parameter 2 is not a dict type.')

            return

        except ValueError as e:
            lines = traceback.format_exc().splitlines()
            logging.error('CContactReports setlinks incompatible input value, %s.\n%s\n%s\n%s', e.args[0],lines[0], lines[1], lines[-1])
            print('CContactReports setlinks incompatible input parameter, ', e.args[0],'\n',lines[0],'\n',lines[1],'\n',lines[-1])

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in CContactReports extend(): %s\n%s\n%s\n%s", e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in CContactReports extend(): ', e.__doc__, '\n', lines[0], '\n', lines[1],'\n', lines[-1])

 
    def extend(self, sightfile):
        """ This specialization of extend() provides specialized methods to format 
            a Contact Locator Report (SIGHTLocator Report).
        """
        regetarget = re.compile('Target:[ ]*')
        regeobsrvr = re.compile('Observer:[ ]*')
        regeutc = re.compile('UTC')
        regedur = re.compile('Duration')
        regesat = re.compile('[sS]at')
        regekey2 = re.compile('[A-Z][A-Za-z]+_')
        regern = re.compile(r'[\r\n]')

        try:
            rpt = Path(sightfile)
            xlfile = rpt.with_suffix('.xlsx')

            nospc = rr.decimate_spaces(rpt)
            reduced = rr.decimate_commas(nospc)
            lines = rr.lines_from_csv(reduced)

            nospc = Path(nospc)
            if nospc.exists():
                nospc.unlink()

            reduced = Path(reduced)
            if reduced.exists():
                reduced.unlink()
            
            wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
            print('Creating Output workbook {0}'.format(xlfile.name))

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

            lengs = list()
            times = list()
            colutc = list()
            startstop = dict()
            for row, rlist in lines.items():
                for col, data in enumerate(rlist):
                    """ Make a list of start and stop times from Contact Locator file. """
                    if len(data) > 0:
                        data = regern.sub('', data) # workaround - where do these raw strings come from?

                        mtarg = regetarget.match(data)
                        if mtarg:
                            """ Target row contains the satellite number and occurs once in each file."""
                            key1 = data[(mtarg.span()[1]):len(data)]
                        
                            xlfpstr = str(xlfile.stem)
                            match = regesat.search(xlfpstr)
                            if match:
                                satinfn = xlfpstr[match.span()[0]:len(str(xlfpstr))]
                            else:
                                raise ValueError('No Satellite number in filename {0}.'.format(xlfile.name))
                        
                            keycheck = None
                            """ Normalize key1 and satellite numbers before comparison. """

                            match = regesat.search(key1)
                            if match.span()[1] == len(key1):
                                """ There is no numeric character at end of key1, implies LEOsat, Sat1. """
                                keycheck = key1 + '1'
                            elif match.span()[1] < len(key1):
                                """ one or more characters at end of key1 """
                                if key1[match.span()[1]].isnumeric():
                                    keycheck = key1
                            else:
                                """ key1 does not contain a recognizable satellite number. """
                                raise ValueError('No recognizable satellite number in key1: {0}'.format(key1))

                            regekeymatch = re.compile(satinfn, re.IGNORECASE)
                            if regekeymatch.search(keycheck):
                                pass
                            else:
                                raise ValueError('Satellite number in filename {0} does not match data from file.'.format(xlfile.name))

                        mobs = regeobsrvr.match(data)
                        if mobs:
                            """ Observer row contains the AOI and recurs for each data set in the file. """
                            obs = (data[(mobs.span()[1]):len(data)])

                            match = regekey2.search(obs)
                            if match:
                                key2 = obs[match.span()[1]:len(obs)]
                                sheet = wb.add_worksheet(key2)

                                key = key1 +'@'+ key2
                                self.aoi.update({key:xlfile})
                                """ Keep track of files written, by compound key of satellite and AOI. """
                                
                                continue
                            else:
                                raise ValueError('Observer string %s does not contain key2 as expected.', data)
                        
                        match = regeutc.search(data)
                        """ The 'UTC' heading row identifies the start/stop columns and recurs for each data set in the file. """
                        if match:
                            if len(colutc) == 0:
                                colutc.append(col)
                                continue
                            elif len(colutc) == 1:
                                colutc.append(col)
                                continue
                            else:
                                """ only do this one time for the current file. """
                                pass
                        
                        match = regedur.search(data)
                        if match:
                            continue

                        match = rr.regetime.match(data)
                        if match:
                            """ Time data occurs for each visibility between a Target and Observer."""
                            if col == colutc[0]:
                                """ A  contact Start Time in Excel datetime form. """
                                starttime = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                                continue
                            elif col == colutc[1]:
                                """ A contact Stop Time in Excel datetime form. """
                                stoptime = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                                startstop.update({str(key):[starttime, stoptime]})
                                """ Multiple start/stop times for each key. """
                                continue
                            else:
                                raise ValueError('Unexpected time field found in data column, %d.', col)
                
            print('Start/Stop times determined. Decimating Link Report.')

            for key, times in startstop.items():
                """ Use the ContactLocator start/stop times to select data from the Link Report"""
                try:
                    linkfile = self.links[key]

                except KeyError as e:
                    KeyError('Warning: Key %s Not found in self.links.  Continuing with next key.', key)
                    print('Key {0} not associated with Link Report file in links dictionary. Continuing.'.format(key))
                    continue
                
                if Path(linkfile).exists:
                    """ Get and open the Link ReportFile using xlwings. """
                    print('Opening input Link File {0} in XLWings.'.format(linkfile.name))

                    with xw.App() as excel:
                        excel.visible=False

                        aoi = key.split('@')[1]
                        """ Worksheet keys are shortened to just the AOI. """
                        
                        sheet = wb.get_worksheet_by_name(aoi)
                        """ The sheets to be written are created and named after the keys. """
                        print('Accessing {0} worksheet {1} for writing.'.format(xlfile.name, aoi))
                        
                        bk = excel.books.open(str(linkfile))
                        try:
                            lrsheet = bk.sheets['Report']
                            """ The presence of the GMAT output report in a tab named 'Report'
                                is a mandatory condition. 
                            """
                        except pwin.com_error as ouch:
                            logging.warning('Access to Link Report sheet raised Windows com error. {0}, {1}'\
                                .format(type(ouch), ouch.args[1]))
                            print('Warning: Access to Link Report sheet raised Windows com error. {0}, {1}'\
                                .format(type(ouch), ouch.args[1])) 
                            continue

                        writerow = 0
                        lreprng = lrsheet.range('A1').expand()
                        for row, des in enumerate(lreprng.rows):
                            """ Link Report data row 0 is text, subsequent rows are either datetime or numeric. """
                            data = des.value
                            """ data is a list of the cells in the row.
                                cellval = data[0] is under A1Gregorian
                                cellval = data[2] is under Earth Fixed Planetodetic LAT
                                cellval = data[3] is under Earth Fixed Planetodetic LON
                                cellval = data[4] is under Earth Altitude
                                cellval = data[5] is under [AOI] X
                                cellval = data[6] is under [AOI] Y
                                cellval = data[7] is under [AOI] Z
                            """

                            """ @TODO:Performance enhancement. iterate through data[0] 
                            and identify the row for each start time.
                            """
                  
                            if row == 0:
                                for col, cellval in enumerate(data):
                                    cellval, leng = rr.heading_row(cellval)
                                    lengs.append(leng)
                                    sheet.set_column(col, col, leng)
                                    
                                    sheet.write(row, col, cellval, cell_heading)
                                continue

                            elif row > 0:                        
                                timegtstart = data[0] >= times[0]
                                timeltstop = data[0] <= times[1]
                                if timegtstart & timeltstop:                                  
                                    for col, cellval in enumerate(data):
                                        if col == 0:
                                            writerow += 1
                                            """ Only for col == 0 """

                                            cellstr = cellval.strftime(rr.dtdict['GMAT1'][3])
                                            print('Writing Link Report data for time {0}'.format(cellstr))
                                            leng = len(cellstr)
                                            if len(lengs) < col + 1:
                                                """ There is no element of lengs corresponding to the (zero based) column. """
                                                lengs.append(leng)
                                            elif leng > lengs[col]:
                                                """ Only update the column width if current data is longer than previous. """
                                                lengs[col] = leng

                                            sheet.set_column(col, col, leng)
                                            sheet.write(writerow, col, cellval, cell_datetime)
                                            continue
                                    
                                        elif col > 0:
                                            cellstr = '{: 0.4f}'.format(cellval)
                                            leng = len(cellstr) + 2
                                            if len(lengs) < col + 1:
                                                """ There is no element of lengs corresponding to the (zero based) column. """
                                                lengs.append(leng)

                                            elif leng > lengs[col]:
                                                """ Only update the column width if current data is longer than previous. """
                                                lengs[col] = leng
                                            
                                            sheet.set_column(col, col, leng)
                                            sheet.write(writerow, col, cellval, cell_4plnum)
                                            continue
                                else:
                                    """ Only write data that is between the ContactLocator start and stop times."""
                                    continue

                        sheet.freeze_panes('A2')
                        """ Lock the first row, first column after formatting of all rows and columns is done. """                               

            logging.info('CContactReports extend() completed for filename: %s.', rpt.name)
            print ('CContactReports extend() completed  for filename:', rpt.name)

            return

        except OSError as e:
            logging.error("OS error: %s in CContactReports extend() for filename %s", e.strerror, e.filename)
            print('OS error: ', e.strerror,' in CContactReports extend() for filename ', e.filename)

        except ValueError as e:
            lines = traceback.format_exc().splitlines()
            logging.error('%s, Incompatible input value, %s.\n%s\n%s\n%s', rpt.name, e.args[0],lines[0], lines[1], lines[-1])
            print(rpt.name, ', Incompatible input value, ',e.args[0], '\n', lines[0], '\n', lines[1],'\n', lines[-1])

        except pwin.com_error as ouch:
            lines = traceback.format_exc().splitlines()
            logging.error('Access to excel raised Windows com error. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))
            print('Access to excel raised Windows com error. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in CContactReports extend(): %s\n%s\n%s\n%s", e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in CContactReports extend(): ', e.__doc__, '\n', lines[0], '\n', lines[1],'\n', lines[-1])
        
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
    import getpass
    import platform
    from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

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

    try:
        gmat_paths = CGMATParticulars()
        o_path = gmat_paths.get_output_path()
        """ o_path is an instance of Path that locates the GMAT output directory. """

        qApp = QApplication([])

        fname = QFileDialog().getOpenFileName(None, 'Open Link Reports batch file', 
                        o_path,
                        filter='text files(*.batch)')

        logging.info('Link Report batch file is %s', fname[0])

        batchfile = Path(fname[0])

        lr_inst = CLinkReports()
        lr_inst.do_batch(batchfile)

        logging.info ('Link Reports completed.')
        print('Link Reports completed.')
        fname = QFileDialog().getOpenFileName(None, 'Open (SIGHT) Contact Locater Reports batch file.', 
                        o_path,
                        filter='Batch files(*.batch)')

        logging.info('Contact Locater Report batch file is %s', fname[0])

        sightfile = Path(fname[0])

        cr_inst = CContactReports()
        cr_inst.setlinks(lr_inst.links)        
        cr_inst.do_batch(sightfile)

        logging.info ('Contact Locater Reports completed.')
        print('Contact Locater Reports completed.')
        print('Contact Reports Test Case Completed.')

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s caught at top level:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[-1])
        
    finally:
        qApp.quit()

    