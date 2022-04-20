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
import bisect as bi
import copy as cp
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
        regenumevt = re.compile('Number of events')
        regedur = re.compile('Duration')
        regesat = re.compile('[sS]at')
        regekey2 = re.compile('[A-Z][A-Za-z]+_')
        regern = re.compile(r'[\r\n]')

        rpt = Path(sightfile)
        xlfile = rpt.with_suffix('.xlsx')

        print('Creating Output workbook {0}'.format(xlfile.name))
        try:
            wb = xwrt.Workbook(xlfile, {'constant_memory':True, 'strings_to_numbers':True, 'nan_inf_to_errors': True})
            
        except OSError as e:
            lines = traceback.format_exc().splitlines()
            logging.error("OS error: %s in CContactReports extend() for filename %s.\n%s\n%s\n%s", e.strerror, e.filename,\
                lines[0], lines[1], lines[-1])
            print('OS error: ', e.strerror,' in CContactReports extend() for filename ', e.filename,\
                '\n', lines[0], '\n', lines[1], '\n', lines[-1])
            return # Let do_batch() try another file.  

        except pwin.com_error as ouch:
            lines = traceback.format_exc().splitlines()
            logging.error('Excel Workbook raised Windows com error in CContactReports. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))
            print('Excel Workbook raised Windows com error in CContactReports. {0}, {1}\n{2}\n{3}\n{4}'\
                .format(type(ouch), ouch.args[1], lines[0], lines[1], lines[-1]))
            return # Let do_batch() try another file.

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in CContactReports extend(): %s\n%s\n%s\n%s", e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in CContactReports extend(): ', e.__doc__, '\n', lines[0], '\n', lines[1],'\n', lines[-1])
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
            nospc = rr.decimate_spaces(rpt)
            reduced = rr.decimate_commas(nospc)

            nospc = Path(nospc)
            if nospc.exists():
                nospc.unlink()

            lines = rr.lines_from_csv(reduced)

            reduced = Path(reduced)
            if reduced.exists():
                reduced.unlink()

            lengs = list()
            times = list()
            colutc = list()
            startstop = dict()
            contact = False #Kludge to include duration into the startstop dictionary.

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
                        # End If for Target match (satellite)

                        mobs = regeobsrvr.match(data)
                        if mobs:
                            """ Observer row contains the AOI and recurs for each set of contacts in the file. """

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
                        # End If for Observer Match (AOI or Ground Station)

                        match = regeutc.search(data)
                        """ The 'UTC' heading row identifies the start/stop columns and recurs for each data set in the file. """
                        if match:
                            if len(colutc) == 0:
                                colutc.append(col)
                                continue
                            elif len(colutc) == 1:
                                colutc.append(col)
                                continue
                        # End If for UTC Heading match

                        mdura = regedur.search(data)
                        if mdura:
                            if len(colutc) == 2:
                                colutc.append(col)
                                continue
                        # End If for Duration heading match

                        mtime = rr.regetime.match(data)
                        if mtime:
                            """ Time data occurs for each visibility between a Target and Observer."""
                            if col == colutc[0]:
                                starttime = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                                """ A  contact Start Time in Excel datetime form. """
                                continue
                            elif col == colutc[1]:
                                stoptime = dt.datetime.strptime(data, rr.dtdict['GMAT1'][3])
                                """ A contact Stop Time in Excel datetime form. """
                                contact = True
                                """ data = duration does not match the regetime pattern and match
                                    objects are not valid for conditional And and Or operators. 
                                    So we extent the contact time detection logic using this kludge. """
                                continue
                            else:
                                raise ValueError('Unexpected time field found in data column, %d.', col)
                        # End If for datetime pattern match

                        if contact:
                            if col == colutc[2]:
                                contact = False
                                duration = data
                                times.append([starttime, stoptime, duration])
                                continue
                        # End If for Kludge

                        mevts = regenumevt.match(data)
                        if mevts:
                            if len(times) > 0:
                                startstop.update(cp.deepcopy({key:times}))
                                """ Update start/stop times for the previous key each time 'Number of events' is found. """
                                times.clear()
                            continue

                        # End If for text occurring at end of contact
                    # End If data length > 0
                # End iteration over Contact Report row data
            # End iteration over lines from Contact Report

            print('Contact Start/Stop times determined for Report {0}.\n'.format(rpt.name))

            with xw.App() as excel: # This line takes significant I/O time.
                excel.visible=False  
                for key, contacts in startstop.items():
                    """ contacts is a 2 x 3 list of all startstop times associated with the given key.
                        Use the Contact Locator start/stop times to select data from the Link Report 
                        also associated with the key.
                    """
                    try:
                        linkfile = self.links[key]

                    except KeyError as e:
                        KeyError('Warning: Key %s Not found in self.links.  Continuing with next key.', key)
                        print('Key {0} not associated with Link Report file in links dictionary. Continuing.'.format(key))
                        continue
                    
                    if Path(linkfile).exists:
                        """ Get and open the Link ReportFile using xlwings. """
                        print('Opening input Link File {0} in XLWings for {1}.'.format(linkfile.name, key))

                        aoi = key.split('@')[1]
                        """ Worksheet keys are shortened to just the AOI. """
                        sheet = wb.get_worksheet_by_name(aoi)
                        """ The sheet to be written is named the same as the keys. """

                        print('Accessing {0} worksheet {1} for writing.'.format(xlfile.name, aoi))

                        try:
                            bk = excel.books.open(str(linkfile)) # This step requires a lot of time.

                            lrsheet = bk.sheets['Report']
                            """ The presence of the GMAT output report in a tab named 'Report'
                                is a mandatory condition. 
                            """
                        except pwin.com_error as ouch:
                            logging.warning('Attempted access to worksheet \'Report\' raised Windows com error. {0}, {1}'\
                                .format(type(ouch), ouch.args[1]))
                            print('Warning: attempted access to worksheet \'Report\' raised Windows com error. {0}, {1}'\
                                .format(type(ouch), ouch.args[1])) 
                            continue
                        
                        lreprng = lrsheet.range('A1').expand()
                        a1gregs = lreprng.columns[0].value 
                        """ By default aigregs is a List. It is ordered by datetime. """
                        
                        data = lreprng.rows[0].value
                        """ Headings List:
                            cellval = data[0] is A1Gregorian
                            cellval = data[2] is Earth Fixed Planetodetic LAT
                            cellval = data[3] is Earth Fixed Planetodetic LON
                            cellval = data[4] is Earth Altitude
                            cellval = data[5] is [AOI] X
                            cellval = data[6] is [AOI] Y
                            cellval = data[7] is [AOI] Z
                        """
                        
                        data.append('Slant.Range.(km)')
                        data.append('Azimuth.(deg)')
                        data.append('Elevation.(deg)')
                        """ Headings for custom formulas. """
                        writerow = 0
                        for col, cellval in enumerate(data):
                            cellval, leng = rr.heading_row(cellval)
                            lengs.append(leng)
                            sheet.set_column(col, col, leng)
                            sheet.write(writerow, col, cellval, cell_heading)
                        # End iteration over row 0 for headings

                        for times in contacts:
                            writerow += 1
                            for col, cellval in enumerate(times):
                                if col == 0:
                                    starttime = cellval
                                    rowstart = bi.bisect_left(a1gregs, starttime)
                                    """ One row before the start point in the Link Report data range. """

                                    sheet.write(writerow, col, starttime, cell_datetime)
                                    sheet.write(writerow, col+1, 'Start', cell_wrap)
                                if col == 1:
                                    stoptime = cellval
                                    rowstop = bi.bisect_right(a1gregs, stoptime)
                                    """ One row after the stop time in the Link Report data range. """
                                if col == 2:
                                    duration = cellval
                                    """ Preserve the SPICE computed duration of contact from the Event Locator. """
                                # End column cases
                            # End iteration over rows of Link Report A1 Gregorian

                            for row in range(rowstart, rowstop):
                                """ Write out the Link Report attributes between the start/stop times."""
                                
                                data = lreprng.rows[row].value
                                basedata = len(data)

                                writerow += 1
                                formrow = writerow +1
                                """ Excel Rows are 1-based. """
                                data.append('=SQRT(E{0}^2+F{0}^2+G{0}^2)'.format(formrow))
                                """Formula for Slant Range (km)"""
                                data.append('=DEGREES(ATAN(F{0}/E{0}))'.format(formrow))
                                """Formula for Azimuth (deg)"""
                                data.append('=DEGREES(ATAN(G{0}/(SQRT(E{0}^2+F{0}^2))))'.format(formrow))
                                """Formula for Elevation (deg)"""
                                for col, cellval in enumerate(data):            
                                    if col == 0:
                                        """ Ephemeris """
                                        cellstr = cellval.strftime(rr.dtdict['GMAT1'][3]) 
                                        leng = len(cellstr) * 0.85
                                        if len(lengs) < col + 1:
                                            lengs.append(leng)
                                        if leng > lengs[col]:
                                            lengs[col] = leng
                                            """ Only update the column width if current data is longer than previous. """
                                        else:
                                            leng = lengs[col]
                                            
                                        sheet.set_column(col, col, leng)
                                        sheet.write(writerow, col, cellval, cell_datetime)

                                        continue

                                    elif col in range(1, basedata):
                                        """ Attributes """
                                        cellstr = '{: 0.4f}'.format(cellval)
                                        leng = len(cellstr) + 1
                                        if len(lengs) < col + 1:
                                            lengs.append(leng)
                                        if leng > lengs[col]:
                                            lengs[col] = leng
                                            """ Only update the column width if current data is longer than previous. """
                                        else:
                                            leng = lengs[col]
                                            
                                        sheet.set_column(col, col, leng)
                                        sheet.write(writerow, col, cellval, cell_4plnum)

                                        continue

                                    elif col >= basedata:
                                        """ Write the custom formulas on this row, starting with nextcolumn"""
                                        leng = 10 # We do not want the column length = length of the formula string.
                                        if len(lengs) < col + 1:
                                            lengs.append(leng)
                                        if leng > lengs[col]:
                                            lengs[col] = leng

                                        sheet.set_column(col, col, leng)
                                        sheet.write_formula(writerow, col, cellval, cell_4plnum)

                                        continue

                                    else:
                                        """ Should be impossible to reach, but if so, that's a bad thing. """
                                        raise IndexError('Column {0} is out of range for {1}'.format(col, Path(linkfile).name))
                                    # End column cases
                                # End iteration over row columns
                            # End iteration over start/stop range
                                  
                            writerow += 1
                            sheet.write(writerow, 0, stoptime, cell_datetime)
                            sheet.write(writerow, 1, 'Stop', cell_wrap)
                            sheet.write(writerow, 2, 'Duration:', cell_wrap)
                            sheet.write(writerow, 3, duration, cell_2plnum)
                            sheet.write(writerow, 4, 'secs', cell_wrap)
                    # End iteration over all Contact Locator Report rows

                    sheet.freeze_panes('A2')
                    """ Lock the first row, first column after formatting of all rows and columns is done. """                               
                # End iteration over contact file

                logging.info('CContactReports extend() completed for filename: %s.', rpt.name)
                print ('CContactReports extend() completed  for filename:', rpt.name, '\n')
                
            # Close workbook
            return

        except OSError as e:
            lines = traceback.format_exc().splitlines()
            logging.error("OS error: %s\n%s\n%s\n%s in CContactReports extend() for filename %s", e.strerror, e.filename,\
                lines[0], lines[1], lines[-1])
            print('OS error: ', e.strerror,' in CContactReports extend() for filename ', e.filename,\
                '\n',lines[0], '\n',lines[1], '\n',lines[-1])

        except ValueError as e:
            lines = traceback.format_exc().splitlines()
            logging.error('%s, Incompatible input value, %s.\n%s\n%s\n%s', rpt.name, e.args[0],lines[0], lines[1], lines[-1])
            print(rpt.name, ', Incompatible input value, ',e.args[0], '\n', lines[0], '\n', lines[1],'\n', lines[-1])

        except IndexError as e:
            lines = traceback.format_exc().splitlines()
            logging.error('%s, %s \n%s \n%s \n%s',e.__doc__, e.args[0], lines[0], lines[1], lines[-1])
            print(e.__doc__, e.args[0], '\n', lines[0], '\n', lines[1],'\n', lines[-1])

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
        print('Link Reports completed.\n')

        fname = QFileDialog().getOpenFileName(None, 'Open (SIGHT) Contact Locater Reports batch file.', 
                        o_path,
                        filter='Batch files(*.batch)')

        logging.info('Contact Locater Report batch file is %s', fname[0])

        sightfile = Path(fname[0])

        cr_inst = CContactReports()
        cr_inst.setlinks(lr_inst.links)     
        cr_inst.do_batch(sightfile)

        logging.info ('Contact Reports Completed.')
        print('Contact Reports Completed.')

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s caught at top level:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[-1])
        
    finally:
        qApp.quit()

    