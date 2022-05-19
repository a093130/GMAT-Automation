#! Python
# -*- coding: utf-8 -*-
""" 
	@file DownlinkCapacity.py
	@brief module container for class definition CDownlinkCapacity.
	
	@copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.
	@author  Colin Helms, colinhelms@outlook.com, [CCH]

	@details Class def Downlink Capacity extends the class CLinkBudgets to insert useful Excel
    formulas for calculating Image Download volume and onbaord storage needs.  
    The parent class provides virtual methods
    moreheadings(self) and moreformulas(self, row, aoi) to be extended herein.

	@remark Change History
        18 May 2022, [CCH] File created, repository GMAT-Automation.
        19 May 2022, [CCH] Excel Formulas incorporated and (infomally) Validated
"""
import logging
import traceback
from pathlib import Path
from gmatautomation import CGmatParticulars
from gmatautomation import CLinkBudgets
#from LinkBudgets import CLinkBudgets # for development only
from gmatautomation import CLinkReports


class CDownlinkCapacity(CLinkBudgets):

    def __init__(self, **args):
        super().__init__(**args)

    def moreheadings(self):
        """ further specialization of formula headings used in Link Budgets."""
        data = list()

        data.append(r'NITF MPEG2000 Data Transfer Rate (bps)')
        data.append(r'Elapsed Time (Sec)')
        data.append(r'Pixel Xfr Rate (pixels/s)')
        data.append(r'Pixels Xfr')
        data.append(r'Full Res Images Xfr')
        data.append(r'Total Images D/L')
        data.append(r'Max  Collect Pixels')
        data.append(r'Residual Image  pixels')
        data.append(r'Storage Requirement (bytes)')
        

        return data

    def moreformulas(self, formrow, aoi):
        """ further specialization of formula headings used in Link Budgets.
            Do not increment formrow.
        """
        data = list()

        """ The cell format for the resultant values are returned as an array index:
            cellformats[0] = cell_heading
            cellformats[1] = cell_datetime
            cellformats[2] = cell_wrap
            cellformats[3] = cell_4plnum
            cellformats[4] = cell_2plnum
            cellformats[5] = cell_1digint
            cellformats[6] = cell_sep3digint

            DO NOT put a space before the equal sign in the formula.
        """
        data.append([r'=$Z{0} * Compression_ratio'.format(formrow), 6])
        """Formula for NITF MPEG2000 Data Transfer Rate (bps)"""

        previousrow = formrow - 1

        data.append([r'=3600*(HOUR($A{0})-HOUR($A{1}))+60*(MINUTE($A{0})-MINUTE($A{1}))+(SECOND($A{0})-SECOND($A{1}))'.format(formrow, previousrow), 4])
        """Formula for Elapsed Time (Sec) - Tricky one here, PyLint wants to correct parenthetical expressions"""

        data.append([r'= $AA{0}/Bits_per_Pixel'.format(formrow), 6])
        """Formula for Pixel Xfr Rate (pixels/s)"""

        data.append([r'=$AC{0} * $AB{0}'.format(formrow), 6])
        """Formula for Pixels Xfr"""

        data.append([r'=$AD{0}/Pixels_per_FR'.format(formrow), 4])
        """Formula for Full Res Images Xfr"""

        data.append([r'=$AE{0} + $AF{1}'.format(formrow, previousrow), 4])
        """Formula for Total Images"""

        data.append([r'= FRI_Pixel_Rate * AB{0}'.format(formrow), 6])
        """Max  Collect Pixels"""

        data.append([r'= $AG{0} - $AD{0}'.format(formrow), 6])
        """Residual Image  pixels"""

        data.append([r'=$AH{0} + $AI{1}'.format(formrow, previousrow), 6])
        """Storage Requirement (bytes)"""        
        return data

if __name__ == "__main__":
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
        """
    import getpass
    import platform
    from PyQt5.QtWidgets import(QApplication, QFileDialog)

    logging.basicConfig(
            filename='./DownlinkCapacity.log',
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
        gmat_paths = CGmatParticulars()
        o_path = gmat_paths.get_output_path()
        """ o_path is an instance of Path that locates the GMAT output directory. """

        qApp = QApplication([])

        fname = QFileDialog().getOpenFileName(None, 'Open Link Reports batch file', 
                        str(o_path),
                        filter='text files(*.batch)')

        logging.info('Link Report batch file is %s', fname[0])

        batchfile = Path(fname[0])

        lr_inst = CLinkReports()
        lr_inst.do_batch(batchfile)

        logging.info ('Link Reports completed.')
        print('Link Reports completed.\n')

        fname = QFileDialog().getOpenFileName(None, 'Open (SIGHT) Contact Locater Reports batch file.', 
                        str(o_path),
                        filter='Batch files(*.batch)')

        logging.info('Contact Locater Report batch file is %s', fname[0])

        sightfile = Path(fname[0])

        cr_inst = CDownlinkCapacity()
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