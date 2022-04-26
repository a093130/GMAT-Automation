#! Python
# -*- coding: utf-8 -*-
"""
	@file CleanUpData.py
	@brief File is module container for class definition CCleanUpData.
	
	@copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

	@author  Colin Helms, colinhelms@outlook.com, [CCH]

    @details Provides minimal ReportFile formatting that yields a memory resident dictionary  
        rather than an Excel workbook file.
   
    @remark Change History
        Fri 29 Mar 2022, [CCH] File created, GitHub repository GMAT-Automation.
        Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.
    
    @bug https://github.com/a093130/GMAT-Automation/issues
"""
import logging
import traceback
from pathlib import Path
from gmatautomation import reduce_report as rr
from gmatautomation import CCleanUpReports
from gmatautomation import CGmatParticulars

class CCleanUpData(CCleanUpReports): # syntax is module.class.
    """ Base Class. Downstream processing of the generated dictionary may be specialized 
        using work_on(). File processing may also be modified using extend().  
        The default work_on() simply prints the size of each input dictionary.
        Note that the dictionary size may be very large.
    """

    def __init__(self, **args):
        super().__init__(**args)
        return


    def do_batch(self, batchfile, **args):
        """Call parent class do_batch() but the LinkReports.extend function should be called. """

        super().do_batch(batchfile, **args)
        """ Delegate up the MRO chain.
            See: https://stackoverflow.com/questions/32014260
            See: https://rhettinger.wordpress.com/2011/05/26/super-considered-super
        """
        return

    def extend(self, rpt):
        try:
            nospc = rr.decimate_spaces(rpt)
            reduced = rr.decimate_commas(nospc)
            data = rr.lines_from_csv(reduced)

            nospc = Path(nospc)
            if nospc.exists():
                nospc.unlink()

            reduced = Path(reduced)
            if reduced.exists():
                reduced.unlink()
        
            logging.info('Cleaned up file: {0}, row data returned in dictionary'.format(rpt))

            self.work_on(data)

            return

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)
            print('OS error',  e.strerror, ' for filename', e.filename)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Exception %s :\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception', e.__doc__, ':\n', lines[0], '\n', lines[1], '\n', lines[-1])

        
    def work_on(self, linedict: dict):
        """ Virtual function to enable specialized processing of lines_from_csv() output dictionary. """
        try:
            if isinstance(linedict, dict):
                print('Argument type for work_on() is validated as Dictionary.')
                print('size of dictionary:', len(linedict))
            else:
                print('Invalid parameter, is not Dictionary. Type of parameter is ', linedict.__class__)

            return

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)
            print('OS error',  e.strerror, ' for filename', e.filename)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Exception %s :\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception', e.__doc__, ' :\n', lines[0], '\n', lines[1], '\n', lines[-1])
            


if __name__ == "__main__":
    """ Unit Tests for module. """
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
    

    gmat_paths = CGmatParticulars()
    o_path = gmat_paths.get_output_path()
    """ o_path is an instance of Path that locates the GMAT output directory. """

    qApp = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open REPORT batchile.', 
                    o_path,
                    filter='text files(*.batch)')

    logging.info('Input report file is %s', fname[0])

    reports = CCleanUpData()
    batchfile = Path(fname[0])

    try:
        """ Test Case: Repeats Test Case 2 in reduce_report by calling class instance instead."""
        reports.do_batch(batchfile)

        logging.info('Test Case for CleanUpData complete.')
        print('Test Case for CleanUpData complete.')

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s :\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception', e.__doc__, ' :\n', lines[0], '\n', lines[1], '\n', lines[-1])

    finally:
        qApp.quit()