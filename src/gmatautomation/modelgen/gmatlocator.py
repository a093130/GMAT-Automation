# -*- coding: utf-8 -*-
"""
    @file: gmatlocator.py

    @brief: This module provides a base class that is useful for locating
     the GMAT executable directory.  The current version is Windows specific.

    @copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

    @author  Colin Helms, colinhelms@outlook.com, [CCH]

    @details The key to this utility class is the find_gmat() function, which starting with the
    %LOCALAPPDATA% directory, searches for gmat.exe and captures its parent file path. 
    
    @remark Change History
        Tue Apr 30 16:20:54 2019, File Created
        30 Apr 2019, refactored from gmat_batcher.py.
        26 May 2019, Include CGMATParticulars from modelgen.py. Used by reducereport, modelgen.
        Mar 08 2022, committed to GitHub repository GMAT-Automation.
        Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.

    @bug https://github.com/a093130/GMAT-Automation/issues   
"""
import os
import sys
import re
import logging
import traceback
from pathlib import Path

class CGmatParticulars():
    """ This class initializes its instance with the script output path taken from
    the gmat_startup_file.txt.
    """
    def __init__(self,**args):
        logging.debug('Instance of class GMAT_Particulars constructed.')
        super().__init__(**args)

        #self.p_gmat = Path(os.getenv('LOCALAPPDATA'))/'GMAT'
        self.p_gmat = Path(os.getenv('LOCALAPPDATA'))
        # Fix for gmatautomation project issue #4.
        self.executable_path = None
        self.startup_file_path = None
        self.output_path = None                     

    def find_gmat(self):
        """ Method searches for GMAT.exe. """
        logging.debug('Method get_executable_path() called.')
        
        p = self.p_gmat
        """ Starting place to look. """

        try:
            gmat_ex_paths = list(p.glob('**/GMAT.exe'.casefold()))
            """ Perform case insensitive match."""

            old_mtime = 0
            if len(gmat_ex_paths) >= 1:          
                for pth in gmat_ex_paths:
                    """ Where multiple GMAT.exe instances are found, use the last modified. """          
                      
                    mtime = pth.stat().st_mtime
        
                    if mtime - old_mtime > 0:
                        self.executable_path = pth
                        old_mtime = mtime
                    else:
                        continue
                # End iteration on gmat_ex_paths
                
                self.p_gmat = self.executable_path.parents[1]
                """ Also update startup_file path. """

                logging.info('The GMAT executable path is %s.', self.executable_path)
            else:
                logging.info('No GMAT executable path is found.')
                print('No GMAT executable path is found.')
                raise RuntimeError('No GMAT executable path is found.')
            
            return self.executable_path

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Exception %s caught at top level:\n%s\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[3], lines[-1])
            print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[3], '\n', lines[-1])

    def get_root_path(self):
        """ This path is initialized and then updated"""
        return self.p_gmat
     
    def get_executable_path(self):
        """ A simple accessor method. """
        if self.executable_path == None:
            return self.find_gmat()
        else:
            return self.executable_path
     
    def get_startup_file_path(self):
        """ The gmat_startup_file.txt is located in the same directory as GMAT.exe. """
        
        if self.startup_file_path == None:
            ex_file_path = self.get_executable_path()
            self.startup_file_path = ex_file_path.parents[0]/'gmat_startup_file.txt'
        else:
            pass

        return self.startup_file_path
        
    def get_output_path(self):
        """ The path defined for all manner of output in gmat_startup_file.txt """
        logging.debug('Method get_output_path() called.')
        
        regeoutpath = re.compile('^OUTPUT_PATH')
        regerelative = re.compile(r'\.\.[\\/][a-zA-Z]+')
        regecrlf = re.compile('\s*')

        su_path = self.get_startup_file_path()
                
        try:
            with open(su_path) as f:
                """ Extract path string text assigned to OUTPUT_PATH in file. """
                for line in f:
                    matchrelpath = regerelative.search(line)
                    if regeoutpath.match(line):
                        """ found a line that contains 'OUTPUT_PATH'."""
                        line = regecrlf.sub('', str(line))
                        if matchrelpath:
                            """ OUTPUT_PATH line contains a relative path that needs to be resolved. """
                            outp = matchrelpath[0]
                            """ Extract the output stem."""
                            rootpath = Path(su_path).parents[0]
                            """ Determine the absolute path for the start-up file. """
                            outp = rootpath/outp
                            self.output_path = outp.resolve()
                        else:
                            line = line.split('=')[1] # Fix to remove 'OutpuPath='
                            self.output_path = Path(line)
                    else:
                        continue
                    # endif
                # end iteration for lines in su_path

            logging.info('The GMAT output path is %s.', self.output_path)
            return self.output_path

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)
            print('OS error: ', e.strerror,' in gmatlocator get_output_path() for filename ', e.filename)
            sys.exit(-1)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error('Exception %s caught at top level:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[-1])
            sys.exit(-1)



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
            filename='./gmatlocator.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! GMAT locator Started !!!!!!!!!!")
    
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)

    try:
        locator = CGmatParticulars()

        outpath = locator.get_output_path()
        """ This call also executes get_startup_file_path(), get_executable_path() and find_gmat(). """

        print('Startup path {0}'.format(outpath))
        print ('CGmatParticulars instance variables:',\
            '\np_gmat: {0}\nexecutable_path: {1}\nstartup_file_path: {2}\noutput_path: {3}'\
            .format(locator.p_gmat, locator.executable_path, locator.startup_file_path, locator.output_path))

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s caught at top level:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[-1])