# -*- coding: utf-8 -*-
"""
Created on Wed Feb  6 19:17:02 2019

@author: colinhelms@outlook.com

@Description:
    This routine reads the batchfile created by modelgen.py, and executes GMAT in
    command line mode for each filename found in the batchfile.
    
    Assumption 1: GMAT.exe is in the executable path.
    Assumption 2: A batchfile exists in the user's directory 
        consisting of a line-by-line list of model filenames to be executed.
    Assumption 3: The user's platform is capable of executing each model file
        with 5 minutes.  If not, change the cpto global variable below.
    
@Change Log:
    10 Jan 2019, Initial baseline, Integration branch.
                 
"""

import os
import sys
import re
#import shlex
import platform
import logging
import traceback
#import tempfile
#import threading as task
import getpass
from pathlib import Path
import subprocess as sp
from PyQt5.QtWidgets import(QApplication, QFileDialog)

cpto = 300
""" Child process timeout = 10 minutes: more than sufficient on dual 2.13GHz E5506 XEON, 
16 Gbyte workstation with GTX 750 GPU 
"""
def parasite(proc):
    """ Read stdout from the proc and write lines to given file """
    with open ("c:\\temp\\gmat_stdout.log",'+a') as fout:
        for line in iter(proc.stdout.readline, b''):
            fout.write('got line: {0}'.format(line.decode('utf-8')))
        
class GMAT_Path:
    """ This class initializes its instance with the GMAT root path using the 
    'LOCALAPPDATA' environment variable.  Current version is Windows specific.
    """
    def __init__(self):
        logging.debug('Instance of class GMAT_Path constructed.')
        
        self.p_gmat = os.getenv('LOCALAPPDATA')+'\\GMAT'
        self.executable_path = ''
                       
    def get_executable_path(self):
        """ Convenience function which searches for GMAT.exe. """
        logging.debug('Method get_executable_path() called.')
        
        p = Path(self.p_gmat)
        
        gmat_ex_paths = list(p.glob('**/GMAT.exe'))
        
        self.executable_path = str(gmat_ex_paths[0])
        """ Initialize startup_file path. """
        
        for pth in gmat_ex_paths:
            """ Where multiple GMAT.exe instances are found, use the last modified. """          
            old_p = Path(self.executable_path)
            old_mtime = old_p.stat().st_mtime
            
            p = Path(pth)
            mtime = p.stat().st_mtime

            if mtime - old_mtime > 0:
                self.executable_path = str(pth + 'GMAT.exe')
            else:
                continue

        logging.info('The GMAT executable path is %s.', self.executable_path)
        
        return self.executable_path

if __name__ == "__main__":
    """ Retrieve the batch file and run GMAT for each model file listed """
    logging.basicConfig(
            filename='./appLog.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', datefmt='%d%B%Y_%H:%M:%S')

    logging.info('******************** GMAT Batch Execution Started ********************')
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
    app = QApplication([])
    
    fname = QFileDialog().getOpenFileName(None, 'Open master batch file.', 
                       os.getenv('USERPROFILE'))
        
    logging.info('Batch file is %s', fname[0])
    
    gmat_args = ()
    
    try:
        with open(fname[0]) as f:    
            for filename in f:
                gmat_args = os.path.normpath(filename)
                rege = re.compile('\n')
                gmat_args = rege.sub('', gmat_args)

                logging.debug("Path to scriptfile is %s", gmat_args)
                scriptname = os.path.basename(filename)
                logging.info("GMAT will be called for script %s", scriptname)
               
                proc = sp.Popen(['gmat', '-m', '-ns', '-x', '-r', str(gmat_args)], stdout=sp.PIPE, stderr=sp.STDOUT)   
                """ Run GMAT for each file in batch.
                    Arguments:
                    -m: Start GMAT with a minimized interface.
                    -ns: Start GMAT without the splash screen showing.
                    -x: Exit GMAT after running the specified script.
                    -r: Automatically run the specified script after loading.
                Note: The buffer passed to Popen() defaults to io.DEFAULT_BUFFER_SIZE, usually 62526 bytes.
                If this is exceeded, the child process hangs with write pending for the buffer to be read.
                https://thraxil.org/users/anders/posts/2008/03/13/Subprocess-Hanging-PIPE-is-your-enemy/
                """
                try:
                    (outs, errors) = proc.communicate(cpto)
                    """Timeout in cpto seconds if process does not complete."""
                    
                except sp.TimeoutExpired as e:
                    logging.error('GMAT timed out in child process. Time allowed was %s secs, continuing', str(cpto))
                    
                    logging.info("Process %s being terminated.", str(proc.pid))
                    proc.kill()
                    """ The child process is not killed by the system. """
                    
                    (outs, errors) = proc.communicate()
                    """ And the stdout buffer must be flushed. """
                           
    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except sp.CalledProcessError as e:       
        logging.error('GMAT error code, %s returned.\nError is: %s.', e.returncode, e.stderr)
    
    
    except sp.SubprocessError as e:
        logging.error('GMAT generic child process error %s.', e.args)
        
    except ValueError as e:
        logging.error('Subprocess called with incorrect arguments: %s.', e.args)
        
    except AttributeError as e:
        tb = sys.exc_info()
        lines = traceback.format_exc().splitlines()
        logging.error('%s, Cause: %s, Context: %s\n%s%s', e.__doc__, e.__cause__, e.__context__, lines[0], lines[-1])

    except Exception as e:
        tb = sys.exc_info()
        lines = traceback.format_exc().splitlines()
        logging.error('%s Cause: %s, Context: %s\n%s%s', e.__doc__,  e.__cause__, e.__context__, lines[0], lines[-1])

    except:
        logging.error("Unknown error:\n%s", sys.exc_info())

            
            
            
            
            
