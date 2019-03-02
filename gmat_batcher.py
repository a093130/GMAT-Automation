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
    28 Feb 2019, Working Copy committed to Integration.
    01 Mar 2019, Multiprocessing Enhancement.
                 
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
import multiprocessing as mp
from PyQt5.QtWidgets import(QApplication, QFileDialog)

cpto = 300
rsrv_cpus = 2

""" Child process timeout = 10 minutes: more than sufficient on dual 2.13GHz E5506 XEON, 
16 Gbyte workstation with GTX 750 GPU 
"""
#def parasite(proc):
#    """ Read stdout from the proc and write lines to given file """
#    with open ("c:\\temp\\gmat_stdout.log",'+a') as fout:
#        for line in iter(proc.stdout.readline, b''):
#            fout.write('got line: {0}'.format(line.decode('utf-8')))

def run_gmat(gmat_arg):
    msg[0] = "Running GMAT with File: " + str(gmat_arg)
    msg[1] = ""
    
    proc = sp.Popen(['gmat', '-m', '-ns', '-x', '-r', str(gmat_arg)], stdout=sp.PIPE, stderr=sp.STDOUT)   
    """ Run GMAT for batch file path names as gmat_args.
    GMAT Arguments:
    -m: Start GMAT with a minimized interface.
    -ns: Start GMAT without the splash screen showing.
    -x: Exit GMAT after running the specified script.
    -r: Automatically run the specified script after loading.
    """
    try:
        """ The buffer passed to Popen() defaults to io.DEFAULT_BUFFER_SIZE, usually 62526 bytes.
        If this is exceeded, the child process hangs with write pending for the buffer to be read.
        https://thraxil.org/users/anders/posts/2008/03/13/Subprocess-Hanging-PIPE-is-your-enemy/
        This try block will attempt to maintain the buffer by reading it frequently, otherwise the timeout
        value should be long enough for GMAT to complete, the TimeoutExpired exeception allows 
        processing to continue under this assumption.
        """
        (outs, errors) = proc.communicate(cpto)
        """Timeout in cpto seconds if process does not complete."""
        
    except sp.TimeoutExpired as e:
        """ This function is meant to be called in the multiprocess context.  Logging
        threads are dangerous, because the thread context from the parent process is not passed
        to the child process.  Logging must be done in the parent process.
        """
        msg[1] += "GMAT timed out in child process " + str(proc.pid) + "."
        
        proc.kill()
        """ The child process is not killed by the system, so clean it up here."""
        (outs, errors) = proc.communicate()
        """ And the stdout buffer must be flushed. """
            
    finally:
        return(msg)
    
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
    __spec__ = None
    """ Tweak to get Spyder IPython to execute this code. See:
    https://stackoverflow.com/questions/45720153/python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
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
    
    gmat_arg = ()
    gmat_args = list()
    ncpus = mp.cpu_count()
    nrunp = ncpus - rsrv_cpus
    ninstances = 4
    
    try:
        with open(fname[0]) as f:
            """ This is the master batch file selected in QtFileDialog. """
            for filename in f:
                """ It must be cleaned up for GMAT to recognize it. """
                gmat_arg = os.path.normpath(filename)
                rege = re.compile('\n')
                gmat_arg = rege.sub('', gmat_arg)

                logging.debug("Path to scriptfile is %s", gmat_arg)
                scriptname = os.path.basename(filename)
                
                gmat_args.append(gmat_arg)
                                      
        with mp.Pool(processes=nrunp, maxtasksperchild=20 ) as workers:
            results_iter = workers.imap(run_gmat, gmat_args, ninstances)
                
        for msg in results_iter:
            logging.info(msg[0])
            if msg[1].len > 0:
                logging.error(msg[1])
            
    except sp.CalledProcessError as e:       
        msg[1] += "GMAT error code, " + e.returncode + " returned."
        msg[1] += "\nError is: " + e.stderr    
    
    except sp.SubprocessError as e:
        msg[1] += "Popen generic child process error, " + e.args + "."
        
    except ValueError as e:
        msg[1] += "Subprocess called with incorrect arguments: " + e.args + "."
        
    except AttributeError as e:
#        tb = sys.exc_info()      
        lines = traceback.format_exc().splitlines()
        msg[1] += "Cause: " + e.__doc__ + "Context: " + e.__cause__ + "Traceback:\n"
        msg[1] += lines[0] + lines[-1]

    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
#        tb = sys.exc_info()
        lines = traceback.format_exc().splitlines()
        logging.error('%s Cause: %s, Context: %s\n%s%s', e.__doc__,  e.__cause__, e.__context__, lines[0], lines[-1])
                
    except:
        logging.error("Unknown error:\n%s", sys.exc_info())
            
    finally:
        workers.close()
        workers.join()
        logging.info('******************** GMAT Batch Execution Completed ********************')
            
            
            
            
            
