#! python
# -*- coding: utf-8 -*-
"""
    @file gmat_batcher.py

    @brief: This routine reads the batchfile created by modelgen.py, 
    and executes GMAT in command line mode for each filename found in the batchfile.  

    @copyright: Copyright (C) 2019 - 2022 Freelance Rocket Science, All rights reserved.

    @author  Colin Helms, colinhelms@outlook.com, [CCH]

    @details:
    Assumption 1: GMAT.exe is in the executable path.
    Assumption 2: A batchfile exists in the user's directory 
        consisting of a line-by-line list of model filenames to be executed.
    Assumption 3: The user's platform is capable of executing each model file
        with 5 minutes.  If not, change the cpto global variable below.
    
    @remark Change History
        Wed Feb  6 2019, Created.
        10 Jan 2019, [CCH] commit to GitHub repository GMAT-Automation, Integration Branch.
        01 Mar 2019, [CCH] Multiprocessing Enhancement.
        03 Jun 2019, [CCH] cleanup by quitting the QtApp.
        31 July 2019, [CCH] Added progress indicator using Manager.Queue. Updated cpto to 315.
        Wed Apr 20 14:54:49 2022, [CCH] reorganized and included in sdist
        Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.

    @bug https://github.com/a093130/GMAT-Automation/issues                
"""
from __future__ import division
""" Must be at the top of the file - maps features supported in future python versions """
import os
import sys
import time
import re
import platform
import logging
import traceback
import getpass
import random
import subprocess as sp
from multiprocessing import Pool
from multiprocessing import cpu_count
from multiprocessing import Manager
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)

cpto = 315
""" Child process timeout = 5 minutes 15 secs: more than sufficient on dual 2.13GHz E5506 XEON, 
16 Gbyte workstation with GTX 750 GPU. For the JS&R paper, run1 07/25/2019, a maximum GMAT run
time of 294 seconds was recorded for Batch_64HET6342W_36000.0kg_22Sep2020_R6.13_51.6deg_-0.832_J206_235431.script.
Change 31 Jul 2019: Was cpto = 300, Is cpto = 315.
"""
rsrv_cpus = 2
""" Reserve 2 cores for system processes and services (daemons). Spikes on process context swap. """

def delay_run():
    """ Helper to randomize start of child processes. Minimizes GMAT log file collisions. """
    numerator = random.randrange(1,6,1)
    denominator = random.randrange(7,12,1)
    delay = round(numerator/denominator,3)
    time.sleep(delay)
    
def run_gmat(args):
    """ GMAT wrapper to allow multiprocess.Pool to parallelize the execution of GMAT.
    Input arguments are contained in a list as follows:
        gmat_arg[0] is the GMAT script file path,
        gmat_arg[1] is the managed output queue connecting the child process to the main process.
    """
    delay_run()
    
    q = args[1]
    """ Output Queue """
    
    scriptname = os.path.basename(args[0])
        
    try:
        proc = sp.Popen(['gmat', '-m', '-ns', '-x', '-r', str(args[0])], stdout=sp.PIPE, stderr=sp.STDOUT)   
        """ Run GMAT for path names passed as args[0].  The GMAT executable must be in the system path.
        GMAT Arguments:
        -m: Start GMAT with a minimized interface.
        -ns: Start GMAT without the splash screen showing.
        -x: Exit GMAT after running the specified script.
        -r: Automatically run the specified script after loading.
        """
        pid = proc.pid
        (outs, errors) = proc.communicate(timeout=cpto)
        """ The buffer passed to Popen() defaults to io.DEFAULT_BUFFER_SIZE, usually 62526 bytes.
        If this is exceeded, the child process hangs with write pending for the buffer to be read.
        https://thraxil.org/users/anders/posts/2008/03/13/Subprocess-Hanging-PIPE-is-your-enemy/
        
        Attempt to maintain the buffer by reading it frequently, the timeout
        value should be long enough for GMAT to complete.  Check the GMAT output from 
        communicate() to be certain.
        """
        outs = outs.decode('UTF-8')
        message = 'From PID: {0} '.format(pid) + filter_outs(outs, scriptname)
        q.put(message)
                              
    except sp.TimeoutExpired as e:
        """ This function is meant to be called in the multiprocess context.  Logging
        threads are dangerous, because the thread context from the parent process is not passed
        to the child process.  Logging must be done in the parent process.
        """
        q.put("GMAT: Timeout Expired, File: %s" % scriptname)
                            
    except sp.CalledProcessError as e:       
        q.put("GMAT: Called ProcessError, File: %s" % scriptname)  
    
    except sp.SubprocessError as e:
        q.put("GMAT: Subprocess Error, File: %s" % scriptname)
        
    except Exception as e:
        q.put("GMAT: Unanticipated Exception " + e.__doc__ + ", File: " + scriptname)
        
    finally:                          
        (outs, errors) = proc.communicate()
        """ Flush the stdout buffer. """
        
        message = 'Terminating PID: {0} '.format(pid) + scriptname
        q.put(message)
        
        proc.terminate()
        """ Was 'proc.kill()' but proc.terminate() is better on Windows. """        
 
def filter_outs(outs:str, id:str):
    """ Reduce the logging size of the gmat output message.
    
    Parameters:
        UTF-8 decoded message
        id string, recommend the scriptname for id, but could be the PID.
    """
    loglines = outs.split()
    loglines = loglines[-20:]
    
    outs = " ".join(loglines)
    loglines = id + "\n" + outs
    
    rege = re.compile("====")
    loglines = rege.sub("", loglines)
    
    return loglines

if __name__ == "__main__":
    """ Retrieve the batch file and run GMAT for each model file listed """
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
    """
    logging.basicConfig(
            filename='./BatcherLog.log',
            level=logging.DEBUG,
            format='%(asctime)s %(filename)s %(levelname)s:\n%(message)s', datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! GMAT Batch Execution Started !!!!!!!!!!")
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)
    
    QApp = QApplication([])
    QApp.processEvents()
    
    fname = QFileDialog().getOpenFileName(None, 'Open master batch file.', 
                       os.getenv('USERPROFILE'))
        
    logging.info('Batch file is %s', fname[0])
    
    gmat_arg = ()
    gmat_args = list()
    ncpus = cpu_count()
    nrunp = ncpus - rsrv_cpus
#    ninstances = 4
    ninstances = 1
    nmsg = 0
    
    mgr = Manager()
    task_queue = mgr.Queue()
    stat_queue = mgr.Queue()
    
    print('Please wait, setting up jobs in {0}.'.format(os.path.basename(fname[0])))
    
    try:
        with open(fname[0]) as f:
            """ This is the master batch file selected in QtFileDialog. """
            numjobs = 0
            for filename in f:
                """ It must be cleaned up for GMAT to recognize it. """
                gmat_arg = os.path.normpath(filename)
                rege = re.compile('\n')
                gmat_arg = rege.sub('', gmat_arg)

                logging.debug("Path to scriptfile is %s", gmat_arg)
                scriptname = os.path.basename(filename)
                
                gmat_args.append([gmat_arg, task_queue])
                
                numjobs += 1
                        
        progress = QProgressDialog("{0} Jobs ...".format(numjobs), "Cancel", 0, numjobs)
        progress.setWindowTitle('Executing Batch')
        progress.setValue(0)
#        progress.show()
                                  
        pool = Pool(processes=nrunp, maxtasksperchild=20)
        """ In the single process execution of GMAT it was found that the process would
        timeout after a max of 20 processes.
        """
        
        """ Multi-processing Progress Bar. See, 
        https://stackoverflow.com/questions/5666576/show-the-progress-of-a-python-multiprocessing-pool-map-call
        by jfs, note further comment about difference between map and imap by simonmacmullen.
        """
        progress_ctr = 0
        for rs, _ in enumerate(pool.imap_unordered(run_gmat, gmat_args, 1)):
            """ rs is an instance of class multiprocessing.pool.AsyncResult """
            
            progress_ctr = rs + 1
#           The Qt progress counter does not update. Needs to be modeless.
#           https://stackoverflow.com/questions/38309803/pyqt-non-modal-dialog-always-modal
#            progress.setValue(progress_ctr)

            print('Completed {0} of {1} tasks.'.format(progress_ctr, numjobs))
        
        pool.close()
        
        print('Reading Job output queue, finishing up.')
        
        while(True):           
            qout = task_queue.get(cpto)
#            
            logging.info('GMAT reports: {0}'.format(qout))
#                        
            if task_queue.qsize() < 1:
                break
        
    except RuntimeError as e:
        lines = traceback.format_exc().splitlines()
        logging.error("RuntimeError: %s\n%s", lines[0], lines[-1])
        
    except ValueError as e:
        lines = traceback.format_exc().splitlines()
        logging.error("ValueError: %s\n%s", lines[0], lines[-1])
        
    except AttributeError as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Attribute Error: %s\n%s", lines[0], lines[-1])
        
    except OSError as e:
        logging.error("OS error: %s for filename %s", e.strerror, e.filename)

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
                            
    finally:
        QApp.quit()
        logging.info("!!!!!!!!!! GMAT Batch Execution Completed !!!!!!!!!!\n\n")
        sys.exit()
            
            
            
            
            
