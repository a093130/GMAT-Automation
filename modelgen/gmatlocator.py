# -*- coding: utf-8 -*-
"""
Created on Tue Apr 30 16:20:54 2019

@author: Colin Helms

@Description:
    This module provides a base class that is useful for locating the GMAT executable
    directory.  
    Current version is Windows specific, since it uses environment variable %LOCALAPPDATA%.
    
@Change Log:
    30 Apr 2019, refactored from gmat_batcher.py.
"""
import os
import logging
from pathlib import Path

class CGmatPath:
    """ This class initializes its instance with the GMAT root path using the 
    'LOCALAPPDATA' environment variable.  
    """
    def __init__(self):
        logging.debug('Instance of class GMAT_Path constructed.')
        
        self.p_gmat = os.path.join(os.getenv('LOCALAPPDATA'), 'GMAT')
        #TODO: find another more portable way to do this,
        #LOCALAPPDATA is a Windows 7,8,10 dependency.
        #Perhaps glob through the system path.
        
        self.executable_path = None
 
    def get_root_path(self):
        return self.p_gmat                      

    def get_executable_path(self):
        """ A simple accessor method. """
        if self.executable_path == None:
            self.find_gmat()
        return self.executable_path
    
    def find_gmat(self):
        """ Method searches for GMAT.exe. """
        logging.debug('Method get_executable_path() called.')
        
        p = Path(self.p_gmat)
        
        gmat_ex_paths = list(p.glob('**/GMAT.exe'))
        
        if len(gmat_ex_paths) >= 1:
            self.executable_path = str(gmat_ex_paths[0])
            """ Initialize startup_file path. """
            
            for pth in gmat_ex_paths:
                """ Where multiple GMAT.exe instances are found, use the last modified. """          
                old_p = Path(self.executable_path)
                old_mtime = old_p.stat().st_mtime
                
                p = Path(pth)
                mtime = p.stat().st_mtime
    
                if mtime - old_mtime > 0:
                    self.executable_path = os.path.join(str(pth), 'GMAT.exe')
                else:
                    continue
    
            logging.info('The GMAT executable path is %s.', self.executable_path)
        else:
            logging.info('No GMAT executable path is found.')
        
        return self.executable_path