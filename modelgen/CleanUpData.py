#! Python
# -*- coding: utf-8 -*-
"""
Created on 29 Mar 2022

@author: colinhelms@outlook.com

@description: module container for class definition CleanUpData.


@Copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.

XlWings Copyright (C) Zoomer Analytics LLC. All rights reserved.
https://docs.xlwings.org/en/stable/license.html
   
@change:
    
"""
import os
import time
import re
import platform
import logging
import traceback
import getpass
import csv
from pathlib import Path
from pathlib import PurePath
import datetime as dt
import xlsxwriter as xwrt
import xlsxwriter.utility as xlut
from gmatlocator import CGMATParticulars
import CleanUpReports
import reduce_report as rr
from PyQt5.QtWidgets import(QApplication, QFileDialog, QProgressDialog)


class CleanUpData(CleanUpReports):
    """ Provides methods to decimate spaces and commas in text files, creates a dictionary. """

    def __init__(self):
        self.data = {}
        """ In this derived class extend is overloaded to create a data dictionary from the input
            file.  This instance attribute may be used to return the current data.
        """
        
    def extend(self, rpt):
        try:
            nospc = rr.decimate_spaces(rpt)
            reduced = rr.decimate_commas(nospc)
            datadict = rr.lines_from_csv(reduced)

            nospc = Path(nospc)
            if nospc.exists():
                nospc.unlink()

            reduced = Path(reduced)
            if reduced.exists():
                reduced.unlink()
        
            logging.info('Cleaned up file: {0}, row data returned in dictionary'.format(rpt))

            self.data.clear()
            """ extend() may be used in the parent do_batch(), the following will overwrite data with the same keys.
                The keys are row numbers in the global implementation of lines_from_csv().
            """
            self.data.update(datadict)
            """ Note that the keys in this dictionary are row numbers"""

            return datadict

        except OSError as e:
            logging.error("OS error: %s for filename %s", e.strerror, e.filename)

        except Exception as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception: %s\n%s\n%s", e.__doc__, lines[0], lines[-1])
