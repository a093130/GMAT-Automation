#!python
# # -*- coding: utf-8 -*-
"""
@Name: GmatAutomation 
@Version: v0.2a

Created on Sun Apr 28 15:28:56 2019
Deployed to <PYTHONPATH>/GmatAutomation

@author: Colin Helms
@author_email: colinhelms@outlook.com

@Description: This package contains procedures to autoformat
various types of reports fro raw GMAT ReportFiles and Contact Locator files.

"""
from __future__ import absolute_import

import sys

__all__ = ["batch_alfano_rep", "reduce_report", "CleanUpData", "CleanUpReports", "ContactReports", "LinkReports", "LinkBudgets"]

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required for Alfano (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys

from .modelgen import gmatlocator