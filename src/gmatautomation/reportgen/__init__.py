#!python
# # -*- coding: utf-8 -*-
"""
@Name: gmatautomation 
@Version: v0.3b1

@author: Colin Helms
@author_email: colinhelms@outlook.com

@description: This package contains procedures to autoformat
various types of reports fro raw GMAT ReportFiles and Contact Locator files.

"""
from __future__ import absolute_import

import sys

__all__ = ["batch_alfano_rep", "reduce_report", "CleanUpData", "CCleanUpData", "CleanUpReports", "CCleanUpReports",\
"ContactReports", "CContactReports", "LinkReports", "CLinkReports", "LinkBudgets", "CLinkBudgets"]

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required. (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys