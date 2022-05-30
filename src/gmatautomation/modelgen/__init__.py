#!python
# # -*- coding: utf-8 -*-
"""
@Name: GmatAutomation 
@Version: v0.2a1

Created on Sun Apr 28 15:28:56 2019
Deployed to <PYTHONPATH>/GmatAutomation

@author: Colin Helms
@author_email: colinhelms@outlook.com

@Description: This package contains procedures to autogenerate
GMAT models from an Excel spredsheet configuration specification.

"""
from __future__ import absolute_import

import sys

__all__ = ["CGmatParticulars", "gmatlocator", "fromconfigsheet",  "modelgen", "modelpov"]

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys
