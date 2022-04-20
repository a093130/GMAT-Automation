#!python
# # -*- coding: utf-8 -*-
"""
@Name: GmatAutomation
@Version: v0.2a

Created on Sun Apr 28 15:28:56 2019
Deployed to <PYTHONPATH>/GmatAutomation

@author: Colin Helms
@author_email: colinhelms@outlook.com

@Description: This package contains a GMAT batchfile execution procedure.

"""
from __future__ import absolute_import

import sys

__all__ = [gmat_batcher"]

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required for Alfano (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys