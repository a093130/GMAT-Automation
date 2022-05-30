#!python
# # -*- coding: utf-8 -*-
"""
@name: gmatautomation
@version: v0.3b1

@author: Colin Helms
@author_email: colinhelms@outlook.com

@description: This package contains a GMAT batchfile execution procedure.

"""
from __future__ import absolute_import

import sys

__all__ = ["gmat_batcher"]

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required. (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys