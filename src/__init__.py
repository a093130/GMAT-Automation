#!python
# # -*- coding: utf-8 -*-
"""
@Name: GmatAutomation
@Version: 0.2a

Created on Sun Apr 28 15:28:56 2019
Deployed to <PYTHONPATH>/GmatAutomation

@author: Colin Helms
@author_email: colinhelms@outlook.com

"""
from __future__ import absolute_import

import sys

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required for Alfano (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys

from .modelgen import fromconfigsheet
from .modelgen import gmatlocator
from .modelgen import modelgen
from .modelgen import modelpov
from .modelcontrol import gmat_batcher
from .reportgen import reduce_report
from .reportgen import batch_alfano_rep
from .reportgen import CleanUpData
from .reportgen import CleanUpReports
from .reportgen import ContactReports
from .reportgen import LinkReports

