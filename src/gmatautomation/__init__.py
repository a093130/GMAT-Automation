#!python
# # -*- coding: utf-8 -*-
"""
@Name: gmatautomation
@Version: 0.3b2

@author: Colin Helms
@author_email: colinhelms@outlook.com
"""
from __future__ import absolute_import

import sys

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required. (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys

""" The following must be imported in order of dependency. """
from .modelgen import gmatlocator
from .modelgen.gmatlocator import CGmatParticulars
from .modelgen import modelpov
from .modelgen import modelspec
from .modelgen.modelspec import CModelSpec
from .modelgen import fromconfigsheet
from .modelgen import modelgen
from .modelgen import userexceptions
from .modelgen.userexceptions import Ultima
from .modelcontrol import gmat_batcher
from .reportgen import reduce_report

from .reportgen import CleanUpReports
from .reportgen.CleanUpReports import CCleanUpReports
from .reportgen import CleanUpData
from .reportgen.CleanUpData import CCleanUpData
from .reportgen import LinkReports
from .reportgen.LinkReports import CLinkReports
from .reportgen import ContactReports
from .reportgen.ContactReports import CContactReports
