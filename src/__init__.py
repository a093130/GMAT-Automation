# -*- coding: utf-8 -*-
"""
@name: gmatautomation 
@version: 0.3b2

@author: Colin Helms
@author_email: colinhelms@outlook.com
"""
from __future__ import absolute_import

import sys

if sys.version_info[:2] < (3, 4):
    m = "Python 3.4 or later is required for Alfano (%d.%d detected)."
    raise ImportError(m % sys.version_info[:2])
del sys

from gmatautomation import modelcontrol
from gmatautomation import modelgen
from gmatautomation import reportgen
