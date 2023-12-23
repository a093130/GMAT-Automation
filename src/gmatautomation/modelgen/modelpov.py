# -*- coding: utf-8 -*-
"""
    @file modelgen.py

    @brief: Provides a module container for the CModelPov class.  

	@copyright 2023 Astroforge-Incorporated
	@author  Colin Helms, [CCH]
	@authoremail colin@astroforge.io,colinhelms@outlook.com

    @version 0.4b0

    @details:  The implemented points of variation in the base class are listed here.
    Users may derive classes from CModelPov for other GMAT points of variation.

    The dictionary form is [heading: (top model name) (sub-model name).(resource)].

    @remark Change History
        Sat Nov 17 09:49:45 2018 Created
        17 Nov 2018, baseline set of model resources supporting paper AIAA-2018-4718
        09 Feb 2019, commit to GitHub repository GMAT-Automation, Integration Branch.
        10 Apr 2019, [CCH] Use GMAT user variable Costate to store Costates.  
            Removed kludge using 'EOTV.ID' to store payload mass.
            Associate 'EOTV.ID' with 'SID'.
            Use GMAT user variable 'PL_MASS' to store 'Payload'
        Wed Apr 20 14:54:49 2022, [CCH] reorganized and included in sdist
        26 Apr 2022 - [CCH] Version 0.2a1, Buildable package, locally deployable.
        05 May 2022 - [CCH] Build version 0.3b3 for PyPI and upload as open source.
        16 Jun 2022 - [CCH] https://github.com/a093130/GMAT-Automation/issues/1 (refactor pov)
        05 Dec 2023, [CCH] version 0.4b0, refactored to eliminate Alfano specializations

    @bug https://github.com/a093130/GMAT-Automation/issues
"""
class CModelPov:
    def __init__(self, **args):
        """  
            The CModelPov class is a parent class container for derived classes which will
            encapsulate the Points of Variation for GMAT modelgeneration. 
            It is used by "fromconfigsheet.py" and "modelgen.py.    
        """

        self.mission_name = 'Mission'

        self.model_name = 'Mission_model'

        """ Specialized classes will create a unique root for each generated batch file and report file. """
        self.nameroot =''

        """ This dictionary identifies GMAT Worksheets. """
        self.sheetnames = dict()

        """ List of GMAT file paths associated with specialized resource names. """
        self.filedefs = []
        
        """ This dictionary maps GMAT resource names to workbook named Ranges 
        e.g., {'Sat.X': 'r20__km', 'Sat.Y': 'r21__km', 'Sat.Z': 'r22__km'}
        """
        self.var2gmat  = dict()
        
        """ This dictionary contains continuation variables and values as read-in from the
        'Mission_Parameters e.g. initial mass to orbit: '{Sat.InitialMass : 100, Sat.InitialMass : 110}.
        """
        self.varmultiplier = dict()
        
        """ This list contains the GMAT variables to be output in the ReportFile."""
        self.reportvars = list()

        """ Any user variables that need to be created on-the-fly 
        are created by the specialized class, e.g. {varname : GMAT Create Var|String}. """
        self.varset = {}  

