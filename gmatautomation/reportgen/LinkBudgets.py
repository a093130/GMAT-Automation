#! Python
# -*- coding: utf-8 -*-
""" 
	@file LinkBudgets.py
	@brief module container for class definition CLinkBudgets.
	
	@copyright: Copyright (C) 2022 Freelance Rocket Science, All rights reserved.
	@author  Colin Helms, colinhelms@outlook.com, [CCH]

	@details Class def CLinkBudgets extends the class CContactReports to insert useful Exel
    formulas for calculating Link Budgets.  The parent class provide virtual methods
    formulaheadings(self, row) and formulas(self, row) to be extended herein.
	
	@remark Change History
		22 April 2022: [CCH] File created, GitHub repository GMAT-Automation.
        Tue Apr 26 2022 [CCH] Version 0.2a1, Buildable package, locally deployable.

		
	@bug https://github.com/a093130/GMAT-Automation/issues

"""
from gmatautomation import CContactReports

class CLinkBudgets(CContactReports):
    def __init__(self,**args):
        super().__init__(**args)

    def formulaheadings(self):
        """ Trivial method to permit specialization of formulas used in Contact Report. """
        data = list()

        data.append('Slant.Range.(km)')
        data.append('Azimuth.(deg)')
        data.append('Elevation.(deg)')

        return data

    def formulas(self, writerow):
        """ Trivial method to permit specialization of formulas used in Contact Report. """
        formrow = writerow +1
        """ Excel Rows are 1-based. """
        data = list()     

        data.append('=SQRT(E{0}^2+F{0}^2+G{0}^2)'.format(formrow))
        """Formula for Slant Range (km)"""
        data.append('=DEGREES(ATAN(F{0}/E{0}))'.format(formrow))
        """Formula for Azimuth (deg)"""
        data.append('=DEGREES(ATAN(G{0}/(SQRT(E{0}^2+F{0}^2))))'.format(formrow))
        """Formula for Elevation (deg)"""

        return data
                                