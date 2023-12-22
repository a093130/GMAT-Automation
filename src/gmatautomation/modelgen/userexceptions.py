#! Python
# -*- coding: utf-8 -*-

#Tailored Header and Implementation file strategic leading comment as follows:
#Tailored for doxygen.
""" 
	@file userexceptions.py
	@brief custom exception classes
	
	@copyright 2023 AstroForge-Incorporated
	@author  <name>, [CCH]
	@authoremail colin@astroforge.io, colinhelms@outlook.com

	@version 0.1a0

	@details <expanded description>.
	
	@remark Change History
		05 Dec 2023: [CCH] File created

	@bug [<initials>] <backlog item>

"""
import logging

"""
@defgroup Globals
@brief Python global constants are persistent across calls to functions.  The ALLCAPS case
is reserved for identification of globals.
"""
# @{
""" Insert global definitions in group. """

# @} End of Globals


"""	@defgroup Classes
	@brief Class definitions are grouped for easy reference in documentation.
"""
# @{
""" Insert Class defs in group. """

#Tailored class strategic comment as follows:
"""
	@brief description
	@details detailed description
"""	
class Ultima(Exception):
    """
        @brief useful user exception
        @details meant as an enclosing exception to ensure that module is exited and cleanup occurs.
        For instance when an exception is caught by a function down in the stack and it is intended to 
        unroll the stack this exception may be re-raised.
    """
    def __init__(self, source='filename', message='User Exception Utltima raised.', **args):
        super().__init__(**args)
        self.source = source
        self.message = message
        logging.warning(self.message)

# @} End of Class Group