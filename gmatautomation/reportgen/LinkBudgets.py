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
        18 May 2022: [CCH] Completed Link Budget Formulas.
        19 May 2022: [CCH] Corrected math issues related to elevation and azimuth.
        19 May 2022: [CCH] Integrated child extension class CDownlinkCapacity.


"""
import logging
import traceback
from pathlib import Path
from gmatautomation import CContactReports
#from ContactReports import CContactReports #For development only
from gmatautomation import CLinkReports
from gmatautomation import CGmatParticulars

class CLinkBudgets(CContactReports):
    def __init__(self,**args):
        super().__init__(**args)

    def formulaheadings(self):
        """ Trivial method to permit specialization of formulas used in Contact Report. """
        data = list()

        data.append(r'Slant.Range.(km)')
        data.append(r'Azimuth.(deg)')
        data.append(r'Elevation.(deg)')
        data.append(r'Free space Loss (db)')
        data.append(r'Antenna Mis-alignment')
        data.append(r'Clear Sky Loss (db)')
        data.append(r'Rain Loss 98% Avail. (db)')
        data.append(r'Total Losses w/ Rain (db)')
        data.append(r'Total Losses Clear-Sky (db)')
        data.append(r'ModCod')
        data.append(r'M')
        data.append(r'Es/No Min (db)')
        data.append(r'C/No Min (db)')
        data.append(r'Coded Bit Rate (bps)')
        data.append(r'Min Rcvr Signal Pwr (dbm)')
        data.append(r'Antenna Gains (db)')
        data.append(r'Required Xmit Power (dbm)')
        data.append(r'Required Xmit Power (mwatt)')
        data.append(r'Effective Message bit rate (bps)')

        hformulas = self.moreheadings()
        if len(hformulas) > 0:
            for h in hformulas:
                data.append(h)        
            
        return data

    def formulas(self, writerow, aoi):
        """ Specialization of formulas, using GMAT data compiled by CContactReports.
            Only formulas related to RF link budgets should go here.
            Further specialization is enabled by virtual methods moreformulas() and 
            moreheadings().
        """

        formrow = writerow + 1
        """ Excel Rows are 1-based. """
        data = list()  

        """ The cell format for the resultant values are returned as an array index:
            cellformats[0] = cell_heading
            cellformats[1] = cell_datetime
            cellformats[2] = cell_wrap
            cellformats[3] = cell_4plnum
            cellformats[4] = cell_2plnum
            cellformats[5] = cell_1digint
            cellformats[6] = cell_sep3digint

            DO NOT put a space before the equal sign in the formula.
        """
        data.append([r'= SQRT($E{0}^2 + $F{0}^2 + $G{0}^2)'.format(formrow), 3])
        """Formula for Slant Range (km)"""

        data.append([r'=DEGREES(ATAN2(E{0},F{0}))'.format(formrow), 3])
        """Formula for Azimuth (deg)"""

        data.append([r'=DEGREES(ATAN2(SQRT(E{0}^2+F{0}^2),D{0}))'.format(formrow), 3])
        """Formula for Elevation (deg)"""

        data.append([r'= -(32.4 + 20*LOG10($H{0}) + 20*LOG10(freq))'.format(formrow), 3])
        """Free space Loss (dB)"""

        data.append([r'= XLOOKUP($J{0}, Gnd_Antenna_Elevation_Range, Gnd_Antenna_Misalign_Loss, 0, 1)'.format(formrow), 4])
        """Antenna Misalignment"""

        data.append([r'= XLOOKUP("{0}", Rain_AOI, Total_Clear_Impaired, -0.1)'.format(aoi), 4])
        """Clear Sky Loss (dB)"""

        data.append([r'= XLOOKUP("{0}", Rain_AOI, Total_Rain_Impaired, -0.5)'.format(aoi), 4])
        """Rain Loss (dB)"""

        data.append([r'= $K{0} + $L{0} + $N{0} - Link_Margin'.format(formrow), 3])
        """Total Losses with Rain (dB)"""

        data.append([r'= $K{0} + $L{0} + $M{0} - Link_Margin'.format(formrow), 4])
        """Total Losses Clear-Sky (dB)"""

        data.append([r'=XLOOKUP($J{0}, VCM_elevation, VCM_ModCod, 2, -1)'.format(formrow), 5])
        """ModCod"""

        data.append([r'=XLOOKUP($J{0}, VCM_elevation, VCM_M, 2, 1)'.format(formrow), 5])
        """M"""

        data.append([r'= XLOOKUP($Q{0}, ModCod_ModCod, ModCod_Es_No, 6)'.format(formrow), 4])
        """Es/No Min (dB)"""

        data.append([r'= 10*LOG(B) + $S{0}'.format(formrow), 4])
        """C/No Min (dB)"""

        data.append([r'=Symbol_Rate * XLOOKUP($Q{0}, VCM_ModCod,VCM_BW_Eff, 1)'.format(formrow), 6])
        """Coded Bit Rate (BPS)"""

        data.append([r'= $T{0} + Total_Rcvr_Noise_P + XLOOKUP("Nairobi",Ground_Site, Noise_Figure_dB, 1)'.format(formrow), 4])
        """Min Rcvr Signal Pwr (dBm)"""

        data.append([r'= XLOOKUP("{0}", Ground_Site, X_band_Gnd_Ant_G) + SatAnt_Gain'.format(aoi), 4])
        """Antenna Gains"""

        data.append([r'=$V{0} - $W{0} - $O{0}'.format(formrow), 4])
        """Required Xmit Power (dBm)"""

        data.append([r'= 10^(0.1 * $X{0})'.format(formrow), 4])
        """Required Xmit Power (mW)"""

        data.append([r'=$U{0}*XLOOKUP($Q{0},ModCod_ModCod, Code_Rate__r,1)'.format(formrow), 6])
        """Effective Message bit rate (bps)"""

        formulas = self.moreformulas(formrow, aoi)
        """ Additional columns of custom Excel formulas. """

        try:
            if len(formulas) > 0:
                for f in formulas:
                    """ f should itself be a two-element array with the first a value, the second a cell format index, as above."""
                    if len(f) < 2:
                        print('moreformulas() return value is too short, derived class mismatch.')
                        raise RuntimeError("moreformulas() return value is too short, derived class mismatch.")
                    elif len(f) > 2:
                        print('moreformulas() return value is too long, derived class mismatch.')
                        raise RuntimeError("moreformulas() return value is too long, derived class mismatch.")
                    else:
                        data.append(f)

            return data

        except RuntimeError as e:
            lines = traceback.format_exc().splitlines()
            logging.error("Exception in CLinkBudgets moreformulas(): %s\n%s\n%s\n%s", e.__doc__, lines[0], lines[1], lines[-1])
            print('Exception in CLinkBudgets moreformulas(): ', e.__doc__, '\n', lines[0], '\n', lines[1],'\n', lines[-1])
 
    def moreheadings(self):
        """ further specialization of formula headings used in Link Budgets."""
        data = list()

        return data

    def moreformulas(self, formrow, aoi):
        """ further specialization of formula headings used in Link Budgets.
            Do not increment formrow.
        """
        data = list()

        return data

if __name__ == "__main__":
    __spec__ = None
    """ Necessary tweak to get Spyder IPython to execute this code. 
    See:
    https://stackoverflow.com/questions/45720153/
    python-multiprocessing-error-attributeerror-module-main-has-no-attribute
        """
    import getpass
    import platform
    from PyQt5.QtWidgets import(QApplication, QFileDialog)

    logging.basicConfig(
            filename='./LinkBudgets.log',
            level=logging.INFO,
            format='%(asctime)s %(filename)s \n %(message)s', 
            datefmt='%d%B%Y_%H:%M:%S')

    logging.info("!!!!!!!!!! Batch Report Format Execution Started !!!!!!!!!!")
    
    host_attr = platform.uname()
    logging.info('User Id: %s\nNetwork Node: %s\nSystem: %s, %s, \nProcessor: %s', \
                 getpass.getuser(), \
                 host_attr.node, \
                 host_attr.system, \
                 host_attr.version, \
                 host_attr.processor)

    try:
        gmat_paths = CGmatParticulars()
        o_path = gmat_paths.get_output_path()
        """ o_path is an instance of Path that locates the GMAT output directory. """

        qApp = QApplication([])

        fname = QFileDialog().getOpenFileName(None, 'Open Link Reports batch file', 
                        str(o_path),
                        filter='text files(*.batch)')

        logging.info('Link Report batch file is %s', fname[0])

        batchfile = Path(fname[0])

        lr_inst = CLinkReports()
        lr_inst.do_batch(batchfile)

        logging.info ('Link Reports completed.')
        print('Link Reports completed.\n')

        fname = QFileDialog().getOpenFileName(None, 'Open (SIGHT) Contact Locater Reports batch file.', 
                        str(o_path),
                        filter='Batch files(*.batch)')

        logging.info('Contact Locater Report batch file is %s', fname[0])

        sightfile = Path(fname[0])

        cr_inst = CLinkBudgets()
        cr_inst.setlinks(lr_inst.links)     
        cr_inst.do_batch(sightfile)

        logging.info ('Contact Reports Completed.')
        print('Contact Reports Completed.')

    except Exception as e:
        lines = traceback.format_exc().splitlines()
        logging.error('Exception %s caught at top level:\n%s\n%s\n%s', e.__doc__, lines[0], lines[1], lines[-1])
        print('Exception ', e.__doc__,' caught at top level: ', lines[0],'\n', lines[1], '\n', lines[-1])
        
    finally:
        qApp.quit()
                                