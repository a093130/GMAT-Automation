# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 09:49:45 2018

@author: colinhelms@outlook.com

@Description:
    This module encapsulates the Points of Variation dictionary for GMAT model
    generation.  It is created by "fromconfigsheet.py" and used by "modelgen.py.
    
    The implemented points of variation in the model file are listed here.
    The list form is [heading: (top model name) (sub-model name).(resource)].
    
    Points of variation:
        Dry mass: Spacecraft EOTV.DryMass
        Starting Epoch: Spacecraft EOTV.Epoch
            This is a list of epoch values to be executed.
        Inclination: Spacecraft EOTV.INC
            This is a list of inclination values  to be executed.
        Costate: The constraint value for final SMA and Inclination change
            This is not actually a model parameter but is used to select a JSON
            control file.
        Max Thrust Power: ElectricThruster HET1.MaximumUsablePower
        Min Thrust Power: ElectricThruster HET1.MinimumUsablePower
        Efficiency: ElectricThruster HET1.FixedEfficiency
        Isp: ElectricThruster HET1.Isp
        Thrust: ElectricThruster HET1.ConstantThrust
        Available Power: SolarPowerSystem EOTVSolarArrays.InitialMaxPower
        Propellant: ElectricTank RAPTank1.FuelMass
        Output: ReportFile1 ReportFile1.Filename
        Viewpoint: Orbit View DefaultOrbitView.ViewPointVector

    Notes:
       1. In order to cover the range of eclipse conditions the EOTV Epoch is
          typically varied for the four seasons:
                20 Mar 2020 03:49 UTC
                20 Jun 2020 21:43 UTC
                22 Sep 2020 13:30 UTC
                21 Dec 2020 10:02 UTC
          The epoch dates are specified in the configsheet workbook named as
          Epoch_1, Epoch_2, Epoch_3, Epoch_4.
       2. The Viewpoint is superfluous in most cases, since model execution
          is intended to be in batch mode for this system. However, if graphic
          output is desired, the OrbitView viewpoint is algorithmically
          varied associated with the Starting Epoch as follows:
                20 Mar 2020 03:49 UTC, (80000, 0, 20000)
                20 Jun 2020 21:43 UTC, (0, 80000, 20000)
                22 Sep 2020 13:30 UTC, (0, -80000, 20000)
                21 Dec 2020 10:02 UTC, (-80000, 0, 20000)
       3. The ReportFile.Filename and the Model name are generated by concatenating
          the configuration, the Starting Epoch and the Inclination
@Change Log:
    15 Nov 2018, baseline set of model resources supporting paper AIAA-2018-4718
    09 Feb 2019, Integration, Fix ReportFile parameter, should be "ReportFile1".
    10 Apr 2019, Use GMAT user variable Costate to store Costates.  Removed kludge
        using 'EOTV.ID' to store payload mass.  Associate 'EOTV.ID' with 'SID'.
        Use GMAT user variable 'PL_MASS' to store 'Payload'
    
"""
def getvarnames():
    """ This dictionary maps GMAT resource names to the workbook name, which is the
    table heading in row 1 of the worksheet named 'GMAT'. 
    """
    var2gmat = dict([
                       ('ReportFile1.Filename', 'Configuration'),
                       ('EOTV.DryMass','Dry Mass'),
                       ('PL_MASS', 'Payload'),
                       ('HET1.MaximumUsablePower', 'Max Thrust Power'),
                       ('HET1.MinimumUsablePower', 'Min Thrust Power'),
                       ('HET1.FixedEfficiency', 'Efficiency'),
                       ('HET1.Isp', 'Isp'),
                       ('HET1.ConstantThrust', 'Thrust'),
                       ('EOTVSolarArrays.InitialMaxPower', 'Available Power'),
                       ('RAPTank1.FuelMass', 'Propellant')
                       ])
    
    return var2gmat.copy()

def getrecursives():
    """ This dictionary  which are named ranges of the worksheet named 'Mission_Paramsarameter. """
    varmultiplier = dict([
                        ('EOTV.Epoch', 'Starting Epoch'),
                        ('EOTV.INC', 'Inclinations'),
                        ('COSTATE', 'Inclinations'),
                        ('SMA_INIT', 'Altitude'),
                        ('SMA_END', 'Altitude')
                        ])

    return varmultiplier.copy()
