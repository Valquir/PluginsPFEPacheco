# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *

# VARIABLES
app = __revit__.Application

# Functions
def convert_internal_to_m(length):
    """Function to convert internal units to meters.
    :param length : Length in internal Revit Units
    :return : Length in Meters, rounded to 2nd digit"""

    rvt_year = int(app.VersionNumber)

    # REVIT < 2022
    if rvt_year < 2022:
        return UnitUtils.Convert(length, DisplayUnitType.DUT_DECIMAL_FEET,DisplayUnitType.DUT_METERS)

    # REVIT >= 2022
    else:
        return UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Meters)

def convert_m_to_internal(length):
    """Function to convert internal units to meters.
    :param length : Length in internal Revit Units
    :return : Length in Meters, rounded to 2nd digit"""

    rvt_year = int(app.VersionNumber)

    # REVIT < 2022
    if rvt_year < 2022:
        return UnitUtils.Convert(length, DisplayUnitType.DUT_METERS, DisplayUnitType.DUT_DECIMAL_FEET)

    # REVIT >= 2022
    else:
        return UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Feet)
