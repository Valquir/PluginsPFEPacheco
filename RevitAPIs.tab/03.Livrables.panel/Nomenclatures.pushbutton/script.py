# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "Nomenclatures"
__author__ = "Valquir Pacheco"
__min_revit_ver__ = 2019
__max_revit_ver = 2024
__doc__ = """"Test"""
__helpurl__ = "www.google.com"

# IMPORTS -----------------------------------------------------------------------------------------------
# .Net Imports
import clr
clr.AddReference("System")
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('Microsoft.Office.Interop.Excel')
from System.Collections.Generic import List

# Regular
import os, sys
from os import path, mkdir
from collections import OrderedDict
import ctypes
import System
# print("User Current Version:-", sys.version)
sys.path.append(r'C:\Users\vpacheco\AppData\Roaming\pyRevit-Master\site-packages')
sys.path.append(r'C:\Users\vpacheco\AppData\Roaming\pyRevit-Master\pyrevitlib')

# pyRevit + AutoDesk
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Structure
import pyrevit
from pyrevit import forms, revit, UI, script
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)

# Custom imports


# FUNCTIONS -----------------------------------------------------------------------------------------------

def exampleeee():
    """Creates a shared parameter file if it does not exist."""

    return 0

# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)
# if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit",0)

random_rebar = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().FirstElement()

# CODE  ----------------------------------------------------------------------------------------------------------

# MANUTENTION
# Nomenclature des aciers de manutentions pour les 3 éléments
# Récapitulatif des aciers de manutentions pour les 3 éléments

# MANCHON
# Nomenclature des manchons de l'élément X - cage XXXX
# Récapitulatif des manchons de l'élément X - cage XXXX (AVEC IMAGE)
# Récapitulatif des manchons de l'élément X - cage XXXX (TABLEAU)

# LONG ET TRANS
# Nomenclature des aciers de l'élément X
# Récapitulatif des aciers de l'élément X

# try:
tt = Transaction(doc, "Creer nomenclature")
tt.Start()

# Define the category (Rebar)
category = Category.GetCategory(doc, BuiltInCategory.OST_Rebar)
# #
# # Create the schedule view
schedule = ViewSchedule.CreateSchedule(doc, category.Id)
#
# # Define the schedule name
nom_feuille = "Rebar Quantity Schedule01"
schedule.Name = nom_feuille

# Add fields to the schedule
# FieldExample = pyrevit.revit.db.SchedulableField.GetField(doc, category.Id, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
# print(FieldExample)
# definitionNomenclature = ScheduleDefinition.AddField(FieldExample)
# print(definitionNomenclature)
definitionNomenclature = schedule.AddField(ScheduleFieldType.Instance, ElementId(-1010106))
print(definitionNomenclature)



# # Save the schedule
# schedule.Save()

tt.Commit()

ctypes.windll.user32.MessageBoxW(0, "Fonction Nomenclature terminé !", "Plugin Revit", 0)


# except (KeyError, AttributeError, EnvironmentError):
#     ctypes.windll.user32.MessageBoxW(0, "Sortir inattendu du logiciel, Plugin arrêté", "Plugin Revit", 0)

# except Exception as e:
#     print("Un erreur a etait produit:")
#     print(e)
#
#     if str(e) == "Name must be unique.":
#         ctypes.windll.user32.MessageBoxW(0, "Il existe des feuilles duplique avec nom : "+nom_feuille, "Plugin", 0)
#
#     else:
#         ctypes.windll.user32.MessageBoxW(0, "Action inattendu, plugin arrêté", "Plugin Revit", 0)

end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))

