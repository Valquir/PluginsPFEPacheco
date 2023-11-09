# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "01.Mur"
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
from Microsoft.Office.Interop import Excel
import csv

# pyRevit + AutoDesk
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Structure
import pyrevit
from pyrevit import forms, revit, UI, script
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)

# Custom imports
from Snippets._Mur import *
from Snippets._Excel import *




# FUNCTIONS -----------------------------------------------------------------------------------------------


# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)
# if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit",0)

# CODE  ----------------------------------------------------------------------------------------------------------

# PROCESSUS DE DEPART ************************************************************************
# Donnée d'entrée

try :
    questions = questions_depart()
    print(questions)

    # Verifier Famille Type de PM
    PMid = create_family_type_wall(questions['EP'])

    # Prendre Vecteur et points de depart pour créer Acier
    vectorDIR, vector, middle = select_structural_wall()

    # Créer chaque Element Mur en consideration des cages et division longueur
    wallsID = create_diaphragm_wall(vectorDIR, middle, questions['HauteursElements'], questions['LongueursDivisions'],
                                    PMid, questions['NC'])

    # Créer et Modifier Enrobage Mur
    create_and_apply_cover(questions['EN'], wallsID)

    # Sauvegarder informations sur fichier
    save_excel_file(questions)

    ctypes.windll.user32.MessageBoxW(0, "Mur crée avec réussite!", "Plugin Revit", 0)

except (KeyError, EnvironmentError):
    ctypes.windll.user32.MessageBoxW(0, "Sortir inattendu du logiciel, Plugin arrêté", "Plugin Revit", 0)

# except TypeError:
#     ctypes.windll.user32.MessageBoxW(0, "Donner un Fichier Excel s'il vous plait", "Plugin Revit", 0)
#
# except AttributeError:
#     ctypes.windll.user32.MessageBoxW(0, "Verifier si paramètres Coupe, Element et Cage sont presents", "Plugin Revit", 0)
#
#
# except Exception as e:
#     print("Un erreur a etait produit:")
#     print(e)
#
#     if str(e) == "The user aborted the pick operation.":
#         ctypes.windll.user32.MessageBoxW(0, "Operation Murs aborté", "Plugin Revit", 0)
#
#     else:
#         ctypes.windll.user32.MessageBoxW(0, "Action inattendu, plugin arrêté", "Plugin Revit", 0)

end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))

