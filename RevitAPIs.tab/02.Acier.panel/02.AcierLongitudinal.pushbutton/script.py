# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "02.Acier Longitudinal"
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
from Snippets._Rebar import *
from Snippets._Excel import *

# FUNCTIONS -----------------------------------------------------------------------------------------------

def is_not_none(element):
    return element is not None

# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)
if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit", 0)

numbers = ['01', '02', '03', '04', '05', '06']

Rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType()
Murs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()

# CODE  ----------------------------------------------------------------------------------------------------------

try:

    # PROCESSUS ACIERS ************************************************************************
    wallsID = []
    DonneeQuestions = []


    # Lire EXCEL
    donne_excel = read_excel_file()
    for i in donne_excel["Depart"]:
        DonneeQuestions.append(filter(is_not_none, i))
    lignesLong = donne_excel["Longitudinale"]

    CoupeDesiree = DonneeQuestions[0][1]
    Enrobage = float(DonneeQuestions[3][1])
    AS = float(DonneeQuestions[1][1])
    FM = float(DonneeQuestions[2][1])
    del DonneeQuestions[4][0]
    Recouvrement = DonneeQuestions[4]

    # Prendre Vecteur et points de depart pour créer Acier
    vectorPER, vectorPAR, middlePoint, element = choisir_mur_acier()
    CageDesiree = element.LookupParameter("Cage").AsValueString()

    for Mur in Murs:
        if Mur.LookupParameter("Coupe").AsValueString() == CoupeDesiree:
            wallsID.append(Mur.Id)

    if wallsID is None:
        print("Coupe non existant !")

    # ACIER Long
    tl = Transaction(doc, "Acier Longitudinal")
    tl.Start()

    for bar in lignesLong:
        index = lignesLong.index(bar) + 1
        DonneeDepart = {}
        DonneeRenfort = {}
        AcierExistent = []
        NewRebarLongP = "No"
        NewRebarLongR = "No"

        DonneeDepart["Type"] = bar[0]
        DonneeDepart["Zone"] = bar[1]
        DonneeDepart["NomBarre"] = bar[2]
        DonneeDepart["QAcierLong"] = bar[3]
        DonneeDepart["DiamLong"] = bar[4]

        for rebar in Rebars:
            if rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).AsValueString() == DonneeDepart["NomBarre"]\
               and rebar.LookupParameter("Coupe").AsValueString() == CoupeDesiree:
                AcierExistent.append(rebar)

        if len(AcierExistent) >= 1:
            boolAsk = forms.ask_for_one_item(['Oui', 'Non'], default='Oui', prompt='Le acier ' +
                                             DonneeDepart["NomBarre"] +
                                             ' existe deja, voulez vous refaire ?', title='Acier')
            if boolAsk == "Oui":
                for i in AcierExistent:
                    doc.Delete(i.Id)
            else:
                continue

        if DonneeDepart["Type"] == "Principal":
            NewRebarLongP = rebar_longitudinal_prin(wallsID, DonneeDepart, vectorPER, vectorPAR, Enrobage, Recouvrement)
            if NewRebarLongP == []:
                print("Element " + str(index) + ", avec nom : " + DonneeDepart["NomBarre"] + ", non crée")
            else:
                print("Element " + str(index) + ", avec nom : " + DonneeDepart["NomBarre"] + ", crée !")

        else:
            DonneeRenfort["Coupe"] = CoupeDesiree
            DonneeRenfort["Cage"] = "0"+str(int(bar[5]))
            DonneeRenfort["Element"] = "0"+str(int(bar[6]))
            DonneeRenfort["AraseSUP"] = float(bar[7])
            DonneeRenfort["AraseINF"] = float(bar[8])
            DonneeRenfort["Groupe"] = "0"+str(int(bar[9]))
            DonneeRenfort["Lit"] = "0"+str(int(bar[10]))

            if (DonneeRenfort["AraseSUP"] - DonneeRenfort["AraseINF"] > AS - FM) or \
               (DonneeRenfort["AraseSUP"] > AS) or (DonneeRenfort["AraseINF"] < FM):
                ctypes.windll.user32.MessageBoxW(0, "Donne sur Arase SUP/INF incorrecte, verifier ! ", "Plugin Revit", 0)

            elif (DonneeRenfort["Cage"] or DonneeRenfort["Element"] or DonneeRenfort["Groupe"] or DonneeRenfort["Lit"]) \
                  not in numbers:
                ctypes.windll.user32.MessageBoxW(0, "Verifier, Cage, Element, Groupe et Lit ! ", "Plugin Revit", 0)

            else:
                NewRebarLongR = rebar_longitudinal_renfort(DonneeDepart, DonneeRenfort, vectorPER, vectorPAR, Enrobage,
                                                           AS, FM)
                if NewRebarLongR == []:
                    print("Element Element " + str(index) + ", avec nom : " + DonneeDepart["NomBarre"] + ", non crée")
                else:
                    print("Element " + str(index) + ", avec nom : " + DonneeDepart["NomBarre"] + ", crée !")

    tl.Commit()

    ctypes.windll.user32.MessageBoxW(0, "Aciers crée avec réussite !", "Plugin Revit", 0)

except (KeyError, AttributeError, EnvironmentError):
    ctypes.windll.user32.MessageBoxW(0, "Sortir inattendu du logiciel, Plugin arrêté", "Plugin Revit", 0)

except TypeError:
    ctypes.windll.user32.MessageBoxW(0, "Donner un Fichier Excel s'il vous plait", "Plugin Revit", 0)

except IndexError:
    ctypes.windll.user32.MessageBoxW(0, "Verifier s'il existe des elements dupliquée dans le MN.", "Plugin Revit", 0)
    # Créer fonction pour analyser s'il existe ces elements dupliquées

except Exception as e:
    print("Un erreur a etait produit:")
    print(e)

    if str(e) == "The user aborted the pick operation.":
        ctypes.windll.user32.MessageBoxW(0, "Operation Murs aborté", "Plugin Revit", 0)

    else:
        ctypes.windll.user32.MessageBoxW(0, "Action inattendu, plugin arrêté", "Plugin Revit", 0)

end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))

