# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "05. Auxiliare"
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
# if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit",0)

# create_shared_parameter_file()
# new_params = ["Cage", "Element", "Coupe"]
# create_new_shared_param(new_params, "Rebars")
# add_shared_param_to_model()

numbers = ['01', '02', '03', '04', '05', '06']

Rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType()
Murs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()

# CODE  ----------------------------------------------------------------------------------------------------------

msg = 'Work in progress !'
forms.alert(msg, exitscript=True)

# PROCESSUS ACIERS ************************************************************************
# SéPARER EN TYPE D'ACIER, MANCHON (FOR AVEC LES DALLES), AUXILIAIRE (CHAQUE CAGE), LONGITUDINAL ET TRANSVERSAL
# wallsID = []
# DonneeQuestions = []
#
# # try:
# # Lire EXCEL
# donne_excel = read_excel_file()
# for i in donne_excel["Depart"]:
#     DonneeQuestions.append(filter(is_not_none, i))
# lignesManchons = donne_excel["Auxiliar"]
#
# CoupeDesiree = DonneeQuestions[0][1]
# Enrobage = float(DonneeQuestions[3][1])
# AS = float(DonneeQuestions[1][1])
# FM = float(DonneeQuestions[2][1])
# del DonneeQuestions[4][0]
# Recouvrement = DonneeQuestions[4]
#
# # Prendre Vecteur et points de depart pour créer Acier
# vectorPER, vectorPAR, middlePoint, element = choisir_mur_acier()
#
# for Mur in Murs:
#     if Mur.LookupParameter("Coupe").AsValueString() == CoupeDesiree:
#         wallsID.append(Mur.Id)
#
# if wallsID is None:
#     print("Coupe non existant !")
#
# # Transversal
# for bar in lignesTrans:
#     DonneeTrans = {}
#     AcierTExistent = []
#     NewRebarLong = []
#     NewRebarTrans = "No"
#     NewRebarLongR = "No"
#
#     DonneeTrans["Cage"] = "0"+str(int(bar[0]))
#     DonneeTrans["Zonage"] = "0"+str(int(bar[1]))
#     DonneeTrans["NomBarre"] = bar[2]
#     DonneeTrans["QAcierTrans"] = int(bar[4])
#     DonneeTrans["Espacement"] = float(bar[5]) * 3.28084
#     DonneeTrans["AraseSup"] = float(bar[3])
#     DonneeTrans["DiamTrans"] = bar[6]
#     DonneeTrans["Debut_long"] = bar[7]
#     DonneeTrans["Fin_long"] = float(bar[8])
#     DonneeTrans["Fin_long"] = float(bar[9])
#     DonneeTrans["Crochet"] = verifier_creer_crochet_acier(int(bar[10]))
#     DonneeTrans["CrochetLong"] = float(bar[5]) * 3.28084
#
#     print(DonneeTrans["AraseSup"] * 3.28084 - DonneeTrans["QAcierTrans"] * DonneeTrans["Espacement"])
#     print(FM * 3.28084)
#     if DonneeTrans["AraseSup"] * 3.28084 - DonneeTrans["QAcierTrans"] * DonneeTrans["Espacement"] < FM * 3.28084:
#         print("SVP verifier quantité d'acier, espacement pour être entre Arase SUP et la fiche mécanique.")
#
#     elif DonneeTrans["AraseSup"] > AS:
#         ctypes.windll.user32.MessageBoxW(0, "SVP entre debut acier et Arase superieur", "Plugin Revit", 0)
#
#     else:
#         tt = Transaction(doc, "Acier Transversal")
#         tt.Start()
#
#         for rebar in Rebars:
#             if rebar.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).AsValueString() == DonneeTrans["NomBarre"]:
#                 AcierTExistent.append(rebar)
#
#         if len(AcierTExistent) >= 1:
#             boolAsk = forms.ask_for_one_item(['Oui', 'Non'], default='Oui', prompt='Le acier ' + DonneeTrans["NomBarre"] +
#                                              ' existe deja, voulez vous refaire ?', title='Acier')
#             if boolAsk == "Oui":
#                 for i in AcierTExistent:
#                     doc.Delete(i.Id)
#             else:
#                 continue
#
#         if DonneeTrans["Zonage"] == "01":
#             NewRebarLong = rebar_transversal_principal(DonneeTrans, vectorPER, vectorPAR, Enrobage, AS, wallsID)
#
#         tt.Commit()
#
#         if not NewRebarLong:
#             print("Acier Nom " + DonneeTrans["NomBarre"] + " non crée, il faut verifier !")
#         else:
#             print("Acier Nom " + DonneeTrans["NomBarre"] + " crée avec réussite")
#
# ctypes.windll.user32.MessageBoxW(0, "Fonction créer aciers transversaux terminé !", "Plugin Revit", 0)
#
# except (KeyError, AttributeError, EnvironmentError):
#     ctypes.windll.user32.MessageBoxW(0, "Sortir inattendu du logiciel, Plugin arrêté", "Plugin Revit", 0)
#
# except TypeError:
#     ctypes.windll.user32.MessageBoxW(0, "Donner un Fichier Excel s'il vous plait", "Plugin Revit", 0)
#
# except IndexError:
#     ctypes.windll.user32.MessageBoxW(0, "Verifier s'il existe des elements dupliquée dans le MN.", "Plugin Revit", 0)
#     # Créer fonction pour analyser s'il existe ces elements dupliquées
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

