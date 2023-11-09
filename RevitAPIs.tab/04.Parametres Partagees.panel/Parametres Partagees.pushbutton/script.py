# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "Parametres partagees"
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

def is_not_none(element):
    return element is not None

def create_shared_parameter_file():
    """
    Creates a shared parameter file if it does not exist.
    """
    shared_param_file_path = r'C:\Users\vpacheco\Desktop\PFE\9.Logiciels\Revit\Parametres.txt'
    # Create a new shared parameter file
    param_file = DB.SharedParameterFile.Create(shared_param_file_path)
    if param_file is None:
        print("Failed to create the shared parameter file.")
        return

    # Save the shared parameter file
    param_file.SaveAs(shared_param_file_path)

    print("Fichier de Parametree crée avec réussite.")

def create_new_shared_param(new_params, param_group):
    """Create shared_param in File."""

    sp_file = app.OpenSharedParameterFile()
    sp_groups = sp_file.Groups
    sp_groups_names = [g.Name for g in sp_groups]

    if param_group in sp_groups_names:
        new_group = sp_file.Groups.get_Item(param_group)
    else:
        new_group = sp_file.Groups.Create(param_group)

    for new_p_name in new_params:
        param_names = [p_def.Name for p_def in new_group.Definitions]

        if new_p_name not in param_names:
            print('Creating parameter [{}]'.format(new_p_name))

            if rvt_year >= 2023:
                option = ExternalDefinitionCreationOptions(new_p_name, SpecTypeId.String.Text)
            else:
                option = ExternalDefinitionCreationOptions(new_p_name, ParameterType.Text)

            new_group.Definitions.Create(option)

def add_shared_param_to_model():
    """Create shared_param in Shared Parameter File."""
    cats = doc.Settings.Categories
    cat_rebars = cats.get_Item(BuiltInCategory.OST_Rebar)
    cat_walls = cats.get_Item(BuiltInCategory.OST_Walls)

    cat_set01 = app.Create.NewCategorySet()
    cat_set01.Insert(cat_rebars)
    cat_set01.Insert(cat_walls)

    # cat_set02 = app.Create.NewCategorySet()
    # cat_set02.Insert(cat_walls)

    sp_file = app.OpenSharedParameterFile()
    sp_groups = sp_file.Groups

    print(sp_groups)

    t = Transaction(doc, 'Add Shared Parameter')
    t.Start()

    for d_group in sp_groups:
        for p_def in d_group.Definitions:
            if d_group.Name == "Rebar" :
                new_instance_binding = app.Create.NewInstanceBinding(cat_set01)
            else:
                new_instance_binding = app.Create.NewInstanceBinding(cat_set01)

            doc.ParameterBindings.Insert(p_def, new_instance_binding, BuiltInParameterGroup.PG_TEXT)

    t.Commit()

    ctypes.windll.user32.MessageBoxW(0, "Paramètres crée!", "Plugin Revit", 0)

def check_loaded_shared_params(list_parameters):
    """
    Check if required Wall and Rebar parameters ar loaded. If not, Load them.
    :param list_parameters:
    :return:
    """
    shared_params = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()
    sp_names = [p.Name for p in shared_params]
    missing_params = [p_name for p_name in list_parameters if p_name not in sp_names]

    # Load Parameters if necessary
    if missing_params:
        # Open SharedParameter File
        sp_file = app.OpenSharedParameterFile()
        if not sp_file:
            msg = 'Fichier de paramètres partagées non trouvée \nSi il vous plait verifier et essayer une autre fois'
            forms.alert(msg, exitscript=True)

        # Confirm add parameters
        confirmed = forms.alert("Il y a {} paramètres chargé sur le modele. \n {} \n \n "
                                "Voulez vous essayer de les mettre ?"
                                "".format(len(missing_params), '\n'.join(missing_params), sp_file.Filename),
                                ok=False, yes=True, no=True)
        print(confirmed)

        if confirmed:
            t = Transaction(doc, 'Ajouter Paramètre partagée')
            t.Start()

            # Create Category Set for Adding Parameter (Modify Parameters)
            cat_set = app.Create.NewCategorySet()
            cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls))
            cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar))

            # Create Instance/Type Bing
            # new_type_binding = app.Create.NewTypeBinding(cat_set)
            new_instance_binding = app.Create.NewInstanceBinding(cat_set)

            print("add missing params -------------------------------")
            # Add Missing Instance Sharded Parameters
            for d_group in sp_file.Groups:
                for p_def in d_group.Definitions:
                    if p_def.Name in missing_params:
                        print("passou params")
                        doc.ParameterBindings.Insert(p_def, new_instance_binding, BuiltInParameterGroup.PG_TEXT)

            print("check rebar -----------------------------------")
            for p in random_wall.Parameters:
                if p.Definition.Name in missing_params:
                    try:
                        print("passou rebar")
                        p.Definintion.SetAllowVaryBetweenGroups(doc, True)
                    except:
                        pass

            t.Commit()
            msg = 'Paramètres partagées cree avec réussite'
            forms.alert(msg, exitscript=False)

    else:
        msg = 'Paramètres partagées deja existente'
        forms.alert(msg, exitscript=False)

    return sp_names

# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)
# if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit",0)

random_wall = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().FirstElement()

# CODE  ----------------------------------------------------------------------------------------------------------

params_paroi_moulee = ["Coupe", "Element", "Cage"]
addedPARAMS = check_loaded_shared_params(params_paroi_moulee)
#
# boolAsk = forms.ask_for_one_item(['Oui', 'Non'], default='Oui',
#                                  prompt='Voulez vous ajouter de autre paramètre ?',
#                                  title='Paramètres partagées')
# if boolAsk == "Oui":

# Newparams = []
# addedPARAMS = []
# condition = True
# condition01 = True
#
# while condition01:
#     QuantParams = int(forms.ask_for_string(default='2', prompt='Combien ?',
#                                            title='Quantité paramètres partagées'))
#     if QuantParams <= 0:
#         msg0 = 'Donner au moins 1 paramètre'
#         forms.alert(msg0, exitscript=False)
#     elif QuantParams >= 10:
#         msg10 = 'Maximum 10 paramètres'
#         forms.alert(msg10, exitscript=False)
#     else:
#         condition01 = False
#
# while condition:
#     for z in range(QuantParams):
#         NomParamDummy = forms.ask_for_string(default='Example', prompt='Quel nom du paramètre  ' + str(z+1) + " ?",
#                                             title='Nom du paramètres partagées')
#         if NomParamDummy in addedPARAMS:
#             msg1 = 'Paramètre deja existent'
#             forms.alert(msg1, exitscript=False)
#             break
#
#         else:
#             Newparams.append(NomParamDummy)
#
#         if z == (QuantParams-1):
#             condition = False
#
# addedPARAMS.append(check_loaded_shared_params(Newparams))


# new_params = ["Cage", "Element", "Coupe"]
# create_new_shared_param(new_params, "Rebars")
# add_shared_param_to_model()

ctypes.windll.user32.MessageBoxW(0, "Fonction Paramètres partagées terminé !", "Plugin Revit", 0)
end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))

