# -*- coding: utf-8 -*-

from datetime import datetime
start_time = datetime.now()

__title__ = "00.Obsolete functions"  # Name of button
__author__ = "Valquir Pacheco"
__min_revit_ver__ = 2019
__max_revit_ver = 2024
__doc__ = """"Test"""
# __helpurl__ = "www.google.com"

# IMPORTS -----------------------------------------------------------------------------------------------
# .Net Imports
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# Regular
import os, sys
from os import path, mkdir
from collections import OrderedDict
import ctypes
import System
# print("User Current Version:-", sys.version)

# pyRevit + AutoDesk
sys.path.append(r'C:\Users\Valqu\AppData\Roaming\pyRevit-Master\pyrevitlib')
sys.path.append(r'C:\Users\Valqu\AppData\Roaming\pyRevit-Master\site-packages')
from Autodesk.Revit.DB import *
from pyrevit import forms, revit
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,Separator, Button, CheckBox)

# Custom imports
from Snippets._ConvertUnits import convert_internal_to_m
clr.AddReference('Microsoft.Office.Interop.Excel')
from Microsoft.Office.Interop import Excel
import csv

# FUNCTIONS -----------------------------------------------------------------------------------------------
def select_wall_lign():
  """Requests the user to select a line, created from exterior to interior in Revit."""
  faire_loop = True
  while faire_loop:

      with forms.WarningBar(title='Choisir ligne crée à la droit de la face extérieur qui est en haut :'):
        element = revit.pick_element()

      if element.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE):
        faire_loop = False
      else:
        ctypes.windll.user32.MessageBoxW(0, "S'il vous plait choisir Ligne", "Plugin Revit", 0)
  return element.GeometryCurve

def select_structural_wall_ancienne():
  """Requests the user to select a structural wall in Revit."""
  faire_loop = True
  while faire_loop:

      with forms.WarningBar(title='Choisir paroi moulée:'):
          element = revit.pick_element()

      if element.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT):
          if element.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger() == 1:
              faire_loop = False
      else:
          ctypes.windll.user32.MessageBoxW(0, "S'il vous plait choisir Mur structural", "Plugin Revit", 0)
  return element

def get_parameters_wall(wallrevit):
  """Collect basic information/parameters about Wall."""
  colonnes = ['ID', 'Categorie', 'Famille et type', 'Nivel inferior', 'Nivel superior',
              'Elevation base', 'Hauteur totale', 'Longueur', 'Largeur', 'Cote terre']

  dictionarydummy = OrderedDict()
  for i in colonnes:
      dictionarydummy[i] = None

  dictionarydummy['ID'] = wallrevit.Id.ToString()
  dictionarydummy['Categorie'] = wallrevit.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString()
  dictionarydummy['Famille et type'] = wallrevit.get_Parameter(
      BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
  dictionarydummy['Nivel inferior'] = wallrevit.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsValueString()
  dictionarydummy['Nivel superior'] = wallrevit.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsValueString()
  dictionarydummy['Elevation base'] = convert_internal_to_m(
      doc.GetElement(wallrevit.LevelId).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())
  dictionarydummy['Hauteur totale'] = convert_internal_to_m(
      wallrevit.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble())
  dictionarydummy['Longueur'] = convert_internal_to_m(
      wallrevit.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble())
  dictionarydummy['Largeur'] = convert_internal_to_m(
      wallrevit.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble())
  dictionarydummy['Cote terre'] = "dummy"

  return dictionarydummy

def get_corners(Mur, vecteur):
    """Create all corners."""
    # 1st letter (Top or Lower [T/L]), 2nd (Left or Right [L/R]), 3rd (Front or Back [F/B])
    # ctypes.windll.user32.MessageBoxW(0, "Processus crée curvature et vecteur", "Plugin Revit", 0)
    view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)
    # if default_3d_view is None: ctypes.windll.user32.MessageBoxW(0, "S'il vous plait créer 1 View 3D !", "Plugin Revit",0)

    MurLargueur = Mur.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    bounding_box = Mur.get_BoundingBox(default_3d_view)
    TRF_point = bounding_box.Max
    LLB_point = bounding_box.Min

    LRF_point = XYZ(TRF_point.X, TRF_point.Y, LLB_point.Z)
    TLB_point = XYZ(LLB_point.X, LLB_point.Y, TRF_point.Z)

    LRB_point = XYZ(LRF_point.X - vecteur.X * MurLargueur, LRF_point.Y - vecteur.Y * MurLargueur, LLB_point.Z)
    LLF_point = XYZ(LLB_point.X + vecteur.X * MurLargueur, LLB_point.Y + vecteur.Y * MurLargueur, LLB_point.Z)
    TRB_point = XYZ(LRB_point.X, LRB_point.Y, TRF_point.Z)
    TLF_point = XYZ(LLF_point.X, LLF_point.Y, TRF_point.Z)

    dict_points = {}
    dict_points["LRF_point"] = LRF_point
    dict_points["LRB_point"] = LRB_point
    dict_points["LLF_point"] = LLF_point
    dict_points["TLB_point"] = TLB_point
    dict_points["TRB_point"] = TRB_point
    dict_points["TLF_point"] = TLF_point

    center_top = XYZ((TRF_point.X + LLB_point.X)/2, (TRF_point.Y+LLB_point.Y)/2, TRF_point.Z)

    return center_top

def rebar_transversal_principal_obs(DonneeTrans, vectorPER, VN, Enrobage, AS, wallsID):
    """All steps for principal transversal rebars."""
    EN = Enrobage * 3.28084
    MurTrans = []
    AcierTrans = []


    # Chercher le TYPE de diamètre
    ElementDiam = 0
    for z in DiamElements:
        if str(int(DonneeTrans["DiamTrans"])) + " mm" == z.get_Parameter(
                BuiltInParameter.REBAR_BAR_DIAMETER).AsValueString():
            ElementDiam = z
            break

    # Chercher Forme
    ElementShape = 0
    for y in FormeElements:
        if y.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() == "6-20":
            ElementShape = y
            break


    # Prendre que les PM de la cage interessé
    MurDummy = 0
    MurTransDebut = 0
    for i in wallsID:
        MurDummy = doc.GetElement(i)
        # Donnée de base de la cage d'intérêt et le premier element de la cage
        if MurDummy.LookupParameter("Cage").AsValueString() == DonneeTrans["Cage"]:
            MurTrans.append(MurDummy)
            if MurDummy.LookupParameter("Element").AsValueString() == '01':
                MurTransDebut = MurDummy


    # Créer les armatures transversales principales pour chaque
    for MurTran in MurTrans:
        NewRebarTR = 0
        MurTLargueur = MurTran.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
        MurTLongueur = MurTran.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()


        # Prendre la hauteur selon unités de Revit
        bounding_box = MurTran.get_BoundingBox(default_3d_view)
        TRF_pointR = bounding_box.Max
        LLB_pointR = bounding_box.Min
        mid_pointR = XYZ((TRF_pointR.X + LLB_pointR.X) / 2, (TRF_pointR.Y + LLB_pointR.Y) / 2, TRF_pointR.Z)
        reference = DonneeTrans["AraseSup"] / AS
        z = mid_pointR.Z * reference - (DonneeTrans["QAcierTrans"] * DonneeTrans["Espacement"]) * 0.328084

        DisplacementEnX = EN - MurTLongueur / 2
        DisplacementEnY = EN - MurTLargueur / 2

        point_start = XYZ(mid_pointR.X + (VN.X * DisplacementEnY) + (vectorPER.X * DisplacementEnX),
                          mid_pointR.Y + (VN.Y * DisplacementEnY) + (vectorPER.Y * DisplacementEnX), z)

        NewRebarTR = Structure.Rebar.CreateFromRebarShape(doc, ElementShape, ElementDiam, MurTran, point_start,
                                                          vectorPER, VN)

        # Paramètres de configuration
        NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)                          # type conf.
        NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(DonneeTrans["QAcierTrans"])# Quantité
        NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(DonneeTrans["Espacement"])  # Espacement
        NewRebarTR.GetShapeDrivenAccessor().BarsOnNormalSide = True

        # NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_HOOK_START_TYPE).Set(DonneeTrans["Crochet"].IntegerValue)
        # NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_HOOK_END_TYPE).Set(DonneeTrans["Crochet"].IntegerValue)

        # Paramètres d'information
        NewRebarTR.LookupParameter("Cage").Set(MurTran.LookupParameter("Cage").AsValueString())
        NewRebarTR.LookupParameter("Element").Set(MurTran.LookupParameter("Element").AsValueString())
        NewRebarTR.LookupParameter("Coupe").Set(MurTran.LookupParameter("Coupe").AsValueString())
        NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).Set(DonneeTrans["NomBarre"])
        NewRebarTR.get_Parameter(BuiltInParameter.NUMBER_PARTITION_PARAM).Set("Transversale principale")

        AcierTrans.append(NewRebarTR)

    return AcierTrans

#EXAMPLE
random_room = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).FirstElemenet()
if not random_room :
    msg = 'There are no rooms in the project. Please try again '

# Function check parameter
def check_loaded_shared_params_obs(list_parameters):
    """
    Check if required Room parameters ar loaded. If not, Load them.
    :param list_parameters:
    :return:
    """
    shared_params = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()
    sp_names = [p.Name for p in shared_params]
    missing_params = [p.Name for p_name in list_parameters if p_name not in sp_names]

    #Load Parameters if necessary
    if missing_params :
        #Open SharedParameter File
        sp_file = app.OpenSharedParameterFile()
        if not sp_file:
            msg = 'Could not find Shared Parameter File \nPlease verify and try again'
            forms.alert(msg, exitscript=True)

        # Confirm add parameters
        confirmed = forms.alert("There are {} missing parameters for the sript. \n {} \n \n "
                                "Would you like to try to load them?"
                                "".format(len(missing_params), '\n'.join(missing_params), sp_file.Filename),
                                ok=False, yes=True, no=True)

        if confirmed:
            t = Transaction(doc, 'Add SharedParameters')
            t.Start()

            # Create Category Set for Adding Parameter (Modify Parameters)
            cat_set = app.Create.NewCategorySet()
            cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms))

            # Create Instance/Type Bing
            # new_type_binding = app.Create.NewTypeBinding(cat_set)
            new_instance_binding = app.Create.NewInstanceBinding(cat_set)

            # Add Missing Instance Sharded Parameters
            for d_group in sp_file.Groups:
                for p_def in d_group.Definitions:
                    if p_def.Name in missing_params:
                        doc.ParameterBindings.Insert(p_def, new_instance_binding, BuiltInParameterGroup.PG_TEXT)

            for p in random_room.Parameters:
                if p.Definition.Name in missing_params:
                    try: p.Definintion.SetAllowVaryBetweenGroups(doc, True)
                    except: pass
            t.Commit()


# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)

Ordre_alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
               'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

dictionaryWall = OrderedDict()

# CODE  -------------------------------------------------------------------------------------------------------

# wall = select_structural_wall()
# dictionaryWall = get_parameters_wall(wall)
# print(dictionaryWall)

# ref = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior) # Mazri's BIM Diary
# face = wall.GetGeometryObjectFromReference(ref[0])
# # print(face)


# ACIERS ------------------------------------
# Nao funciona de jeito nenhum
# NewRebar = Structure.Rebar.CreateFromCurvesAndShape(doc, FormeElements[0], DiamElements[8], HookElements[5],
#                                                     HookElements[5], Mur, vectorDIR, ligne, HookRight, HookRight)

# Parece limitado aos tamanhos
# Aciers = Structure.Rebar.CreateFromRebarShape(doc, FormeElements[1], DiamElements[1],Mur, start, vectorDIR, vectorDIR)

