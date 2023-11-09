# -*- coding: utf-8 -*-

__title__ = "_Rebar"  # Name of button
__author__ = "Valquir Pacheco"
__min_revit_ver__ = 2019
__max_revit_ver = 2024
__doc__ = """"Test"""

# IMPORTS -----------------------------------------------------------------------------------------------
# .Net Imports
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# Regular
import os, sys
import ctypes

# pyRevit + AutoDesk
import pyrevit
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Structure
from pyrevit import forms, revit, UI


# FUNCTIONS -----------------------------------------------------------------------------------------------

def catch_rebar_details(RebarsDNames, RebarsSNamesList, RebarHookNames):
    """Catch group of Rebar diameters, shapes and hooks."""
    # INNOVATION 01 : demander utilisateur choisir diamètres selon liste, meme chose pour formats, mais avec image
    # ctypes.windll.user32.MessageBoxW(0, "Processus informations acier !", "Plugin Revit", 0)
    RebarShapes = []
    RebarDiam = []
    RebarHook = []
    RebarsDiamChoisi = []
    RebarsShapeChoisi = []
    RebarsHooksChoisi = []

    for i in Rebars:
        if i.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString() == "Barre d'armature":  # Diam
            RebarDiam.append(i)
            if i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() in RebarsDNames:
                RebarsDiamChoisi.append(i)
                # print("Passou : " + i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())

        elif i.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString() == "Crochet d'armature":  # Hook
            if i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() in RebarHookNames:
                RebarsHooksChoisi.append(i)

        else:  # Format
            RebarShapes.append(i)
            if i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() in RebarsSNamesList:
                RebarsShapeChoisi.append(i)
                # print("Passou : " + i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())

    # print("Les différents formes --------------")
    # for i in RebarShapes:
    #     print(i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())
    # print("Les différents diamètre --------------")
    # for i in RebarDiam:
    #     print(i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())
    # print("Les différents ganchos --------------")
    # for i in RebarHook:
    #     print(i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())

    # ctypes.windll.user32.MessageBoxW(0, "Informations d'acier prise avec réussite !", "Plugin Revit", 0)
    return RebarsDiamChoisi, RebarsShapeChoisi, RebarHook

def verifier_creer_crochet_acier(Angle):
    """Créer crochet."""
    Rebars1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType()
    CrochetsDispo = []
    for l in Rebars1:
        if l.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsValueString() == "Crochet d'armature":
            CrochetsDispo.append(l)

    AngleString = "Crochet d'attache/étrier - " + str(int(Angle)) + " deg."
    for p in CrochetsDispo:
        if p.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() == AngleString:
            # ctypes.windll.user32.MessageBoxW(0, "FamilleType Crochet déjà existente !", "Plugin Revit", 0)
            return p.Id

    tc = Transaction(doc, "Ajouter Type Crochet")
    tc.Start()
    newCrochettype = CrochetDefault.Duplicate(AngleString)
    newCrochettype.get_Parameter(BuiltInParameter.REBAR_HOOK_ANGLE).Set(float(Angle))
    tc.Commit()

    return newCrochettype.Id

def rebar_longitudinal_prin(wallsID, DonneeDepart, vectorDIR, VN, Enrobage, RecEle):
    """All steps for principal longitudinal rebars."""
    Recouvrement_dummy = 0
    ElementDiam = 0
    RebarsCreated = []
    DiamPrincipal = float(DonneeDepart["DiamLong"]) / 304.8
    EN = Enrobage * 3.28084

    for i in DiamElements:
        if str(int(DonneeDepart["DiamLong"]))+" mm" == i.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsValueString():
            ElementDiam = i
            break

    for i in wallsID:
        # Donnée de base du mur d'intérêt
        Mur = doc.GetElement(i)
        MurLargueur = Mur.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
        MurLongueur = Mur.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
        MurHauteur = Mur.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()

        # Prendre le milieu de la partie haute du Mur
        bounding_box = Mur.get_BoundingBox(default_3d_view)
        TRF_point = bounding_box.Max
        LLB_point = bounding_box.Min
        mid_point = XYZ((TRF_point.X + LLB_point.X) / 2, (TRF_point.Y + LLB_point.Y) / 2, TRF_point.Z)


        # Si element pair, donc déplacer plus 4 cm
        ElementAnalise = int(Mur.LookupParameter("Element").AsValueString()) - 1

        if ElementAnalise % 2 == 0:
            DisplacementX = 0
        else:
            DisplacementX = DiamPrincipal + 3/304.8    # 3 mm en plus de distance

        espacement = (MurLongueur - ((Enrobage+0.05)*2*3.28084) - DiamPrincipal - DisplacementX) / \
                     (int(DonneeDepart["QAcierLong"])-1)


        # Determiner si la position des aciers selon recouvrement necessaires
        if Recouvrement_dummy == 0:                                                             # Debut de la cage
            recouvrement = RecEle[ElementAnalise] * 3.281 / 2
            Recouvrement_dummy += 1
            PositionCage = 1
        elif Recouvrement_dummy == len(RecEle):                                                 # Fin de la cage
            recouvrement = RecEle[ElementAnalise-1] * 3.281 / 2
            Recouvrement_dummy = 0
            PositionCage = 3
        else:                                                                                   # Milieu
            recouvrement = (RecEle[ElementAnalise] + RecEle[ElementAnalise-1]) * 3.281 / 2
            Recouvrement_dummy += 1
            PositionCage = 2
        longueur_barre = MurHauteur - EN + recouvrement


        # Prendre Ligne et normal de l'acier, INDIVIDUELLE !!
        if PositionCage == 1:
            z = mid_point.Z - EN * 2
        elif PositionCage == 2:
            z = mid_point.Z + recouvrement / 2
        else:
            z = mid_point.Z + recouvrement

        DisplacementEnX = DisplacementX + DispTrans + EN - MurLongueur / 2
        DisplacementEnY = EN + DisplacementY - MurLargueur / 2

        if DonneeDepart["Zone"] == "CT":
            DisplacementEnY = DisplacementEnY * -1

        point_start = XYZ(mid_point.X + (VN.X * DisplacementEnY) + (vectorDIR.X * DisplacementEnX),
                          mid_point.Y + (VN.Y * DisplacementEnY) + (vectorDIR.Y * DisplacementEnX), z)
        point_end = XYZ(point_start.X, point_start.Y, z - longueur_barre)

        ligne = Line.CreateBound(point_start, point_end)
        L_Long = List[Curve]()
        L_Long.Add(ligne)
        norm = VN.CrossProduct(vectorDIR)


        # Créer Armature Long
        NewRebar = Structure.Rebar.CreateFromCurves(doc, RS, ElementDiam, None, None, Mur, vectorDIR, L_Long,
                                                    HookRight, HookLeft, True, True)


        # Modifier pour répeter Acier selon quantité et espacement calculé
        NewRebar.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)                              # type conf.
        NewRebar.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(DonneeDepart["QAcierLong"])  # Quantité
        NewRebar.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(espacement)                     # Espacement
        NewRebar.GetShapeDrivenAccessor().BarsOnNormalSide = True
        NewRebar.LookupParameter("Cage").Set(Mur.LookupParameter("Cage").AsValueString())
        NewRebar.LookupParameter("Element").Set(Mur.LookupParameter("Element").AsValueString())
        NewRebar.LookupParameter("Coupe").Set(Mur.LookupParameter("Coupe").AsValueString())
        NewRebar.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).Set(DonneeDepart["NomBarre"])
        NewRebar.get_Parameter(BuiltInParameter.NUMBER_PARTITION_PARAM).Set("Longitudinal principal")
        RebarsCreated.append(NewRebar)

    return RebarsCreated

def rebar_longitudinal_renfort(DonneeDepart, DonneeRenfort, vectorPER, VN, Enrobage, AS, FM):
    """All steps for principal longitudinal rebars."""
    EN = Enrobage * 3.28084
    MurRenfort = 0
    MurRenfortDebut = 0
    ElementDiam = 0
    AcierRenfort = []
    QuantRF = int(DonneeDepart["QAcierLong"])
    DiamLong = float(DonneeDepart["DiamLong"]) / 304.8


    # Chercher le mur spécifique
    for i in Walls:

        if i.LookupParameter("Cage").AsValueString() == DonneeRenfort["Cage"] and \
           i.LookupParameter("Element").AsValueString() == DonneeRenfort["Element"] and \
           i.LookupParameter("Coupe").AsValueString() == DonneeRenfort["Coupe"]:
            MurRenfort = i
            break

    MurRLargueur = MurRenfort.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    MurRLongueur = MurRenfort.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    long_barre_renfort = (DonneeRenfort["AraseSUP"] - DonneeRenfort["AraseINF"]) * 3.28084


    # Prendre le Element 01 de la cage d'interet
    for i in Walls:
        if i.LookupParameter("Cage").AsValueString() == DonneeRenfort["Cage"] and \
           i.LookupParameter("Element").AsValueString() == '01' and \
           i.LookupParameter("Coupe").AsValueString() == DonneeRenfort["Coupe"]:
            MurRenfortDebut = i


    # Prendre la hauteur selon unites de Revit
    bounding_box = MurRenfortDebut.get_BoundingBox(default_3d_view)
    TRF_pointR = bounding_box.Max
    LLB_pointR = bounding_box.Min
    mid_pointR = XYZ((TRF_pointR.X + LLB_pointR.X) / 2, (TRF_pointR.Y + LLB_pointR.Y) / 2, TRF_pointR.Z)
    z = TRF_pointR.Z - (AS - DonneeRenfort["AraseSUP"]) * 3.281


    # Chercher le TYPE de DIAMETRE
    for i in DiamElements:
        if str(int(DonneeDepart["DiamLong"]))+" mm" == i.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsValueString():
            ElementDiam = i
            break


    # Fonction prendre ligne armatures
    DispGroupe = (int(DonneeRenfort["Groupe"])-1)*(float(DonneeDepart["DiamLong"]) + 80)/304.8
    DispLit = (int(DonneeRenfort["Lit"]) - 1) * (float(DonneeDepart["DiamLong"]) + 15) / 304.8


    # Displacement pour les elements paires
    if int(DonneeRenfort["Element"]) % 2 == 0:
        DispX = (float(DonneeDepart["DiamLong"]) + 3) / 304.8  # 3 mm en plus de distance
    else:
        DispX = 0

    espacement = (MurRLongueur - ((Enrobage + 0.05) * 2 * 3.28084) - DiamLong - DispX) / \
                 (int(DonneeDepart["QAcierLong"]) - 1)
    DisplacementEnX = (DispTrans + EN + DispX) - MurRLongueur / 2
    DisplacementEnY = (DispLit + DispGroupe + EN + DisplacementY) - MurRLargueur / 2

    point_start = XYZ(mid_pointR.X + (VN.X * DisplacementEnY) + (vectorPER.X * DisplacementEnX),
                      mid_pointR.Y + (VN.Y * DisplacementEnY) + (vectorPER.Y * DisplacementEnX), z)
    point_end = XYZ(point_start.X, point_start.Y, z - long_barre_renfort)

    ligneRF = Line.CreateBound(point_start, point_end)
    L_Long_Rf = List[Curve]()
    L_Long_Rf.Add(ligneRF)


    # Créer Armature Long
    NewRebarRf = Structure.Rebar.CreateFromCurves(doc, RS, ElementDiam, None, None, MurRenfort, vectorPER, L_Long_Rf,
                                                  HookRight, HookLeft, True, True)

    NewRebarRf.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)                # type conf.
    NewRebarRf.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(QuantRF)     # Quantité
    NewRebarRf.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(espacement)       # Espacement
    NewRebarRf.GetShapeDrivenAccessor().BarsOnNormalSide = True
    NewRebarRf.LookupParameter("Cage").Set(MurRenfort.LookupParameter("Cage").AsValueString())
    NewRebarRf.LookupParameter("Element").Set(MurRenfort.LookupParameter("Element").AsValueString())
    NewRebarRf.LookupParameter("Coupe").Set(MurRenfort.LookupParameter("Coupe").AsValueString())
    NewRebarRf.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).Set(DonneeDepart["NomBarre"])
    NewRebarRf.get_Parameter(BuiltInParameter.NUMBER_PARTITION_PARAM).Set("Longitudinal Renfort")


    AcierRenfort.append(NewRebarRf)

    return AcierRenfort

def modifier_format_posisionnement(NewRebarTR, Enrobage, MurTransDebut, VN, vectorPER):
    """Modifier format et la position d'acier tranversal."""
    MurTLargueur = MurTransDebut.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    MurTLongueur = MurTransDebut.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()
    EN = Enrobage * 3.28084

    DimA = MurTLongueur - EN *2 - 20/304.8
    DimB = MurTLargueur - EN *2 - 15/304.8
    DimC = DimA - 20/304.8
    DimD = DimA
    DimE = DimB

    NewRebarTR.LookupParameter("A").Set(DimA)
    NewRebarTR.LookupParameter("B").Set(DimB)
    NewRebarTR.LookupParameter("C").Set(DimC)
    NewRebarTR.LookupParameter("D").Set(DimD)
    NewRebarTR.LookupParameter("E").Set(DimE)

    NewRebarTR.LookupParameter("C1").Set(0)
    NewRebarTR.LookupParameter("C2").Set(0)
    NewRebarTR.LookupParameter("D1").Set(0)
    NewRebarTR.LookupParameter("D2").Set(0)
    NewRebarTR.LookupParameter("E1").Set(0)
    NewRebarTR.LookupParameter("E2").Set(0)
    NewRebarTR.LookupParameter("F").Set(0)
    NewRebarTR.LookupParameter("F1").Set(0)
    NewRebarTR.LookupParameter("F2").Set(0)
    NewRebarTR.LookupParameter("G").Set(0)
    NewRebarTR.LookupParameter("G1").Set(0)
    NewRebarTR.LookupParameter("G2").Set(0)
    NewRebarTR.LookupParameter("H").Set(0)
    NewRebarTR.LookupParameter("J").Set(0)
    NewRebarTR.LookupParameter("K").Set(0)
    NewRebarTR.LookupParameter("L").Set(0)
    NewRebarTR.LookupParameter("R").Set(0)

    # Prendre la hauteur selon unités de Revit
    bounding_box = NewRebarTR.get_BoundingBox(default_3d_view)
    TRF_pointR = bounding_box.Max
    LLB_pointR = bounding_box.Min
    mid_pointR = XYZ((TRF_pointR.X + LLB_pointR.X) / 2, (TRF_pointR.Y + LLB_pointR.Y) / 2, TRF_pointR.Z)

    # Prendre la hauteur selon unités de Revit
    bounding_box_Mur = MurTransDebut.get_BoundingBox(default_3d_view)
    TRF_pointR_Mur = bounding_box_Mur.Max
    LLB_pointR_Mur = bounding_box_Mur.Min
    mid_pointR_Mur = XYZ((TRF_pointR_Mur.X + LLB_pointR_Mur.X) / 2, (TRF_pointR_Mur.Y + LLB_pointR_Mur.Y) / 2,
                         LLB_pointR_Mur.Z)

    new_loc = XYZ((mid_pointR_Mur.X-mid_pointR.X)*2, (mid_pointR_Mur.Y-mid_pointR.Y)*2, 0)

    ElementTransformUtils.MoveElement(doc, NewRebarTR.Id, new_loc)

def rebar_transversal_principal(DonneeTrans, vectorPER, VN, Enrobage, AS, wallsID):
    """All steps for principal transversal rebars."""
    # INNOVATION 01 : PRENDRE MUR OU L'ARASE EST ET PAS LE DEBUT
    EN = Enrobage * 3.28084

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
    MurTransDebut = 0
    for i in wallsID:
        MurDummy = doc.GetElement(i)
        # Donnée de base de la cage d'intérêt et le premier element de la cage
        if MurDummy.LookupParameter("Cage").AsValueString() == DonneeTrans["Cage"] and \
           MurDummy.LookupParameter("Element").AsValueString() == '01':
            MurTransDebut = MurDummy


    # Créer les armatures transversales principales pour chaque

    NewRebarTR = 0
    MurTLargueur = MurTransDebut.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()
    MurTLongueur = MurTransDebut.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()

    # Prendre la hauteur selon unités de Revit
    bounding_box = MurTransDebut.get_BoundingBox(default_3d_view)
    TRF_pointR = bounding_box.Max
    LLB_pointR = bounding_box.Min
    mid_pointR = XYZ((TRF_pointR.X + LLB_pointR.X) / 2, (TRF_pointR.Y + LLB_pointR.Y) / 2, TRF_pointR.Z)
    reference = (AS - DonneeTrans["AraseSup"]) * 3.28084
    z = mid_pointR.Z - reference - DonneeTrans["QAcierTrans"] * DonneeTrans["Espacement"]

    DisplacementEnX = EN - MurTLongueur / 2
    DisplacementEnY = EN - MurTLargueur / 2

    point_start = XYZ(mid_pointR.X + (VN.X * DisplacementEnY) + (vectorPER.X * DisplacementEnX),
                      mid_pointR.Y + (VN.Y * DisplacementEnY) + (vectorPER.Y * DisplacementEnX), z)

    NewRebarTR = Structure.Rebar.CreateFromRebarShape(doc, ElementShape, ElementDiam, MurTransDebut, point_start,
                                                      vectorPER, VN)

    # Paramètres de configuration
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_LAYOUT_RULE).Set(3)                                # type conf.
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS).Set(DonneeTrans["QAcierTrans"])  # Quantité
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_BAR_SPACING).Set(DonneeTrans["Espacement"])        # Espacement
    NewRebarTR.GetShapeDrivenAccessor().BarsOnNormalSide = True
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_HOOK_LENGTH_OVERRIDE).Set(DonneeTrans["CrochetLong"])
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_HOOK_START_TYPE).Set(DonneeTrans["Crochet"].IntegerValue)
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_HOOK_END_TYPE).Set(DonneeTrans["Crochet"].IntegerValue)



    # Paramètres d'information
    NewRebarTR.LookupParameter("Cage").Set(MurTransDebut.LookupParameter("Cage").AsValueString())
    NewRebarTR.LookupParameter("Element").Set(MurTransDebut.LookupParameter("Element").AsValueString())
    NewRebarTR.LookupParameter("Coupe").Set(MurTransDebut.LookupParameter("Coupe").AsValueString())
    NewRebarTR.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK).Set(DonneeTrans["NomBarre"])
    NewRebarTR.get_Parameter(BuiltInParameter.NUMBER_PARTITION_PARAM).Set("Transversale principale")

    modifier_format_posisionnement(NewRebarTR, Enrobage, MurTransDebut, VN, vectorPER)

    return NewRebarTR


# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

RebarsShapeNames = {'S1': 'A FAIRE', 'S2': '25', 'S3': '2-00', 'S4': '3-00', 'S5': '3-00', 'T1': '6-20','T2': '5-27',
                    'M1': '1-01', 'L1': '0-00', 'L2': '2-05'}

RebarsShapeNamesList = ["25", '2-00', '3-00', '6-20', '5-27', '1-01', '0-00', '2-05']

RebarsDiamNames = ['HA6 (Fe500)', 'HA8 (Fe500)', 'HA10 (Fe500)', 'HA12 (Fe500)', 'HA14 (Fe500)', 'HA16 (Fe500)',
                   'HA20 (Fe500)', 'HA25 (Fe500)', 'HA32 (Fe500)', 'HA40 (Fe500)']

HookNames = ["90", "87"]

view_collector = FilteredElementCollector(doc).OfClass(View3D).ToElements()
default_3d_view = next((view for view in view_collector if view.IsTemplate == False), None)

HookLeft = Structure.RebarHookOrientation.Left
HookRight = Structure.RebarHookOrientation.Right
RS = Structure.RebarStyle.Standard
RT = Structure.RebarStyle.StirrupTie

Walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
Rebars = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsElementType()

CrochetDefault = doc.GetElement(ElementId(144769))  # Crochet de base  90°CrochetsDispo = []

# Prendre des elements sur Revit des types de Diamètres, formes et crochets
DiamElements, FormeElements, HookElements = catch_rebar_details(RebarsDiamNames, RebarsShapeNamesList, HookNames)

DispTrans = 0.06 * 3.28084
DisplacementY = 0.06 * 3.28084

# CODE  -------------------------------------------------------------------------------------------------------
