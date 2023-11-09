# -*- coding: utf-8 -*-

__title__ = "_MurDonnee"  # Name of button
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
from pyrevit import forms, revit, UI
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)

# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
rvt_year = int(app.VersionNumber)
snap_mode = UI.Selection.ObjectSnapTypes.Endpoints

# FUNCTIONS -----------------------------------------------------------------------------------------------

def questions_depart():
    """Ask for a series of initial info to produce diaphragm wall."""
    # INNOVATION 01 : En generale, il manque faire des "try and except" pour terminer le programme sans donner d'erreurs
    # INNOVATION 02 : Les recouvrements sont normées, choisir d'une liste !
    DictDonnee = {}
    with forms.ProgressBar(tittle='Progress de donnée de entrée') as pb:

        condition = True
        while condition:
            components = [Label('Nom de la coupe :'), TextBox('NC', Text="01"),
                          Separator(),
                          Label('Arrase supérieur :'), TextBox('AS', Text="100.00"),
                          Label('Fiche Mécanique :'), TextBox('FM', Text="70.00"),
                          Separator(),
                          Label('Quantité de éléments [entre 1 et 5] :'), TextBox('QE', Text="3"),
                          Label('Quantité de cages [entre 1 et 5] :'), TextBox('QC', Text="3"),
                          Label('Quantité de dalles [entre 0 et 10] :'), TextBox('QD', Text="3"),
                          Separator(),
                          Label('Enrobage [entre 0.01 et 0.20]:'), TextBox('EN', Text="0.075"),
                          Label('Épaisseur Paroi [entre 0.02 et 2.00]:'), TextBox('EP', Text="1.00"),
                          Label('Longueur du Paroi [entre 0.50 et 20.00]:'), TextBox('LO', Text="7.00"),
                          Button('Choisir')]

            form = FlexForm('Questions de depart tout en [m/unités]', components)

            form.show()
            values = form.values

            DictDonnee['NC'] = values['NC']
            DictDonnee['AS'] = float(values['AS'])
            DictDonnee['FM'] = float(values['FM'])
            DictDonnee['QE'] = int(values['QE'])
            DictDonnee['QC'] = int(values['QC'])
            DictDonnee['QD'] = int(values['QD'])
            DictDonnee['EN'] = float(values['EN'])
            DictDonnee['EP'] = float(values['EP'])
            DictDonnee['LO'] = float(values['LO'])

            if DictDonnee['FM'] > DictDonnee['AS']:
                ctypes.windll.user32.MessageBoxW(0, "Fiche mécanique plus haut que Arrase !", "Plugin Revit", 0)
            elif DictDonnee['AS']-DictDonnee['FM'] > 60:
                ctypes.windll.user32.MessageBoxW(0, "Paroi moulée trop elevée, maximum 60 metres !", "Plugin Revit", 0)
            elif DictDonnee['QE'] < 1 or DictDonnee['QE'] > 5:
                ctypes.windll.user32.MessageBoxW(0, "Quantité d'éléments invalide !", "Plugin Revit", 0)
            elif DictDonnee['QC'] < 1 or DictDonnee['QC'] > 5:
                ctypes.windll.user32.MessageBoxW(0, "Quantité de cages invalide !", "Plugin Revit", 0)
            elif DictDonnee['QD'] < 0 or DictDonnee['QD'] > 10:
                ctypes.windll.user32.MessageBoxW(0, "Quantité de dalles invalide !", "Plugin Revit", 0)
            elif DictDonnee['EN'] < 0.01 or DictDonnee['EN'] > 0.20:
                ctypes.windll.user32.MessageBoxW(0, "Enrobage invalide !", "Plugin Revit", 0)
            elif DictDonnee['EP'] < 0.2 or DictDonnee['EP'] > 2:
                ctypes.windll.user32.MessageBoxW(0, "Épaisseur invalide !", "Plugin Revit", 0)
            elif DictDonnee['LO'] <= 0 or DictDonnee['LO'] > 20:
                ctypes.windll.user32.MessageBoxW(0, "Longueur invalide !", "Plugin Revit", 0)
            else:
                condition = False


    # Comprendre hauteurs elements, division cages et position dalles ----------------------------
        pb.update_progress(50, 100)
        HauteurParoi = DictDonnee['AS']-DictDonnee['FM']

    # Quantité d'éléments ----------------------------------------------------------------------------
    # INNOVATION 01 : Mettre attention utilisateur avec longueur acier + grand que 15m/de ne pas permettre
        conditionElements = True

        if DictDonnee['QE'] == 1:
            HauteursElements = HauteurParoi
            RecouvrementElements = 0

        else:
            ctypes.windll.user32.MessageBoxW(0, "Maintenant, sur la division des éléments", "Plugin Revit", 0)
            while conditionElements:
                conditionQE = False
                HauteursElements = []
                RecouvrementElements = []
                for i in range(DictDonnee['QE']):
                    LongueurDummy = float(forms.ask_for_string(default='10.0',
                                          prompt='Longueur total élément '+str(i+1)+' :',
                                          title='Etude elements'))

                    if LongueurDummy <= 0 or LongueurDummy >= HauteurParoi:
                        ctypes.windll.user32.MessageBoxW(0, "Longueur d'élément invalide !", "Plugin Revit", 0)
                        conditionQE = True
                        break
                    else:
                        if LongueurDummy > 15:
                            ctypes.windll.user32.MessageBoxW(0, "ATTENTION: Longueur plus grande que recommandé de 15m",
                                                             "Plugin Revit", 0)
                        HauteursElements.append(LongueurDummy)

                    if i != 0:
                        RecouvrementDummy = float(forms.ask_for_string(default='1.40',
                                                  prompt='Recouvrement entre element '+str(i+1)+' et '+str(i),
                                                  title='Etude elements'))  # Dans l'avenir pas necessaire demander

                        if RecouvrementDummy <= 0 or RecouvrementDummy > 3:
                            ctypes.windll.user32.MessageBoxW(0, "Recouvrement entre éléments invalide !",
                                                             "Plugin Revit", 0)
                            conditionQE = True
                            break
                        else:
                            RecouvrementElements.append(RecouvrementDummy)

                if LongueurDummy <= 0 or RecouvrementDummy <= 0 or RecouvrementDummy > 3:
                    pass

                elif sum(HauteursElements) != HauteurParoi or conditionQE:
                    ctypes.windll.user32.MessageBoxW(0, "Sommes des hauteurs est " + str(sum(HauteursElements)) +
                                                     ' mais il doit être : ' + str(HauteurParoi), "Plugin Revit", 0)

                else:
                    conditionElements = False



    # Quantité de cages ----------------------------------------------------------------------------------------------
    # INNOVATION 02 : Pas necessaire demander la derniere longueur, il faut juste faire longueurMax - LongueursDivisions
        pb.update_progress(75, 100)
        LongueurMax = DictDonnee['LO']
        conditionCage = True

        if DictDonnee['QC'] == 1:
            LongueursDivisions = DictDonnee['LO']

        else:
            ctypes.windll.user32.MessageBoxW(0, "Maintenant, sur la division des cages", "Plugin Revit", 0)
            while conditionCage:
                conditionQC = False
                LongueursDivisions = []

                boolAsk = forms.ask_for_one_item(['Oui', 'Non'], default='Oui',
                                                 prompt='Est-ce que les cages ont dimensions pareil ?',
                                                 title='Dimensions cages')

                if boolAsk == "Oui":
                    ValeurDummy = LongueurMax / DictDonnee['QC']
                    for i in range(DictDonnee['QC']):
                        LongueursDivisions.append(ValeurDummy)
                    conditionCage = False

                else:
                    for i in range(DictDonnee['QC']):
                        if i == range(DictDonnee['QC']):
                            LongueursCageDummy = LongueurMax - sum(LongueursDivisions)
                            LongueursDivisions.append(LongueursCageDummy)

                        else:
                            LongueursCageDummy = float(forms.ask_for_string(default='1.50',
                                                       prompt='Longueur cage '+str(i+1)+' :',
                                                       title='Etude cages'))

                            if LongueursCageDummy <= 0 or LongueursCageDummy > DictDonnee['LO'] or \
                               sum(LongueursDivisions) > LongueurMax:
                                ctypes.windll.user32.MessageBoxW(0, "Longueur cage invalide !", "Plugin Revit", 0)
                                conditionQC = True
                                break
                            else:
                                LongueursDivisions.append(LongueursCageDummy)

                    if LongueursCageDummy <= 0 or LongueursCageDummy > DictDonnee['LO'] or \
                       sum(LongueursDivisions) > LongueurMax:
                        pass

                    elif conditionQC:
                        ctypes.windll.user32.MessageBoxW(0, "Somme du longueur de cages impossible, refaire !",
                                                         "Plugin Revit", 0)

                    else:
                        conditionCage = False

    # Quantité de dalles ---------------------------------------------------------------------------------
    # INNOVATION 03 : Faire une forms complet en demandant la longueur de la dalle, les efforts sur l'appui et epaisseur
    # INNOVATION 04 : Prendre informations des dalles connectées en revit pour les informations de l'INNOVATION 03
        pb.update_progress(90, 100)
        conditionDalle = True

        if DictDonnee['QD'] == 0:
            HauteurDalles = 0

        else:
            ctypes.windll.user32.MessageBoxW(0, "Maintenant, sur la division des dalles", "Plugin Revit", 0)
            while conditionDalle:
                conditionQD = False
                HauteurDalles = []

                for i in range(DictDonnee['QD']):
                    HauteurCageDummy = float(forms.ask_for_string(default='0.00',
                                                                  prompt='Niveau dalle ' + str(i + 1) +
                                                                  ' [entre '+str(DictDonnee['FM'])+' et '+str(DictDonnee['AS'])+'] :',
                                                                  title='Etude cages'))

                    if HauteurCageDummy < DictDonnee['FM'] or HauteurCageDummy > DictDonnee['AS']:
                        ctypes.windll.user32.MessageBoxW(0, "Niveau dalle invalide !", "Plugin Revit", 0)
                        conditionQD = True
                        break

                    else:
                        HauteurDalles.append(HauteurCageDummy)

                        if 0 <= i-1 < len(HauteurDalles):
                            if HauteurDalles[i-1] < HauteurDalles[i]:
                                ctypes.windll.user32.MessageBoxW(0, "Niveau doit être inférieur de la precedente !",
                                                                 "Plugin Revit", 0)
                                conditionQD = True
                                break

                if conditionQD:
                    conditionDalle = True

                else:
                    conditionDalle = False

    pb.update_progress(100, 100)
    DictDonnee['HauteursElements'] = HauteursElements
    DictDonnee['RecouvrementElements'] = RecouvrementElements
    DictDonnee['LongueursDivisions'] = LongueursDivisions
    DictDonnee['HauteurDalles'] = HauteurDalles
    ctypes.windll.user32.MessageBoxW(0, "Operation terminée avec succès !", "Plugin Revit", 0)
    return DictDonnee

def create_and_apply_cover(enrobage, murIDs):
    """Create Covertype and apply to walls created."""
    # ctypes.windll.user32.MessageBoxW(0, "Processus verifier/creer Nouveau Enrobage type", "Plugin Revit", 0)
    Coverdist = enrobage*3.281
    ListCovers = []
    Covers = FilteredElementCollector(doc).OfCategoryId(ElementId(-2009014))  # .WhereElementIsElementType()
    WallElement = doc.GetElement(murIDs[0])
    WallCoverExample = WallElement.get_Parameter(BuiltInParameter.CLEAR_COVER_EXTERIOR)
    CoverDefault = doc.GetElement(WallCoverExample.AsElementId())
    CoverCondition = True
    Coverdummy = 0

    for i in Covers:
        if i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() == \
                "Enrobage de "+str(int(enrobage*1000))+" mm":
            Coverdummy = i
            CoverCondition = False
            # ctypes.windll.user32.MessageBoxW(0, "FamilleType enrobage déjà existente!", "Plugin Revit", 0)
            break

    t = Transaction(doc, "Modifier Enrobage de PM")
    t.Start()

    if CoverCondition:
        newCovertype = CoverDefault.Duplicate("Enrobage de "+str(int(enrobage*1000))+" mm")
        CoverLength = newCovertype.get_Parameter(BuiltInParameter.COVER_TYPE_LENGTH)
        CoverLength.Set(Coverdist)

    else:
        newCovertype = Coverdummy

    for i in murIDs:
        wallDummy = doc.GetElement(i)
        wallDummy.get_Parameter(BuiltInParameter.CLEAR_COVER_EXTERIOR).Set(newCovertype.Id)
        wallDummy.get_Parameter(BuiltInParameter.CLEAR_COVER_INTERIOR).Set(newCovertype.Id)
        wallDummy.get_Parameter(BuiltInParameter.CLEAR_COVER_OTHER).Set(newCovertype.Id)

    t.Commit()

    # ctypes.windll.user32.MessageBoxW(0, "FamilleType d'Enrobage placé avec réussite!", "Plugin Revit", 0)
    return newCovertype.Id

def create_family_type_wall(epaisseur):
    """Create family type of Wall with specif Parameters."""
    # ctypes.windll.user32.MessageBoxW(0, "Processus Nouveau famille type", "Plugin Revit", 0)

    # all_materials = list(FilteredElementCollector(doc).OfClass(Material).ToElements())
    # dict_materials = {mat.Name: mat for mat in all_materials}
    # all_materials.sort(key=lambda x: x.Name)
    # Béton coulé sur place C25 - 144141
    # Béton coulé sur place  - 1250
    ConditionDummy = False
    walltypeID = "dummy"
    listnames = []
    my_elements = FilteredElementCollector(doc).OfClass(WallType).WhereElementIsElementType()

    for i in my_elements:
        listnames.append(i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString())

        if i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsValueString() == "PM "+str(epaisseur)+"0":
            walltypeID = i.Id
            # ctypes.windll.user32.MessageBoxW(0, "FamilleType deja existente!", "Plugin Revit", 0)
            return walltypeID

    # murbase = doc.GetElement(doc.GetDefaultElementTypeId(ElementTypeGroup.WallType))  # Famille default
    murbase = doc.GetElement(ElementId(451454))   # Il a pris la famille de PM du gabarit

    t = Transaction(doc, "Insérer Famille PM")
    t.Start()

    newwalltype = murbase.Duplicate("PM "+str(epaisseur)+"0")
    newwalltypeID = newwalltype.Id
    cs = newwalltype.GetCompoundStructure()
    i = cs.GetFirstCoreLayerIndex()
    cs.SetLayerWidth(i, epaisseur*3.28083989501)
    # cs.SetMaterialId(i, ElementId(1250))        # Definir matériaux comme Béton coule sur place
    newwalltype.SetCompoundStructure(cs)

    t.Commit()

    # ctypes.windll.user32.MessageBoxW(0, "FamilleType crée avec success!", "Plugin Revit", 0)
    return newwalltypeID

def select_structural_wall():
    """Requests the user to select a structural wall in Revit."""
    ctypes.windll.user32.MessageBoxW(0, "Processus supprimer mur existent et prendre vecteur/point", "Plugin Revit", 0)
    faire_loop = True
    while faire_loop:

        with forms.WarningBar(title='Choisir paroi moulée de base:'):
            ctypes.windll.user32.MessageBoxW(0, "Choisir Mur", "Plugin Revit", 0)
            element = revit.pick_element()

        if element.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT):
            faire_loop = False
        else:
            ctypes.windll.user32.MessageBoxW(0, "S'il vous plait choisir un Mur", "Plugin Revit", 0)

    ctypes.windll.user32.MessageBoxW(0, "Choisir point gauche en haut et apres en bas avec face extérieur en haut",
                                     "Plugin Revit", 0)
    point_ext = uidoc.Selection.PickPoint(snap_mode)
    point_int = uidoc.Selection.PickPoint(snap_mode)

    x = (point_ext.X + point_int.X) / 2
    y = (point_ext.Y + point_int.Y) / 2
    z = (point_ext.Z + point_int.Z) / 2

    # Create the middle point.
    middle_point = XYZ(x, y, z)

    # Vector / Direction
    vector = XYZ(point_ext.X - point_int.X, point_ext.Y - point_int.Y, point_ext.Z - point_int.Z)
    perpendicular_vector = XYZ(vector.Y, -vector.X, vector.Z).Normalize()

    t = Transaction(doc, "Supprimer Mur")
    t.Start()
    doc.Delete(element.Id)
    t.Commit()

    ctypes.windll.user32.MessageBoxW(0, "Mur supprimé et vecteur/point pris avec réussite", "Plugin Revit", 0)
    return perpendicular_vector, vector, middle_point


def choisir_mur_acier():
    """Requests the user to select a structural wall the is parte of the section analised."""
    #ctypes.windll.user32.MessageBoxW(0, "Processus supprimer mur existent et prendre vecteur/point", "Plugin Revit", 0)
    faire_loop = True
    while faire_loop:

        with forms.WarningBar(title='Choisir paroi moulée qui fait partie de la analyse :'):
            ctypes.windll.user32.MessageBoxW(0, "Choisir Mur", "Plugin Revit", 0)
            element = revit.pick_element()

        if element.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT):
            faire_loop = False
        else:
            ctypes.windll.user32.MessageBoxW(0, "S'il vous plait choisir un Mur", "Plugin Revit", 0)

    # DEJA PRISE SUR EXCEL
    # CoupeDesiree = element.LookupParameter("Coupe")
    # Enro = element.get_Parameter(BuiltInParameter.CLEAR_COVER_OTHER).AsValueString()
    # first_space = Enro.find(" ")
    # second_space = Enro.find(" ", first_space + 1)
    # third_space = Enro.find(" ", second_space + 1)
    # Enrobage = float(Enro[second_space + 1:third_space])/1000

    ctypes.windll.user32.MessageBoxW(0, "Choisir point gauche en haut et en bas, face ext. en haut", "Plugin Revit", 0)
    point_ext = uidoc.Selection.PickPoint(snap_mode)
    point_int = uidoc.Selection.PickPoint(snap_mode)

    # Create the middle point.
    middlepoint = XYZ((point_ext.X + point_int.X) / 2, (point_ext.Y + point_int.Y) / 2, (point_ext.Z + point_int.Z) / 2)

    # Vector / Direction
    vectorPAR = XYZ(point_ext.X - point_int.X, point_ext.Y - point_int.Y, point_ext.Z - point_int.Z).Normalize()
    vectorPER = XYZ(vectorPAR.Y, -vectorPAR.X, vectorPAR.Z).Normalize()

    # ctypes.windll.user32.MessageBoxW(0, "Mur  et vecteur/point pris avec réussite", "Plugin Revit", 0)
    return vectorPER, vectorPAR, middlepoint, element


def create_diaphragm_wall(vectorD, start_point, HauteursElements, LongueursDivisions, PMid, NomCoupe):
    """Create diaphragm  wall based considering division by element and cage."""
    ctypes.windll.user32.MessageBoxW(0, "Processus créer Paroi Moulée", "Plugin Revit", 0)

    intensity = 0
    WallIDs = []
    niveau = forms.select_levels(title='Choisir niveau supérieur', button_name='Choisir',
                                 width=500, multiple=False, filterfunc=None, doc=None, use_selection=False)

    t = Transaction(doc, "Créer les PMs")
    t.Start()

    for y in range(len(LongueursDivisions)):    # Chaque cage
        offset = 0
        intensity = LongueursDivisions[y]*3.28084

        end_point = XYZ(start_point.X + vectorD.X * intensity, start_point.Y + vectorD.Y * intensity,
                        start_point.Z + vectorD.Z * intensity)
        ligne = Line.CreateBound(start_point, end_point)
        start_point = end_point

        for i in range(len(HauteursElements)):  # Chaque Element
            # newWall = Wall.Create(doc, ligne, niveau.Id, True) FONCTIONNE, de gauche a droit, extérieur en haut !
            offset = offset - HauteursElements[i] * 3.28084
            wall = Wall.Create(doc, ligne, PMid, niveau.Id, HauteursElements[i]*3.28084, offset, False, True)  # Flip et structural
            WallIDs.append(wall.Id)
            WallUtils.DisallowWallJoinAtEnd(wall, 0)
            wall.LookupParameter("Cage").Set("0"+str(y+1))
            wall.LookupParameter("Element").Set("0" + str(i+1))
            wall.LookupParameter("Coupe").Set(NomCoupe)

    t.Commit()

    # ctypes.windll.user32.MessageBoxW(0, "Paroi Moulée crée avec réussite", "Plugin Revit", 0)
    return WallIDs

# CODE  -------------------------------------------------------------------------------------------------------
