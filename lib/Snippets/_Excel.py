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
import pyrevit
from Autodesk.Revit.DB import *
from pyrevit import forms, revit, UI, DB
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)

# VARIABLES  -----------------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
PATH_SCRIPT = os.path.dirname(__file__)
rvt_year = int(app.VersionNumber)

# FUNCTIONS -----------------------------------------------------------------------------------------------

def get_last_row_and_column(worksheet):
    """Gets the last row and column with data in the specified worksheet."""
    last_row = -1
    last_column = -1


    for row in range(worksheet.Rows.Count):
        conditionDummy = False

        for col in range(worksheet.Columns.Count):
            cell = worksheet.Cells(row+1, col+1)

            if cell.Value2 is None:
                conditionDummy = True
                break

            else:
                last_row = row+1
                last_column = col+1

        cell01 = worksheet.Cells(row+1,1)
        if conditionDummy and cell01.Value2 is None:
            break

    return last_row, last_column

def read_excel_file():
    """
    Read all data in an Excel file and sends back to main script
    """
    ctypes.windll.user32.MessageBoxW(0, "S'il vous plait donner chemin et nom pour du fichier Excel", "Plugin Revit", 0)
    donnee_excel = {}

    # LIRE EXCEL
    # Create Excel application object
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False

    fileImport = forms.pick_file()

    # Open the selected Excel "workbook"
    workbook = excel_app.Workbooks.Open(fileImport)


    # For example, print the value of cell A1 in the first worksheet:
    for worksheet in workbook.Worksheets:
        rows = []
        donnee_excel[worksheet.Name] = []
        rowMax, columnMax = get_last_row_and_column(worksheet)

        for row in range(rowMax):
            columns = []

            for col in range(columnMax):
                cell = worksheet.Cells(row+1, col+1)
                columns.append(cell.Value2)
            rows.append(columns)

        if worksheet.Name != "Depart":
            del rows[0]
        donnee_excel[worksheet.Name].extend(rows)

    # Don't forget to clean up when you're done:
    workbook.Close()
    excel_app.Quit()

    return donnee_excel

def save_excel_file(questions):
    """
    Export data created in Revit to in an Excel file
    """
    ctypes.windll.user32.MessageBoxW(0, "S'il vous plait donner chemin et nom pour du fichier Excel", "Plugin Revit", 0)

    # LIRE EXCEL
    # Create Excel application object
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = True

    fileImport = forms.pick_file()

    # Open the selected Excel "workbook"
    workbook = excel_app.Workbooks.Open(fileImport, ReadOnly=False)

    # Verifier s'il existe un fichier avec feuille "Depart"
    Names = []
    for worksheet in workbook.Worksheets:
        Names.append(worksheet.Name)

    if "Depart" in Names:
        ctypes.windll.user32.MessageBoxW(0, "Feuile Depart existe deja", "Plugin Revit", 0)

    else:
        # Créer feuille
        new_sheet = workbook.Worksheets.Add()
        new_sheet.Name = "Depart"

        new_sheet.Cells(1, 1).Value = "Coupe :"
        new_sheet.Cells(1, 2).Value = questions['NC']

        new_sheet.Cells(2, 1).Value = "Arrase SUP :"
        new_sheet.Cells(2, 2).Value = questions['AS']

        new_sheet.Cells(3, 1).Value = "Fiche Mecanique :"
        new_sheet.Cells(3, 2).Value = questions['FM']

        new_sheet.Cells(4, 1).Value = "Enrobage :"
        new_sheet.Cells(4, 2).Value = questions['EN']

        new_sheet.Cells(5, 1).Value = "Recouvrement :"
        for i in range(len(questions['RecouvrementElements'])):
            new_sheet.Cells(5, i+2).Value = questions['RecouvrementElements'][i]

        new_sheet.Cells(6, 1).Value = "Hauteurs Dalles :"
        for i in range(len(questions['HauteurDalles'])):
            new_sheet.Cells(6, i+2).Value = questions['HauteurDalles'][i]

        # Don't forget to clean up when you're done:
        workbook.Save()

    workbook.Close()
    excel_app.Quit()

    return 0

# CODE  -------------------------------------------------------------------------------------------------------
