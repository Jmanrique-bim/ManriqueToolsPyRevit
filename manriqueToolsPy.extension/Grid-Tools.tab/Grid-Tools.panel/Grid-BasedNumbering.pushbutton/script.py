# -*- coding: utf-8 -*-
"""
pyRevit script: Locate Instance by Grid
Numbers family instances based on grid intersections using a user‐selected first instance
and one of three ordering methods: by X axis, Y axis, or proximity.
"""

import clr
import sys
from System import Double
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms

# ------------------------------
# GridHelper class definition
# ------------------------------

class GridHelper(object):
    @staticmethod
    def IsAlphabetic(input_str):
        if not input_str:
            return False
        return input_str[0].isalpha()

    @staticmethod
    def IsNumeric(input_str):
        if not input_str:
            return False
        return input_str[0].isdigit()

    @staticmethod
    def GetElementLocation(e):
        if e is None:
            return None
        loc = e.Location
        if isinstance(loc, LocationPoint):
            return loc.Point
        elif isinstance(loc, LocationCurve):
            curve = loc.Curve
            return curve.Evaluate(0.5, True)
        return None

    @staticmethod
    def EnsureParameterExists(doc, paramName, catSet):
        """
        Ensures that a shared text parameter with name paramName exists and is bound to the categories in catSet.
        If already bound, does nothing.
        """
        bindings = doc.ParameterBindings
        iter_binding = bindings.ForwardIterator()
        while iter_binding.MoveNext():
            definition = iter_binding.Key
            if definition and definition.Name.lower() == paramName.lower():
                # Parameter already exists; no need to add.
                return

        app = doc.Application
        defFile = app.OpenSharedParameterFile()
        if defFile is None:
            TaskDialog.Show("Error", "No shared parameter file is defined. Please set one in Revit Options.")
            return

        group = defFile.Groups.get_Item("ManriqueBimTools")
        if group is None:
            group = defFile.Groups.Create("ManriqueBimTools")

        # Check if a definition with the same name already exists in the group.
        definition = None
        for defn in group.Definitions:
            if defn.Name.lower() == paramName.lower():
                definition = defn
                break

        if definition is None:
            # Note: In Revit 2023, use SpecTypeId.String for a text parameter.
            options = ExternalDefinitionCreationOptions(paramName, SpecTypeId.String)
            options.Visible = True
            definition = group.Definitions.Create(options)

        binding = app.Create.NewInstanceBinding(catSet)
        success = bindings.Insert(definition, binding)
        if success:
            bindings.ReInsert(definition, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
        else:
            TaskDialog.Show("Error", "Could not bind parameter: " + paramName)

# ------------------------------
# Main command logic
# ------------------------------

# Get the current UIDocument and Document from pyRevit's __revit__ variable.
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = __revit__.Application

# STEP 1: Prompt the user to select the first instance.
try:
    firstRef = uidoc.Selection.PickObject(ObjectType.Element, "Select the first instance for numbering")
except Exception as ex:
    sys.exit("Operation cancelled.")

firstElem = doc.GetElement(firstRef)
if not firstElem:
    TaskDialog.Show("Error", "Could not retrieve the selected element.")
    sys.exit("Failed to get element.")

# STEP 2: Show a simple options form (using pyRevit FlexForm) to choose ordering.
# The form includes three checkboxes.
components = [
    forms.CheckBox("Number on X Axis", default=False),
    forms.CheckBox("Number on Y Axis", default=False),
    forms.CheckBox("Number by Proximity", default=True)
]
title = "Grid Numbering Options"
data = forms.FlexForm(title, components).show()
if not data:
    sys.exit("User cancelled the options.")
numberOnXAxis = data["Number on X Axis"]
numberOnYAxis = data["Number on Y Axis"]
numberByProximity = data["Number by Proximity"]

# STEP 3: Get the category of the first element.
firstCat = firstElem.Category
if firstCat is None:
    TaskDialog.Show("Error", "The selected element does not belong to any category.")
    sys.exit("Failed to get category.")

# STEP 4: Collect all family instances in the document that belong to the same category.
collector = FilteredElementCollector(doc).OfCategoryId(firstCat.Id).OfClass(FamilyInstance).WhereElementIsNotElementType()
familyInstances = [fi for fi in collector if isinstance(fi, FamilyInstance) and fi.SuperComponent is None]
selectedElements = list(familyInstances)

# STEP 5: Create a CategorySet for binding shared parameters.
catSet = uidoc.Application.Create.NewCategorySet()
catSet.Insert(firstCat)

# Ensure that the required shared parameters exist.
GridHelper.EnsureParameterExists(doc, "Grid Square", catSet)
GridHelper.EnsureParameterExists(doc, "Number", catSet)

# STEP 6: Collect all grid elements.
gridCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType()
grids = list(gridCollector)

letterGrids = [g for g in grids if GridHelper.IsAlphabetic(g.Name)]
numberGrids = [g for g in grids if GridHelper.IsNumeric(g.Name)]

# STEP 7: Determine the starting point from the first element.
startPt = GridHelper.GetElementLocation(firstElem)
if not startPt:
    TaskDialog.Show("Error", "Could not determine the location of the first instance.")
    sys.exit("Failed to get location.")

# STEP 8: Sort the family instances according to the chosen method,
# but always force the first element (the one selected by the user) to be at the beginning.
others = [e for e in selectedElements if e.Id != firstElem.Id]
if numberOnXAxis:
    sortedOthers = sorted(others, key=lambda e: (GridHelper.GetElementLocation(e).X if GridHelper.GetElementLocation(e) is not None else 0))
elif numberOnYAxis:
    sortedOthers = sorted(others, key=lambda e: (GridHelper.GetElementLocation(e).Y if GridHelper.GetElementLocation(e) is not None else 0))
elif numberByProximity:
    sortedOthers = sorted(others, key=lambda e: (GridHelper.GetElementLocation(e).DistanceTo(startPt) if GridHelper.GetElementLocation(e) is not None else float('inf')))
else:
    # Default to proximity ordering if no option is selected.
    sortedOthers = sorted(others, key=lambda e: (GridHelper.GetElementLocation(e).DistanceTo(startPt) if GridHelper.GetElementLocation(e) is not None else float('inf')))

sortedElements = [firstElem] + sortedOthers

# STEP 9: Start a transaction to update the parameters on each element.
t = Transaction(doc, "Assign Grid Square and Number")
t.Start()
counter = 1
for elem in sortedElements:
    elemPt = GridHelper.GetElementLocation(elem)
    if elemPt is None:
        continue

    # Find the closest letter grid and number grid.
    closestLetter = None
    closestNumber = None
    if letterGrids:
        sortedLetter = sorted(letterGrids, key=lambda g: g.Curve.Distance(elemPt))
        closestLetter = sortedLetter[0]
    if numberGrids:
        sortedNumber = sorted(numberGrids, key=lambda g: g.Curve.Distance(elemPt))
        closestNumber = sortedNumber[0]

    gridSquare = ""
    if closestLetter and closestNumber:
        gridSquare = "{}-{}".format(closestLetter.Name, closestNumber.Name)

    gridParam = elem.LookupParameter("Grid Square")
    if gridParam and not gridParam.IsReadOnly:
        gridParam.Set(gridSquare)

    numParam = elem.LookupParameter("Number")
    if numParam and not numParam.IsReadOnly:
        numParam.Set(str(counter))
    counter += 1
t.Commit()

TaskDialog.Show("Success", "Elements numbered successfully.")
