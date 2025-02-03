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
    def ensure_parameter_exists(doc, param_name, cat_set):
        with Transaction(doc, f"Ensure Parameter {param_name}") as t:
            t.Start()

            app = doc.Application

            # Open the shared parameter file
            def_file = app.OpenSharedParameterFile()
            if not def_file:
                TaskDialog.Show("Error", "No shared parameter file is defined. Please set one in Revit Options.")
                t.RollBack()
                return

            # Get or create the group "ManriqueBimTools"
            group = def_file.Groups.Item["ManriqueBimTools"] or def_file.Groups.Create("ManriqueBimTools")

            # Check if the parameter exists in the shared parameter group
            definition = None
            for defn in group.Definitions:
                if defn.Name.lower() == param_name.lower():
                    definition = defn
                    break

            # If no existing definition was found, create a new one
            if definition is None:
                options = ExternalDefinitionCreationOptions(param_name, "Autodesk.Revit.DB.SpecTypeId.String.Text")
                options.Visible = True
                definition = group.Definitions.Create(options)

            # Get existing parameter bindings
            param_bindings = doc.ParameterBindings
            existing_binding = None
            it = param_bindings.ForwardIterator()

            while it.MoveNext():
                defn = it.Key
                if defn and defn.Name.lower() == param_name.lower():
                    existing_binding = it.Current
                    break

            if existing_binding:
                existing_cat_set = existing_binding.Categories
                updated = False

                # Add missing categories to the existing binding
                for cat in cat_set:
                    if not existing_cat_set.Contains(cat):
                        existing_cat_set.Insert(cat)
                        updated = True

                if updated:
                    doc.ParameterBindings.ReInsert(definition, existing_binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            else:
                # Create a new instance binding
                new_binding = app.Create.NewInstanceBinding(cat_set)
                success = doc.ParameterBindings.Insert(definition, new_binding, BuiltInParameterGroup.PG_IDENTITY_DATA)

                if not success:
                    TaskDialog.Show("Error", f"Could not bind parameter: {param_name}")

            t.Commit()

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
