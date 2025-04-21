# -*- coding: utf-8 -*-

from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import DB
from pyrevit.revit.db import query
import csv


doc = revit.doc
uidoc = revit.uidoc

# Initialize output window
output = script.get_output()


import clr
import csv
import os
from pyrevit import DB, UI
from pyrevit import forms

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Define the FamilyPlacer class for scalability
class FamilyPlacer:
    def __init__(self, doc):
        self.doc = doc

    def place_family_instance(self, family_name, type_name, x, y, z, rotation, level_name, equipment_id_param, equipment_id_value):
        # Get the family symbol
        family_symbol = self.get_family_symbol(family_name, type_name)
        if not family_symbol:
            raise Exception("Family or type not found: {} - {}".format(family_name, type_name))

        # Ensure the family symbol is active
        if not family_symbol.IsActive:
            family_symbol.Activate()
            self.doc.Regenerate()

        # Get the level
        level = self.get_level(level_name)
        if not level:
            raise Exception("Level not found: {}".format(level_name))

        # Create a placement point
        point = DB.XYZ(x, y, z)

        # Start a transaction
        TransactionManager.Instance.EnsureInTransaction(self.doc)

        # Place the family instance
        instance = self.doc.Create.NewFamilyInstance(point, family_symbol, level, DB.Structure.StructuralType.NonStructural)

        # Apply rotation
        if rotation:
            self.rotate_instance(instance, rotation)

        # Set the Equipment ID parameter
        param = instance.LookupParameter(equipment_id_param)
        if param and param.IsReadOnly == False:
            param.Set(equipment_id_value)

        TransactionManager.Instance.TransactionTaskDone()
        return instance

    def get_family_symbol(self, family_name, type_name):
        collector = DB.FilteredElementCollector(self.doc)
        collector.OfClass(DB.FamilySymbol)
        for symbol in collector:
            if symbol.FamilyName == family_name and symbol.Name == type_name:
                return symbol
        return None

    def get_level(self, level_name):
        collector = DB.FilteredElementCollector(self.doc)
        collector.OfClass(DB.Level)
        for level in collector:
            if level.Name == level_name:
                return level
        return None

    def rotate_instance(self, instance, angle):
        location = instance.Location
        if isinstance(location, DB.LocationPoint):
            point = location.Point
            axis = DB.Line.CreateBound(point, point.Add(DB.XYZ(0, 0, 1)))
            angle_rad = DB.UnitUtils.ConvertToInternalUnits(angle, DB.DisplayUnitType.DUT_DEGREES)
            DB.ElementTransformUtils.RotateElement(self.doc, instance.Id, axis, angle_rad)

# CSV Handler class
class CSVHandler:
    @staticmethod
    def read_csv(file_path):
        data = []
        with open(file_path, 'r') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                data.append(row)
        return data

# Main script execution
def main():
    doc = DocumentManager.Instance.CurrentDBDocument

    # Get the CSV file path
    file_path = forms.pick_file(file_ext='csv', init_dir=os.getcwd())
    if not file_path:
        forms.alert("No file selected. Operation cancelled.", title="Error")
        return

    # Read the CSV data
    csv_data = CSVHandler.read_csv(file_path)

    # Initialize the FamilyPlacer
    family_placer = FamilyPlacer(doc)

    # Process each row in the CSV
    for row in csv_data:
        try:
            family_placer.place_family_instance(
                family_name=row['family'],
                type_name=row['type_name'],
                x=float(row['x']),
                y=float(row['y']),
                z=float(row['z']),
                rotation=float(row['rotation']),
                level_name=row['level'],
                equipment_id_param="Equipment ID_CEDT",
                equipment_id_value=row['id']
            )
        except Exception as e:
            forms.alert("Error placing family: {}\n{}".format(row, str(e)), title="Error")

# Run the script
if __name__ == "__main__":
    main()
