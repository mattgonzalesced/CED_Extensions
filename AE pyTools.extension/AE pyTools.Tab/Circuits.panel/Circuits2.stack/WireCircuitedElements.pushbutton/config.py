from pyrevit import script, forms, DB
from Autodesk.Revit.DB import FilteredElementCollector, Electrical

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Initialize config to save default wire type
config = script.get_config("wire_type_config")

# Retrieve all wire types in the document
wire_types = FilteredElementCollector(doc).OfClass(Electrical.WireType).ToElements()
wire_type_options = {wire_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(): wire_type.Id
                     for wire_type in wire_types}

# Prompt user to select a default wire type
selected_wire_type_name = forms.SelectFromList.show(
    sorted(wire_type_options.keys()),
    title="Select Default Wire Type",
    button_name="Save as Default"
)

# Save selected wire type to config
if selected_wire_type_name:
    config.default_wire_type = selected_wire_type_name
    script.save_config()
    script.get_logger().info("Saved " + selected_wire_type_name + " as the default wire type.")
else:
    script.get_logger().warning("No wire type selected. Default not saved.")
