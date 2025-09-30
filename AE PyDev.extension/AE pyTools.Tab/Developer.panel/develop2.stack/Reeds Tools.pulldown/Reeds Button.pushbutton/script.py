# -*- coding: utf-8 -*-
from pyrevit import script, revit, DB

#these are my typical imports. I add more as necessary. Rather than importing all of Autodesk.DB, i prefer using
#pyrevits DB module, and begin any revit API method with the module. ex: DB.FilteredElementCollector
#pyrevit doesnt have Autodesk.Revit.DB.Electrical, so I import that as DBE to mimic the same convention
#this helps you keep track of the modules youre using. doing a * import can get messy



# Setup
uidoc = revit.uidoc
doc = revit.doc
app = doc.Application

logger = script.get_logger()
logger.debug("Use These within your code to test conditions and print values/variables."
                "\n example:"
                "\n App Version: {}"
                "\n Doc Title: {}".format(app.VersionName, doc.Title))



#this sets up any output window that you want to show. it has functions to control its appearance,
output = script.get_output()
output.close_others()   #prevents windows from building up as you run scripts




#SELECTION
#you can use Revit API selection method
selection = uidoc.Selection
#or use pyrevits wrapper which provides some more functionality
pyrevit_selection = revit.get_selection()





#EXAMPLES
#--------------------------------------------------
#🟠 Read Parameter Properties
def get_param_value(param):
    """Get a value from a Parameter based on its StorageType."""
    value = None
    if param.StorageType == DB.StorageType.Double:      value = param.AsDouble()
    elif param.StorageType == DB.StorageType.ElementId: value = param.AsElementId()
    elif param.StorageType == DB.StorageType.Integer:   value = param.AsInteger()
    elif param.StorageType == DB.StorageType.String:    value = param.AsString()
    return value

# Read All Instance Parameters of an Element
for p in picked_object.Parameters:
    print("Name: {}".format(p.Definition.Name))
    print("ParameterGroup: {}".format(p.Definition.ParameterGroup))
    print("BuiltInParameter: {}".format(p.Definition.BuiltInParameter))
    print("IsReadOnly: {}".format(p.IsReadOnly))
    print("HasValue: {}".format(p.HasValue))
    print("IsShared: {}".format(p.IsShared))
    print("StorageType: {}".format(p.StorageType))
    print("Value: {}".format(get_param_value(p)))
    print("AsValueString(): {}".format(p.AsValueString()))
    print('-'*100)