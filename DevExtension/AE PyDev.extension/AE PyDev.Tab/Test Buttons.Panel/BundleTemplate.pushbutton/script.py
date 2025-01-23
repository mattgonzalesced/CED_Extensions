# -*- coding: utf-8 -*-


import clr
from pyrevit import script
from pyrevit import revit, DB
from pyrevit.revit import query
from collections import defaultdict

# Access the current Revit document
doc = revit.doc

# Set up the output window
output = script.get_output()
output.close_others()
output.set_width(800)
