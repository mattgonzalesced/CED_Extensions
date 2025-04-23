# -*- coding: utf-8 -*-
import os
import re
from pyrevit import forms, script
from pyrevit import revit, DB
from pyrevit import UI
from pyrevit import coreutils

doc = revit.doc
active_view = doc.ActiveView

all_grids = DB.FilteredElementCollector(revit.doc) \
    .OfCategory(DB.BuiltInCategory.OST_Grids) \
    .WhereElementIsNotElementType().ToElements()

selected_option = \
    forms.CommandSwitchWindow.show(
        ['Show All',
         'Show End 1',
         'Show End 2',
         'Hide All'],
        message='Select Grid Bubble Option:'
    )

grids = []
selection = revit.get_selection()

if selection:
    grids = [x for x in selection if isinstance(x, DB.Grid)]
else:
    grids = all_grids

try:
    with revit.Transaction('Toggle Grid Bubbles'):
        for grid in grids:
            if selected_option == 'Show All':
                grid.ShowBubbleInView(DB.DatumEnds.End0, active_view)
                grid.ShowBubbleInView(DB.DatumEnds.End1, active_view)
            elif selected_option == 'Hide All':
                grid.HideBubbleInView(DB.DatumEnds.End0, active_view)
                grid.HideBubbleInView(DB.DatumEnds.End1, active_view)
            elif selected_option == 'Show End 1':
                # only show the “start” end, hide the other
                grid.ShowBubbleInView(DB.DatumEnds.End0, active_view)
                grid.HideBubbleInView(DB.DatumEnds.End1, active_view)
            elif selected_option == 'Show End 2':
                # only show the “finish” end, hide the other
                grid.HideBubbleInView(DB.DatumEnds.End0, active_view)
                grid.ShowBubbleInView(DB.DatumEnds.End1, active_view)

except Exception:
    pass

revit.uidoc.RefreshActiveView()
