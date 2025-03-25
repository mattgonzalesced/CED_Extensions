# -*- coding: utf-8 -*-

from pyrevit import DB, script, forms, revit, output
from pyrevit.revit import query
import clr
from System.Collections.Generic import List
from Snippets import _elecutils as eu
from Autodesk.Revit.DB.Electrical import *

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()

def pick_circuits_from_list():

    ckts = DB.FilteredElementCollector(doc)\
        .OfClass(ElectricalSystem)\
        .WhereElementIsNotElementType()

    print("Total Circuits in Doc: {}".format(ckts.GetElementCount()))

    ckt_options = {" All": []}

    for ckt in ckts:
        ckt_supply = ckt.BaseEquipment.Name
        ckt_number = ckt.CircuitNumber
        ckt_load_name = ckt.LoadName
        ckt_rating = ckt.Rating
        ckt_wiretype = ckt.WireType

        ckt_options[" All"].append(ckt)
        print("{}/{} ({}) - {}".format(ckt_supply, ckt_number, ckt_rating, ckt_load_name))
        if ckt_supply not in ckt_options:
            ckt_options[ckt_supply] = []
        ckt_options[ckt_supply].append(ckt)

    grouped_options = {}
    for group, ckts in ckt_options.items():
        grouped_options[group] = [
            "{} | {} - {}".format(ckt_supply, ckt_number, ckt_load_name) for ckt in ckts
        ]
        grouped_options[group].sort()

        # Show selection dialog
    selected_option = forms.SelectFromList.show(
        grouped_options,
        title="Select a CKT",
        group_selector_title="Panel:",
        multiselect=False
    )

    #print("{}/{} ({}) - {}".format(ckt_supply,ckt_number,ckt_rating,ckt_load_name))
    print(selected_option)




pick_circuits_from_list()