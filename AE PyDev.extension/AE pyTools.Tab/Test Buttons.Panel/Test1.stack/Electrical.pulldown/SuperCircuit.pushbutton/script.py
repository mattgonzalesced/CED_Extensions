# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT"

import re

from System.Collections.Generic import List
from pyrevit import DB, revit
from pyrevit.revit.db import query

from Snippets._elecutils import (
    get_all_light_devices,
    get_all_panels,
    get_all_elec_fixtures,
    get_all_light_fixtures
)

doc = revit.doc
uidoc = revit.uidoc


# ------------------------------------------------------------
# helpers
# ------------------------------------------------------------

def extract_circuit_number(val):
    """Return first integer found in circuit string for sorting."""
    if not val:
        return 9999
    m = re.search(r"\d+", str(val))
    return int(m.group()) if m else 9999


def group_elements_by_panel_and_circuit(elements, panels):
    """Group elements by (panel name, circuit number)."""

    panel_lookup = {
        query.get_param_value(query.get_param(p, "Panel Name")): p
        for p in panels
    }

    groups = {}

    for el in elements:

        #  skip already-circuited elements
        if is_already_circuited(el):
            continue

        panel_name = query.get_param_value(query.get_param(el, "CKT_Panel_CEDT"))
        circuit_num = query.get_param_value(query.get_param(el, "CKT_Circuit Number_CEDT"))

        if not panel_name or not circuit_num:
            continue

        key = (panel_name, circuit_num)

        if key not in groups:
            groups[key] = {
                "elements": [],
                "element_ids": [],
                "panel_element": panel_lookup.get(panel_name),
                "panel_name": panel_name,
                "circuit_number": circuit_num,
                "rating": query.get_param_value(query.get_param(el, "CKT_Rating_CED")),
                "load_name": query.get_param_value(query.get_param(el, "CKT_Load Name_CEDT")),
                "ckt_notes": query.get_param_value(query.get_param(el, "CKT_Schedule Notes_CEDT"))
            }

        groups[key]["elements"].append(el)
        groups[key]["element_ids"].append(el.Id)

    sorted_keys = sorted(
        groups.keys(),
        key=lambda k: (k[0], extract_circuit_number(k[1]))
    )

    return [(k, groups[k]) for k in sorted_keys]


def create_circuit(doc, element_ids, panel):
    if not element_ids or not panel:
        return None

    id_list = List[DB.ElementId](element_ids)
    system = DB.Electrical.ElectricalSystem.Create(
        doc,
        id_list,
        DB.Electrical.ElectricalSystemType.PowerCircuit
    )

    if system:
        system.SelectPanel(panel)

    return system



def is_already_circuited(element):
    try:
        mep = element.MEPModel
        if not mep:
            return False

        systems = mep.GetElectricalSystems()
        return systems and systems.Count > 0
    except:
        return False


# ------------------------------------------------------------
# main
# ------------------------------------------------------------

def main():
    panels = list(get_all_panels(doc))
    fixtures = (
        list(get_all_elec_fixtures(doc)) +
        list(get_all_light_devices(doc)) +
        list(get_all_light_fixtures(doc))
    )

    selection = revit.get_selection()
    elements = selection if selection else fixtures

    grouped = group_elements_by_panel_and_circuit(elements, panels)

    created_systems = {}

    tg = DB.TransactionGroup(doc, "SUPER CIRCUIT")
    tg.Start()

    # -------------------------
    # TX 1: create circuits
    # -------------------------
    with revit.Transaction("Create Circuits"):
        for key, data in grouped:
            print("Creating circuit: Panel={} | Circuit={}".format(
                data["panel_name"], data["circuit_number"]
            ))

            system = create_circuit(
                doc,
                data["element_ids"],
                data["panel_element"]
            )

            if system:
                created_systems[system.Id] = data
            else:
                print("  -> skipped")

    # -------------------------
    # TX 2: update parameters
    # -------------------------
    with revit.Transaction("Update Circuit Parameters"):
        for sys_id, data in created_systems.items():
            system = doc.GetElement(sys_id)
            if not system:
                continue

            if data["rating"]:
                p = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
                if p:
                    p.Set(data["rating"])

            if data["load_name"]:
                p = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
                if p:
                    p.Set(data["load_name"])

            if data["ckt_notes"]:
                p = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
                if p:
                    p.Set(data["ckt_notes"])

    tg.Assimilate()


if __name__ == "__main__":
    main()
