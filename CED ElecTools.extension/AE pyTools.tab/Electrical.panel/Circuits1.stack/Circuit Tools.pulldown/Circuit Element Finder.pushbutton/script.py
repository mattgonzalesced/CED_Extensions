# -*- coding: utf-8 -*-

import os

from pyrevit import forms, revit, script

from UIClasses import pathing as ui_pathing

TITLE = "Circuited Device Finder"

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT:
    forms.alert("Could not locate CEDLib.lib.", title=TITLE, exitscript=True)

from CEDElectrical.ui.circuit_element_finder_action import run_circuit_element_finder


def main():
    logger = script.get_logger()
    result = run_circuit_element_finder(uidoc=revit.uidoc, logger=logger) or {}
    if result.get("status") != "ok":
        logger.debug("Circuited Device Finder ended: {0}".format(result))


main()
