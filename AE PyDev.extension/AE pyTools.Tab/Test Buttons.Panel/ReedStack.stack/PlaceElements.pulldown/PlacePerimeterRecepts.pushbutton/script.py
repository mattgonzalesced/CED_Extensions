# -*- coding: utf-8 -*-
import os, sys
_here = os.path.dirname(__file__)
while _here and not os.path.isdir(os.path.join(_here, 'lib')):
    _here = os.path.dirname(_here)
org = os.path.join(_here, 'lib', 'organized')
if org not in sys.path:
    sys.path.insert(0, org)

# RP button: Place Perimeter Recepts.pushbutton/script.py
from organized.MEPKit.revit.appdoc import get_doc
from organized.MEPKit.core.log import get_logger, alert
from organized.MEPKit.electrical.perimeter_runner import place_perimeter_recepts

open_output("Perimeter Recepts Log", header_md="## Perimeter Recepts — run log")

doc = get_doc()
log = get_logger("PerimeterRecepts", level="INFO")

log.info("---- start ----")
count = place_perimeter_recepts(doc, logger=log)
log.info("---- done; placed {} ----".format(count))

alert("Placed {} perimeter receptacle(s).\nCheck the pyRevit Output panel for details."
      .format(count), title="Perimeter Recepts", warn=False)