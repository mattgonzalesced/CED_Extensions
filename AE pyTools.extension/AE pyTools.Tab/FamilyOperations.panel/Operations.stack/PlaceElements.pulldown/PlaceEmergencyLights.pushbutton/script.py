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
from organized.MEPKit.tools.place_emergency_lighting import place_emergency_lighting

log = get_logger("Emergency Lighting", level="DEBUG", title="Emergency Lighting Log")



doc = get_doc()

log.info("---- start ----")
count = place_emergency_lighting(doc, logger=log)
log.info("---- done; placed {} ----".format(count))

alert("Placed Emergency lighting.\nCheck the pyRevit Output panel for details.")