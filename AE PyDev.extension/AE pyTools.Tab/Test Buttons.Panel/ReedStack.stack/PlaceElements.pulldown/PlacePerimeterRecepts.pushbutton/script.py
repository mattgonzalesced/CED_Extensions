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
from organized.MEPKit.core.log import get_logger, alert, open_output
from organized.MEPKit.electrical.perimeter_runner import place_perimeter_recepts



# --- force-open Output window ---
from pyrevit import script
out = script.get_output()
out.set_title("Perimeter Recepts Log")
out.print_md("## Perimeter Recepts — run log")
try:
    out.center(); out.maximize()
except Exception:
    pass

# --- hard-bind a stdlib logger to this Output panel ---
import logging
class _OutputHandler(logging.Handler):
    def emit(self, record):
        try:
            out.write(self.format(record) + "\n")
        except Exception:
            pass

log = logging.getLogger("PerimeterRecepts")
# clear any old handlers so we don't double print or go to stderr
for h in list(log.handlers):
    log.removeHandler(h)
h = _OutputHandler()
h.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
log.addHandler(h)
log.setLevel(logging.DEBUG)  # show everything


doc = get_doc()

log.info("---- start ----")
count = place_perimeter_recepts(doc, logger=log)
log.info("---- done; placed {} ----".format(count))

alert("Placed {} perimeter receptacle(s).\nCheck the pyRevit Output panel for details."
      .format(count), title="Perimeter Recepts", warn=False)