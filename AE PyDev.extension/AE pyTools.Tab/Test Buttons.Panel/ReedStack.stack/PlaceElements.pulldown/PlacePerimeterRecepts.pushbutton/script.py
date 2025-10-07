# RP button: Place Perimeter Recepts.pushbutton/script.py
from organized.MEPKit.revit.appdoc import get_doc
from organized.MEPKit.core.logging import get_logger
from organized.MEPKit.electrical.perimeter_runner import place_perimeter_recepts

doc = get_doc()
log = get_logger("PerimeterRecepts", level="INFO")
count = place_perimeter_recepts(doc, logger=log)
log.info("Done. Placed: {}".format(count))