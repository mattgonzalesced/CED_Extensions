# -*- coding: utf-8 -*-
"""
pyRevit script
--------------
Compute multiline panel/circuit text for each EF‑U_Junction Box_CED‑HEB
instance whose type name is “REFRIGERATION PLAN” and write the string into:

        • Tag_Text   (instance parameter)
        • CED‑E‑FIXTURE TYPE   (instance parameter; falls back to type parameter)

The write style matches the sample that already works for you: Windows CR+LF
line breaks and a single commit‑once transaction.
"""

from collections import defaultdict
import re
from pyrevit import revit, DB, script

doc    = revit.doc
logger = script.get_logger()

# ----------------------------- helpers ---------------------------------- #

def get_classification(idx, has_defrost=True):
    """First system for a panel -> FAN+LTS+WMR (or FAN+LTS if no defrost);
       all others -> DEFROST."""
    if idx == 1:
        return "FAN+LTS+WMR" if has_defrost else "FAN+LTS"
    return "DEFROST"


def _numeric_key(circ):
    """Return a tuple key for stable sorting: (is_multipole, numeric if present else large, circ)."""
    if not circ:
        return (1, 10**12, "")

    # Detect multi-pole circuit (contains comma)
    is_multipole = 1 if ',' in circ else 0

    m = re.search(r'\d+', circ)
    if m:
        try:
            return (is_multipole, int(m.group()), circ)
        except Exception:
            pass
    return (1, 10**12, circ)


def build_multiline_text(fixt):
    """Return the CR+LF multiline string for one fixture."""
    mep = getattr(fixt, 'MEPModel', None)
    if not mep:
        return None

    elecs = list(mep.GetAssignedElectricalSystems())
    if not elecs:
        elecs = [s for s in mep.GetElectricalSystems()
                 if s.GetType().Name == 'ElectricalSystem']
    if not elecs:
        return None

    # collect circuits by panel, preserve panel encounter order
    panel_circs = defaultdict(list)
    panels_order = []
    for sys in elecs:
        panel = getattr(sys, 'PanelName', '') or ''
        circ  = sys.CircuitNumber or ''
        if panel not in panel_circs:
            panels_order.append(panel)
        panel_circs[panel].append(circ)

    groups, order = {}, []
    # For each panel choose a deterministic "primary" circuit (lowest numeric then lexical)
    for panel in panels_order:
        circs = panel_circs[panel]
        # sort deterministically
        sorted_circs = sorted(circs, key=_numeric_key)
        # primary (first) -> FAN+LTS+WMR (or FAN+LTS if no DEFROST)
        primary = sorted_circs[0]
        remaining = sorted_circs[1:]
        has_defrost = bool(remaining)
        groups[(panel, get_classification(1, has_defrost))] = [primary]
        order.append((panel, get_classification(1, has_defrost)))
        # remaining -> DEFROST (if any)
        if remaining:
            groups[(panel, get_classification(2))] = remaining
            order.append((panel, get_classification(2)))

    lines = []
    for panel, clsf in order:
        circs = groups[(panel, clsf)]
        parts = [('({})'.format(c) if ',' in c else c) for c in circs]
        lines.append('{} - {}'.format(panel, ', '.join(parts)))
        lines.append(clsf)

    # Windows‑style breaks
    return '\r\n'.join(lines)


def set_param(element, param_name, value):
    """Write value into an instance parameter; fall back to the type."""
    p = element.LookupParameter(param_name)
    if p and not p.IsReadOnly:
        p.Set(value)
        return True

    # fallback: type parameter
    try:
        sym = element.Symbol
        p   = sym.LookupParameter(param_name)
        if p and not p.IsReadOnly:
            p.Set(value)
            return True
    except Exception:
        pass

    logger.warning(
        "Cannot write '{0}' on element id {1} (instance or type)".format(
            param_name, element.Id.IntegerValue))
    return False


# ----------------------------- main routine ----------------------------- #

# Collect electrical‑fixture **instances** (not types)
collector = DB.FilteredElementCollector(doc) \
    .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures) \
    .WhereElementIsNotElementType()

fixtures = []
for inst in collector:
    # filter by type name == "REFRIGERATION PLAN"
    type_el    = doc.GetElement(inst.GetTypeId())
    type_name  = None
    if type_el:
        type_param = type_el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        type_name  = type_param.AsString() if type_param else None
    if type_name == "REFRIGERATION PLAN":
        fixtures.append(inst)

if not fixtures:
    logger.info("No 'REFRIGERATION PLAN' fixtures found.")
    script.exit()

logger.info("Found {} 'REFRIGERATION PLAN' fixture(s).".format(len(fixtures)))

with DB.Transaction(doc, "Write multiline Tag_Text + Fixture Type") as tx:
    tx.Start()
    updated_tag = 0
    updated_fix = 0

    for inst in fixtures:
        text_val = build_multiline_text(inst)
        if not text_val:
            logger.warning("Host {} has no electrical systems; skipped."
                           .format(inst.Id.IntegerValue))
            continue

        # Tag_Text
        if set_param(inst, "Tag_Text", text_val):
            updated_tag += 1

        # CED‑E‑FIXTURE TYPE
        # if set_param(inst, "CED-E-FIXTURE TYPE", text_val):
        #     updated_fix += 1

    tx.Commit()

logger.info("Updated Tag_Text on {} fixture(s).".format(updated_tag))
# logger.info("Updated CED‑E‑FIXTURE TYPE on {} fixture(s).".format(updated_fix))
