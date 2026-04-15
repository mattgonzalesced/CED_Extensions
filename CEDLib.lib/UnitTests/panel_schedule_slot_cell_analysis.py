# -*- coding: utf-8 -*-
"""Unit-test utility: compare direct slot cells from Revit API vs PanelScheduleManager."""

import Autodesk.Revit.DB.Electrical as DBE
import clr
from System.Collections.Generic import IList, List
from pyrevit import DB, script
from pyrevit.compat import get_elementid_value_func

from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository as ps_repo
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager


def _direct_cells_by_slot(schedule_view, slot):
    """Return direct GetCellsBySlotNumber coordinates and the successful binding mode."""
    getter = getattr(schedule_view, "GetCellsBySlotNumber", None)
    if getter is None:
        return [], "missing"

    # Pattern A: out-ref
    try:
        row_ref = clr.Reference[IList[int]]()
        col_ref = clr.Reference[IList[int]]()
        getter(int(slot), row_ref, col_ref)
        row_arr = list(row_ref.Value or [])
        col_arr = list(col_ref.Value or [])
        pairs = []
        for idx in range(int(min(len(row_arr), len(col_arr)))):
            pairs.append((int(row_arr[idx]), int(col_arr[idx])))
        if pairs:
            return sorted(set(pairs)), "out-ref"
    except Exception:
        pass

    # Pattern B: tuple-return
    try:
        raw = getter(int(slot))
        if isinstance(raw, tuple) and len(raw) >= 2:
            row_arr = list(raw[0] or [])
            col_arr = list(raw[1] or [])
            pairs = []
            for idx in range(int(min(len(row_arr), len(col_arr)))):
                pairs.append((int(row_arr[idx]), int(col_arr[idx])))
            if pairs:
                return sorted(set(pairs)), "tuple"
    except Exception:
        pass

    # Pattern C: preallocated lists
    try:
        row_arr = List[int]()
        col_arr = List[int]()
        getter(int(slot), row_arr, col_arr)
        pairs = []
        for idx in range(int(min(len(row_arr), len(col_arr)))):
            pairs.append((int(row_arr[idx]), int(col_arr[idx])))
        if pairs:
            return sorted(set(pairs)), "prealloc"
    except Exception:
        pass

    return [], "none"


def run(doc=None, output=None, logger=None):
    """Print comparison tables for all non-template panel schedule views."""
    if doc is None:
        doc = __revit__.ActiveUIDocument.Document
    if output is None:
        output = script.get_output()
    if logger is None:
        logger = script.get_logger()

    get_id_val = get_elementid_value_func()
    options = list(ps_repo.collect_panel_equipment_options(doc, include_without_schedule=True) or [])
    schedule_to_option = {}
    panel_option_lookup = {}
    for option in options:
        panel_id = int(option.get("panel_id", 0) or 0)
        schedule_view = option.get("schedule_view")
        if panel_id > 0:
            panel_option_lookup[panel_id] = option
        if schedule_view is None:
            continue
        sid = int(get_id_val(getattr(schedule_view, "Id", None)))
        if sid > 0:
            schedule_to_option[sid] = option

    psm = PanelScheduleManager(doc, panel_option_lookup=panel_option_lookup, logger=logger)
    views = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DBE.PanelScheduleView)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    views = sorted(
        [v for v in views if not bool(getattr(v, "IsTemplate", False))],
        key=lambda x: str(getattr(x, "Name", "")),
    )
    if not views:
        output.print_md("**No panel schedule views found.**")
        return

    output.print_md("## Slot Cell Analysis")
    output.print_md("Direct API: `PanelScheduleView.GetCellsBySlotNumber` vs `PanelScheduleManager.get_row_col_from_slot`")
    for view in views:
        view_id = int(get_id_val(getattr(view, "Id", None)))
        option = schedule_to_option.get(view_id)
        panel_id = int(option.get("panel_id", 0) or 0) if option else 0
        panel_name = str(getattr(option.get("panel"), "Name", "")) if option else ""
        try:
            table = view.GetTableData()
            max_slot = int(getattr(table, "NumberOfSlots", 0) or 0)
        except Exception:
            max_slot = 0
        if max_slot <= 0:
            continue

        rows = []
        mismatch_count = 0
        binding_modes = set()
        for slot in range(1, max_slot + 1):
            direct_cells, mode = _direct_cells_by_slot(view, int(slot))
            binding_modes.add(str(mode))

            psm_cells = []
            if panel_id > 0:
                try:
                    psm_cells = list(psm.get_row_col_from_slot(panel_id, int(slot)) or [])
                except Exception:
                    psm_cells = []
            psm_cells = sorted(set([(int(r), int(c)) for r, c in list(psm_cells or [])]))

            matches = bool(direct_cells == psm_cells)
            if not matches:
                mismatch_count += 1
            rows.append([
                int(slot),
                int(len(direct_cells)),
                ", ".join(["({0},{1})".format(r, c) for r, c in direct_cells]),
                int(len(psm_cells)),
                ", ".join(["({0},{1})".format(r, c) for r, c in psm_cells]),
                "Yes" if matches else "No",
            ])

        header = "### {0} | Panel: {1} | Slots: {2} | Mismatches: {3}".format(
            str(getattr(view, "Name", "")),
            panel_name if panel_name else "(unmapped)",
            int(max_slot),
            int(mismatch_count),
        )
        output.print_md(header)
        output.print_md("Binding mode(s): {0}".format(", ".join(sorted(list(binding_modes)))))
        output.print_table(rows, ["Slot", "Direct Cnt", "Direct Cells", "PSM Cnt", "PSM Cells", "Match"])


if __name__ == "__main__":
    run()
