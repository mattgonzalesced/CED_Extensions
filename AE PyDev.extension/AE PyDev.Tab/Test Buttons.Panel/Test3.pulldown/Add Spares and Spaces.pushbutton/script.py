# -*- coding: utf-8 -*-
from pyrevit import revit, UI, DB
from Autodesk.Revit.DB import Electrical as DBE
from pyrevit import forms
from pyrevit import script
from pyrevit.revit import query
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
import re
from collections import defaultdict

# Import reusable utilities
from Snippets._elecutils import get_panel_dist_system, get_compatible_panels

# Get the current document

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
console = script.get_output()


def get_view_schedules_from_selection(elements):
    valid_schedules = []
    discard_counter = defaultdict(int)

    for el in elements:
        if isinstance(el, DBE.PanelScheduleSheetInstance):
            vs = doc.GetElement(el.ScheduleId)   # <- this is a PanelScheduleView
            # accept ANY view or be specific:
            # if isinstance(vs, DBE.PanelScheduleView):
            if isinstance(vs, DB.View):
                valid_schedules.append(vs)
        else:
            discard_counter[el.Category.Name if el.Category else "Unknown"] += 1

    for cat, count in discard_counter.items():
        print(u"⚠️  {} {} element(s) discarded.".format(count, cat))

    return valid_schedules


def prompt_user_for_schedules(doc):
    """
    Prompt the user to select one or more panel schedules from the project.
    Returns a list of PanelScheduleView objects.
    Exits if no schedules available or nothing selected.
    """
    # 1. Collect all schedules
    all_schedules = [
        ps for ps in DB.FilteredElementCollector(doc)
        .OfClass(DBE.PanelScheduleView)
        if not ps.IsTemplate
    ]
    if not all_schedules:
        forms.alert("No panel schedules found in the project.", exitscript=True)

    # 2. Wrap each schedule in an object that shows the schedule name in the UI
    class ScheduleOption(object):
        def __init__(self, schedule):
            self.schedule = schedule
        def __str__(self):
            # The text displayed to the user
            return self.schedule.Name

    schedule_options = [ScheduleOption(s) for s in all_schedules]

    # 3. Multi-select from list
    selected = forms.SelectFromList.show(
        schedule_options,
        title="Select Panel Schedule(s) to Modify",
        multiselect=True
    )
    if not selected:
        forms.alert("No schedules selected. Exiting.", exitscript=True)

    # 4. Return the chosen schedule objects
    return [s.schedule for s in selected]

def collect_schedules_to_process():
    schedules_to_process = []

    # Case 1: If active view is a ViewSchedule, use it
    active_view = uidoc.ActiveView
    if isinstance(active_view, DBE.PanelScheduleView):
        print("🔹 Active view is a schedule: {}".format(active_view.Name))
        schedules_to_process.append(active_view)

    else:
        # Case 2: Check selected elements
        selection = revit.get_selection()
        if selection:
            print("🔹 Evaluating selected elements...")
            schedules_to_process.extend(get_view_schedules_from_selection(selection.elements))

        # Case 3: Nothing selected + not in schedule view → prompt user
        if not schedules_to_process:
            print("🔹 No valid schedules selected. Prompting user to choose from list...")
            schedules_to_process = prompt_user_for_schedules(doc)

    # Process all schedules at once
    # clean_schedule_headers(schedules_to_process)

    # print("\n✅ Done. {} schedule(s) processed.".format(len(schedules_to_process)))
    return schedules_to_process


def get_schedule_info(schedules):
    for schedule in schedules:
        try:
            table_data = schedule.GetTableData()
            body_section = table_data.GetSectionData(DB.SectionType.Body)
            total_slots = table_data.NumberOfSlots
            if body_section:
                num_rows = body_section.NumberOfRows
                num_cols = body_section.NumberOfColumns

                print("\nSchedule Name: {0}".format(schedule.Name))
                print("Body Section: Rows = {0}, Cols = {1}".format(num_rows, num_cols))

                for row in range(num_rows):
                    for col in range(num_cols):
                        # You can retrieve cell text or type if needed
                        cell_ckt = schedule.GetCircuitIdByCell(row, col)
                        cell_slot = schedule.GetSlotNumberByCell(row, col)
                        # Do whatever you want with the text or type
                        print("Row: {0:<3} Col: {1:<3} | cell_ckt: {2} | cell_slot: {3}".format(
                            row, col, cell_ckt, cell_slot
                        ))
            else:
                print("⚠ No body section found in schedule: {0}".format(schedule.Name))

        except Exception as e:
            print("⚠ Error processing schedule '{0}': {1}".format(schedule.Name, e))


import clr
from Autodesk.Revit.DB import Transaction, ElementId

def fill_half_spare_half_space(schedules, doc):
    with Transaction(doc, "Fill Spare and Space") as t:
        t.Start()

        for schedule in schedules:
            try:
                table_data = schedule.GetTableData()
                body_section = table_data.GetSectionData(DB.SectionType.Body)  # or DB.SectionType.Body
                if not body_section:
                    print("No body section found in schedule: {0}".format(schedule.Name))
                    continue

                total_slots = table_data.NumberOfSlots
                num_rows   = body_section.NumberOfRows
                num_cols   = body_section.NumberOfColumns

                print("\n--- Processing schedule: {0}".format(schedule.Name))
                print("Body Rows = {0}, Cols = {1}, total_slots = {2}"
                      .format(num_rows, num_cols, total_slots))

                # empties_dict[slotNum] = list of (row, col)
                empties_dict = {}

                # --------------------------------------------------------------------
                # 1) For each row, chunk columns that share the same empty slot
                # --------------------------------------------------------------------
                for row in range(num_rows):
                    # We'll track the current slot across columns
                    # If it changes, or we see a non-empty circuit, we finalize the old chunk
                    active_slot = None
                    col_list    = []  # columns for the currently active slot in this row

                    for col in range(num_cols):
                        ckt_id = schedule.GetCircuitIdByCell(row, col)
                        slot_num = schedule.GetSlotNumberByCell(row, col)

                        # Evaluate: is this cell empty and within valid slot range?
                        is_empty = (ckt_id == ElementId.InvalidElementId and
                                    1 <= slot_num <= total_slots)

                        # Check if it's the same "active_slot" we are currently grouping
                        if is_empty and (slot_num == active_slot):
                            # Keep collecting columns
                            col_list.append(col)

                        else:
                            # Either we've encountered a different slot or a non-empty circuit
                            # If we have a previously active slot with columns, finalize that chunk
                            if active_slot and col_list:
                                # Store all these row-col pairs
                                for c in col_list:
                                    empties_dict.setdefault(active_slot, []).append((row, c))

                            # Reset active slot to none
                            active_slot = None
                            col_list    = []

                            # If the new cell is empty, we start a new chunk
                            if is_empty:
                                active_slot = slot_num
                                col_list    = [col]

                    # Reached the end of this row => finalize any leftover chunk
                    if active_slot and col_list:
                        for c in col_list:
                            empties_dict.setdefault(active_slot, []).append((row, c))

                # --------------------------------------------------------------------
                # 2) We now have empties_dict keyed by slotNum => many (row,col) pairs
                # Convert to a list and assign half to Spare, half to Space
                # --------------------------------------------------------------------
                empties_list = sorted(empties_dict.items(), key=lambda kv: kv[0])
                total_empties = len(empties_list)
                half_count    = total_empties // 2

                print("Found {0} distinct empty slot groups.".format(total_empties))
                for i, (slot_num, rowcol_pairs) in enumerate(empties_list):
                    # Decide Spare vs. Space
                    is_spare = (i < half_count)

                    # We'll attempt each (row,col) until one works
                    assigned = False
                    for (r, c) in rowcol_pairs:
                        try:
                            if is_spare:
                                schedule.AddSpare(r, c)
                                schedule.SetLockSlot(r,c,0)
                                print("  [Slot {0}] SPARE added at row={1}, col={2}".format(slot_num, r, c))
                            else:
                                schedule.AddSpace(r, c)
                                schedule.SetLockSlot(r, c, 0)
                                print("  [Slot {0}] SPACE added at row={1}, col={2}".format(slot_num, r, c))

                            assigned = True
                            break  # if we only need one success, break here
                        except Exception as add_ex:
                            print("    Could not add circuit at row={0}, col={1}, slot={2}. Error: {3}"
                                  .format(r, c, slot_num, add_ex))

                    if not assigned:
                        print("  [Slot {0}] No valid row/col found to add Spare/Space.".format(slot_num))

            except Exception as e:
                print("Error processing schedule '{0}': {1}".format(schedule.Name, e))

        t.Commit()

console.close_others()
console.show()
sched = collect_schedules_to_process()
# get_schedule_info(sched)
fill_half_spare_half_space(sched, doc)
