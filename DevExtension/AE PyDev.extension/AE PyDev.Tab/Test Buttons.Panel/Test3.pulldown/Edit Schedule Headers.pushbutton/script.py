# -*- coding: utf-8 -*-

from pyrevit import revit, DB, forms
import re
from collections import defaultdict

doc = revit.doc
uidoc = revit.uidoc


def clean_schedule_headers(view_schedules):
    """
    Cleans headers in all provided schedules under a single transaction.
    - Converts to uppercase.
    - Removes _CED, _CEDT, _CEDR.
    """
    if not view_schedules:
        print("⚠️ No schedules provided to clean.")
        return

    print("\n▶ Cleaning headers for {} schedule(s)...".format(len(view_schedules)))

    with revit.Transaction("Clean Up Schedule Headers"):
        for vs in view_schedules:
            schedule_def = vs.Definition
            field_count = schedule_def.GetFieldCount()

            print("— Processing '{}': {} fields".format(vs.Name, field_count))

            for field_index in range(field_count):
                schedule_field = schedule_def.GetField(field_index)
                old_heading = schedule_field.ColumnHeading

                new_heading = old_heading.upper()
                new_heading = re.sub(r'(_CEDT|_CEDR|_CED)', '', new_heading).strip()

                if old_heading != new_heading:
                    schedule_field.ColumnHeading = new_heading
                    print("   ✔ Field {} | '{}' → '{}'".format(field_index, old_heading, new_heading))


def get_view_schedules_from_selection(elements):
    """
    Filters selection for ScheduleSheetInstances and returns their associated ViewSchedules.
    Prints discarded non-schedule element info.
    """
    valid_schedules = []
    discard_counter = defaultdict(int)

    for el in elements:
        if isinstance(el, DB.ScheduleSheetInstance):
            vs = doc.GetElement(el.ScheduleId)
            if isinstance(vs, DB.ViewSchedule):
                valid_schedules.append(vs)
        else:
            discard_counter[el.Category.Name if el.Category else "Unknown"] += 1

    for cat, count in discard_counter.items():
        print("⚠️  {} {} element(s) discarded.".format(count, cat))

    return valid_schedules


def prompt_user_for_schedules():
    """
    Prompt the user to select one or more schedules from the project.
    """
    all_schedules = [
        vs for vs in DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSchedule)
        if not vs.IsTemplate
    ]

    if not all_schedules:
        forms.alert("No schedules found in the project.", exitscript=True)


    selected_schedules = forms.select_schedules(
            title="Select Schedule(s) to Modify",
            multiple=True
    )

    if not selected_schedules:
        forms.alert("No schedules selected. Exiting.", exitscript=True)

    return selected_schedules


# --- MAIN LOGIC ---

schedules_to_process = []

# Case 1: If active view is a ViewSchedule, use it
active_view = uidoc.ActiveView
if isinstance(active_view, DB.ViewSchedule):
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
        schedules_to_process = prompt_user_for_schedules()

# Process all schedules at once
clean_schedule_headers(schedules_to_process)

print("\n✅ Done. {} schedule(s) processed.".format(len(schedules_to_process)))
