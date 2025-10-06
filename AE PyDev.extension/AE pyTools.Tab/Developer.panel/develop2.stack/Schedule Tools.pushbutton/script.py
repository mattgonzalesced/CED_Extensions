# -*- coding: utf-8 -*-
import re

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
logger = script.get_logger()
COMMON_FIELDS = [
    "Appears in Schedule",
    "Schedule Filter",
    "Schedule Sort Order",
    "Family",
    "Type",
    "Product Type",
    "Identity Type Mark",
    "Identity Label Separator",
    "Identity Mark",
    "Description",
    "Schedule Description",
    "Installed Location",
    "Space: Name",
    "Space: Number",
    "Schedule Notes"

]

FIELDS_TO_ADD = [
    "Circuit 1 Description_CEDT",
    "Circuit 1 Voltage_CED",
    "Circuit 1 Phase_CED",
    "Circuit 1 Number of Poles_CED",
    "Circuit 1 FLA_CED",
    "Circuit 1 MCA_CED",
    "Circuit 1 MOCP_CED",
    "Circuit 1 Apparent Load_CED",
    "Circuit 1 Remarks_CEDT"
    # "Circuit 2 Connection_CED",
    # "Circuit 2 Description_CEDT",
    # "Circuit 2 Load Classification_CED",
    # "Circuit 2 Voltage_CED",
    # "Circuit 2 Phase_CED",
    # "Circuit 2 Number of Poles_CED",
    # "Circuit 2 FLA_CED",
    # "Circuit 2 MCA_CED",
    # "Circuit 2 MOCP_CED",
    # "Circuit 2 Apparent Load_CED",
    # "Circuit 2 Remarks_CEDT",
]

FIELDS_TO_REMOVE = [
    "FLA_CED",
    "Horsepower_CED",
    "MCA_CED",
    "MOCP_CED",
    "MCA",
    "MOCP",
    "FLA",
    "Number of Poles_CED",
    "Voltage_CED",
    "Watts_CED",
    "Rated Voltage",
    "System Voltage",
    "Phase"
]

FIELDS_TO_UNHIDE = [
    "Identity Type Mark",
    "Identity Label Separator",
    "Identity Mark",
    "Interlock Identity Mark"
]

FIELDS_TO_HIDE = [
    "Circuit 1 Number of Poles_CED",
    "Identity Label Separator",
    "Identity Mark"
]

field_replacement_map = {
    "Equipment Type_CEDT": "Product Type",
    "Equipment Type ID_CEDT": "Identity Type Mark",
    "Equipment ID_CEDT": "Identity Mark",
    "Equipment Remarks_CEDT": "Schedule Notes",
    "Area Served_CEDT": "Area Served"
}


# ---------------------------------------------------

def get_selected_schedules(multiselect=True):
    """Resolve target schedules based on context:
    - Active schedule view
    - Selected schedule graphics on sheet
    - User picks from list
    """
    active_view = revit.active_view
    selected_ids = uidoc.Selection.GetElementIds()

    # Case 1: Active view is a schedule
    if isinstance(active_view, DB.ViewSchedule) and not active_view.IsTitleblockRevisionSchedule:
        return [active_view]

    # Case 2: Selected schedule graphics on a sheet
    if selected_ids:
        schedules = []
        for elid in selected_ids:
            el = doc.GetElement(elid)
            if isinstance(el, DB.ScheduleSheetInstance):
                sched = doc.GetElement(el.ScheduleId)
                if sched and isinstance(sched, DB.ViewSchedule):
                    schedules.append(sched)
        if schedules:
            return schedules

    # Case 3: Prompt user to pick
    all_schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
    picked = forms.SelectFromList.show(
        [s.Name for s in all_schedules],
        multiselect=multiselect,
        title="Select Schedules to Modify"
    )
    if not picked:
        return []
    return [s for s in all_schedules if s.Name in picked]


# Collect all project parameters
def get_all_project_params():
    param_bindings = doc.ParameterBindings
    iterator = param_bindings.ForwardIterator()
    iterator.Reset()
    param_names = []
    while iterator.MoveNext():
        definition = iterator.Key
        if isinstance(definition, DB.InternalDefinition):
            param_names.append(definition.Name)
    return sorted(param_names)

# Get parameter definition by name
def get_definition_by_name(name):
    iterator = doc.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        if isinstance(definition, DB.InternalDefinition) and definition.Name == name:
            return definition
    return None


# Get schedulable instance fields for a given schedule
def get_schedulable_instance_fields(schedule):
    fields = schedule.Definition.GetSchedulableFields()
    return {
        schedulable_field.ParameterId.Value: schedulable_field
        for schedulable_field in fields
        if schedulable_field.FieldType == DB.ScheduleFieldType.Instance
    }


# Add a list of schedulable fields to a schedule
def add_schedulable_fields_to_schedule(schedule, schedulable_fields):
    definition = schedule.Definition
    current_field_ids = definition.GetFieldOrder()
    added_names = []

    for schedulable_field in schedulable_fields:
        parameter_name = get_schedulable_field_name(schedulable_field)
        already_present = any(
            definition.GetField(field_id).GetName() == parameter_name for field_id in current_field_ids)
        if not already_present:
            definition.AddField(schedulable_field)
            added_names.append(parameter_name)

    reorder_fields(definition, added_names)
    output.print_md("- Updated `{}` with: {}".format(schedule.Name, ", ".join(added_names)))


def remove_schedulable_fields_from_schedule(schedule, field_names_to_remove):
    """
    Remove fields from a schedule by their display names.

    Args:
        schedule (DB.ViewSchedule): The schedule to modify.
        field_names_to_remove (list[str]): Names of fields to remove.
    """
    definition = schedule.Definition
    current_field_ids = list(definition.GetFieldOrder())
    remove_ids = []

    # First pass: collect field IDs to delete
    for field_id in current_field_ids:
        field = definition.GetField(field_id)
        if field and field.GetName() in field_names_to_remove:
            remove_ids.append(field_id)

    # Second pass: remove them safely
    removed_names = []
    for fid in remove_ids:
        try:
            field = definition.GetField(fid)
            if field:
                removed_names.append(field.GetName())
            definition.RemoveField(fid)
        except Exception as ex:
            output.print_md("- ❌ Could not remove fieldId {}: {}".format(fid, ex))

    if removed_names:
        output.print_md("- 🗑 Removed from `{}`: {}".format(schedule.Name, ", ".join(removed_names)))




def replace_schedule_fields(schedules, field_replacement_map):
    """
    Replaces specified schedule fields with new ones while preserving field order.

    Args:
        schedules (list[DB.ViewSchedule]): List of Revit schedule elements to process.
        field_replacement_map (dict): Dictionary where keys are existing field names,
                                      and values are the new field names to replace them with.


    Example Usage:
        def main():
        field_replacement_map = {
            "Type": "Equipment Type_CEDT",
            "Identity Label Seperator": "Schedule Notes"
        }

        selected_schedules = forms.select_schedules(title="Pick schedules to replace fields", multiple=True)
        if not selected_schedules:
            forms.alert("No schedules selected.")
            return

        replace_schedule_fields(selected_schedules, field_replacement_map)
        forms.alert("Field replacements complete.")
    """
    with revit.Transaction("Replace Fields in Schedules"):
        for schedule in schedules:
            definition = schedule.Definition
            current_field_ids = definition.GetFieldOrder()
            current_field_names = [definition.GetField(fid).GetName() for fid in current_field_ids]

            original_order = list(current_field_names)
            fields_to_remove_ids = []
            fields_to_add = []

            schedulable_map = get_schedulable_instance_fields(schedule)

            for field_id in current_field_ids:
                field = definition.GetField(field_id)
                if not field:
                    continue
                field_name = field.GetName()

                # Only process fields that exist and have a replacement
                if field_name in field_replacement_map:
                    replacement_name = field_replacement_map[field_name]

                    # Find replacement in schedulable map
                    replacement_found = False
                    for schedulable in schedulable_map.values():
                        if get_schedulable_field_name(schedulable) == replacement_name:
                            fields_to_add.append(schedulable)
                            fields_to_remove_ids.append(field_id)
                            replacement_found = True
                            break

                    if not replacement_found:
                        output.print_md(
                            "- ⚠️ Replacement field `{}` not found in `{}`. Skipping.".format(replacement_name,
                                                                                              schedule.Name))

            # Remove fields only if matched
            for field_id in fields_to_remove_ids:
                definition.RemoveField(field_id)

            add_schedulable_fields_to_schedule(schedule, fields_to_add)

            # Reorder: Replace old names with new ones
            updated_order = []
            for name in original_order:
                if name in field_replacement_map and field_replacement_map[name] in [get_schedulable_field_name(f) for f
                                                                                     in fields_to_add]:
                    updated_order.append(field_replacement_map[name])
                elif name not in field_replacement_map.keys():
                    updated_order.append(name)

            reorder_fields(definition, updated_order)


def add_filter_to_schedule(schedule, parameter_name, filter_type, value, insert_on_top=False):
    definition = schedule.Definition

    # Find field ID by name
    field_id = get_schedule_field_id_by_name(schedule, parameter_name)
    existing_filters = list(definition.GetFilters())
    if not field_id:
        output.print_md("- ❌ Field `{}` not found in `{}`".format(parameter_name, schedule.Name))
        return

    # Create filter by type
    if isinstance(value, str):
        new_filter = DB.ScheduleFilter(field_id, filter_type, value)
    elif isinstance(value, int):
        new_filter = DB.ScheduleFilter(field_id, filter_type, value)
    elif isinstance(value, float):
        new_filter = DB.ScheduleFilter(field_id, filter_type, float(value))
    elif isinstance(value, DB.ElementId):
        new_filter = DB.ScheduleFilter(field_id, filter_type, value)
    else:
        output.print_md("- ❌ Unsupported value type for filter: `{}`".format(type(value)))
        return

    insert_index = 0 if insert_on_top else len(existing_filters)
    definition.InsertFilter(new_filter, insert_index)

    output.print_md(
        "- ✅ Inserted filter on `{}` to `{}` at index `{}`".format(parameter_name, schedule.Name, insert_index))


# Get the ScheduleFieldId of a field by name from a schedule

def get_schedule_field_id_by_name(schedule, field_name):
    definition = schedule.Definition
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        if field and field.GetName() == field_name:
            return field.FieldId
    return None


# Get the display name for a schedulable field
def get_schedulable_field_name(schedulable_field):
    parameter_id = schedulable_field.ParameterId
    parameter_def = doc.GetElement(parameter_id)
    if parameter_def and hasattr(parameter_def, 'Name'):
        return parameter_def.Name
    try:
        built_in_param = DB.BuiltInParameter(parameter_id.Value)
        return DB.LabelUtils.GetLabelFor(built_in_param)
    except:
        return "Param_{}".format(parameter_id.Value)


# Reorder schedule fields to match a preferred list exactly
def reorder_fields(sched_definition, field_names,
                            position="end", after_name=None):
    """
    Reorder schedule fields relative to the current field order.

    Args:
        sched_definition (DB.ScheduleDefinition):
            The schedule definition object to reorder fields in.
        field_names (list[str]):
            List of field display names to move/reorder.
        position (str, optional):
            Where to place the specified fields. Options are:
            - "start": move fields to the beginning.
            - "end": move fields to the end.
            - "after": insert fields immediately after `after_name`.
            Defaults to "end".
        after_name (str, optional):
            The field name after which the reordered fields will be inserted.
            Required if position="after".

    Behavior:
        - Only reorders fields whose names match items in `field_names`.
        - Maintains the relative order of all other fields.
        - If a field name does not exist in the schedule, it is skipped.
    """
    current_fields = sched_definition.GetFieldOrder()

    # Collect matching field ids
    to_move = []
    for name in field_names:
        for field_id in current_fields:
            field = sched_definition.GetField(field_id)
            if field and field.GetName() == name and field_id not in to_move:
                to_move.append(field_id)
                break

    # Build base order without moved fields
    base_order = [fid for fid in current_fields if fid not in to_move]

    if position == "start":
        new_order = to_move + base_order

    elif position == "end":
        new_order = base_order + to_move

    elif position == "after":
        if not after_name:
            raise ValueError("`after_name` must be provided when position='after'")
        insert_index = None
        for idx, fid in enumerate(base_order):
            field = sched_definition.GetField(fid)
            if field and field.GetName() == after_name:
                insert_index = idx + 1
                break
        if insert_index is None:
            # If target not found, just append to end
            new_order = base_order + to_move
        else:
            new_order = base_order[:insert_index] + to_move + base_order[insert_index:]

    else:
        raise ValueError("Invalid position '{}'. Use 'start', 'end', or 'after'.".format(position))

    sched_definition.SetFieldOrder(new_order)


def hide_fields(schedule, field_names_to_hide):
    definition = schedule.Definition
    hidden = []
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        logger.debug("Field to hide: {}".format(field.GetName()))
        if field and field.GetName() in field_names_to_hide:
            if not field.IsHidden:
                logger.debug("FIELD IS NOT HIDDEN! HIDING NOW")
                field.IsHidden = True
                hidden.append(field.GetName())
    if hidden:
        output.print_md("- 👻 Hidden fields in `{}`: {}".format(schedule.Name, ", ".join(hidden)))

def unhide_fields(schedule, field_names_to_unhide):
    definition = schedule.Definition
    unhidden = []
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        if field and field.GetName() in field_names_to_unhide:
            logger.debug("Field to Unhide: {}".format(field.GetName()))
            if field.IsHidden:
                logger.debug("FIELD IS HIDDEN! UNHIDING NOW")
                field.IsHidden = False
                unhidden.append(field.GetName())
    if unhidden:
        output.print_md("- 👻 Hidden fields in `{}`: {}".format(schedule.Name, ", ".join(unhidden)))

# Hide specified parameters in all schedules
def hide_fields_in_all_schedules(param_names_to_hide):
    hidden_fields_report = {}
    schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()

    for schedule in schedules:
        definition = schedule.Definition
        field_ids = definition.GetFieldOrder()
        hidden_names = []
        for field_id in field_ids:
            field = definition.GetField(field_id)
            if not field:
                continue
            field_name = field.GetName()
            if field_name in param_names_to_hide and not field.IsHidden:
                field.IsHidden = True
                hidden_names.append(field_name)
        if hidden_names:
            hidden_fields_report[schedule.Name] = hidden_names
    return hidden_fields_report


# Pick valid model categories for schedule creation
def pick_categories():
    all_categories = doc.Settings.Categories
    schedulable = [category for category in all_categories if
                   category.CategoryType == DB.CategoryType.Model and not category.SubCategories.IsEmpty]
    category_options = ["{} ({})".format(category.Name, category.Id.Value) for category in schedulable]
    selected = forms.SelectFromList.show(category_options, multiselect=True,
                                         title="Select Categories to Create Schedules")
    if not selected:
        return []
    return [category for category in schedulable if "{} ({})".format(category.Name, category.Id.Value) in selected]


# Get schedulable fields by name from a schedule
def get_schedulable_field_name_map(schedule):
    """
    Returns {field_name: schedulable_field} for both instance, type, and related fields.
    """
    field_map = {}
    for sf in schedule.Definition.GetSchedulableFields():
        if sf.FieldType in (DB.ScheduleFieldType.Instance,
                            DB.ScheduleFieldType.Type,
                            DB.ScheduleFieldType.Material):
            # Try resolving by definition
            param_id = sf.ParameterId
            param_def = doc.GetElement(param_id) if param_id != DB.ElementId.InvalidElementId else None
            if param_def and hasattr(param_def, "Name"):
                field_map[param_def.Name] = sf
            else:
                # fallback: use display label
                try:
                    field_map[sf.GetName()] = sf
                except:
                    field_map["Param_{}".format(param_id.IntegerValue)] = sf
    return field_map



# Create a schedule for each selected category and let user pick fields to add
def create_and_populate_schedules():
    created_schedules = {}
    transaction_group = DB.TransactionGroup(doc, "Create & Configure Schedules")
    transaction_group.Start()

    with revit.Transaction("Create Schedules"):
        selected_categories = pick_categories()
        if not selected_categories:
            forms.alert("No categories selected.")
            transaction_group.RollBack()
            return

        for category in selected_categories:
            schedule = DB.ViewSchedule.CreateSchedule(doc, category.Id)
            schedule.Name = "Auto - {}".format(category.Name)
            created_schedules[category.Id.Value] = schedule
            output.print_md("- Created schedule: **{}**".format(schedule.Name))

    with revit.Transaction("Add Fields"):
        all_field_names = set()
        schedule_field_maps = {}

        for category_id, schedule in created_schedules.items():
            field_map_for_schedule = get_schedulable_field_name_map(schedule)
            schedule_field_maps[category_id] = field_map_for_schedule
            all_field_names.update(field_map_for_schedule.keys())

        selected_fields = forms.SelectFromList.show(
            sorted(all_field_names),
            multiselect=True,
            title="Select Parameters to Add to New Schedules"
        )
        if not selected_fields:
            forms.alert("No parameters selected.")
            transaction_group.RollBack()
            return

        for category_id, schedule in created_schedules.items():
            definition = schedule.Definition
            field_map_for_schedule = schedule_field_maps[category_id]
            added = []
            for name in selected_fields:
                if name in field_map_for_schedule:
                    definition.AddField(field_map_for_schedule[name])
                    added.append(name)
            reorder_fields(definition, added)
            output.print_md("- Updated schedule `{}` with: {}".format(schedule.Name, ", ".join(added)))

    transaction_group.Assimilate()
    forms.alert("Schedules created and configured.", title="Success")


# Pick schedulable instance fields from one or more schedules
def pick_schedulable_fields_from_schedules(schedules):
    field_map = {}
    for schedule in schedules:
        for schedulable_field in schedule.Definition.GetSchedulableFields():
            if schedulable_field.FieldType == DB.ScheduleFieldType.Instance:
                parameter_id = schedulable_field.ParameterId
                if parameter_id and parameter_id != DB.ElementId.InvalidElementId:
                    parameter_def = doc.GetElement(parameter_id)
                    if parameter_def and hasattr(parameter_def, 'Name'):
                        parameter_name = parameter_def.Name
                    else:
                        try:
                            built_in_param = DB.BuiltInParameter(parameter_id.Value)
                            parameter_name = DB.LabelUtils.GetLabelFor(built_in_param)
                        except:
                            parameter_name = "Param_{}".format(parameter_id.Value)
                    field_map[parameter_name] = schedulable_field

    selected_names = forms.SelectFromList.show(
        sorted(field_map.keys()),
        multiselect=True,
        title="Select Schedulable Fields"
    )
    if not selected_names:
        return []

    return [field_map[name] for name in selected_names]


def get_schedulable_instance_fields_from_category(category):
    """Returns schedulable instance fields from a given category."""
    schedulable_fields = []
    with revit.Transaction("Get Schedulable Fields from Category"):
        temp_schedule = DB.ViewSchedule.CreateSchedule(doc, category.Id)
        schedulable_fields = temp_schedule.Definition.GetSchedulableFields()
        doc.Delete(temp_schedule.Id)

    return {
        f.ParameterId.Value: f
        for f in schedulable_fields
        if f.FieldType == DB.ScheduleFieldType.Instance
    }


def convert_to_new_category(source_schedule, new_prefix="PE_",
                                 include_space_fields=True,
                                 target_category=DB.BuiltInCategory.OST_PlumbingEquipment,
                                 space_category=DB.BuiltInCategory.OST_MEPSpaces):
    """
    Duplicate a source schedule (typically mechanical) into a Plumbing Equipment schedule.

    Args:
        source_schedule (DB.ViewSchedule): The existing schedule to duplicate from.
        new_prefix (str, optional): Prefix for the new schedule name. Defaults to "PE_".
        include_space_fields (bool, optional): Whether to also match fields from Spaces category.
        target_category (DB.BuiltInCategory or DB.Category, optional):
            Target category for the new schedule. Defaults to PlumbingEquipment.
        space_category (DB.BuiltInCategory or DB.Category, optional):
            Secondary category for space parameters. Only used if include_space_fields=True.

    Returns:
        tuple(DB.ViewSchedule, list[str]):
            - The newly created plumbing schedule.
            - A list of skipped field names that could not be added.

    Behavior:
        - Creates a new schedule for the target plumbing category.
        - Attempts to match field names from the source schedule with schedulable
          fields in the plumbing category (and optionally space category).
        - Preserves field order and hidden/visible states.
        - Copies filters and sort/group settings from the source schedule.
        - Returns a report of any skipped fields.
    """
    # Resolve category objects
    if isinstance(target_category, DB.BuiltInCategory):
        target_category = DB.Category.GetCategory(doc, target_category)
    if include_space_fields and isinstance(space_category, DB.BuiltInCategory):
        space_category = DB.Category.GetCategory(doc, space_category)

    if not target_category:
        raise ValueError("Plumbing Equipment category not found.")

    source_def = source_schedule.Definition
    source_fields = source_def.GetFieldOrder()

    # Collect schedulables from plumbing
    plumbing_schedulable = get_schedulable_instance_fields_from_category(target_category)
    combined_schedulables = plumbing_schedulable.copy()

    if include_space_fields and space_category:
        space_schedulable = get_schedulable_instance_fields_from_category(space_category)
        combined_schedulables.update(space_schedulable)

    # Map source field names to target schedulables
    field_map = {}
    skipped_fields = []

    for field_id in source_fields:
        source_field = source_def.GetField(field_id)
        if not source_field:
            continue
        field_name = source_field.GetName()
        is_hidden = source_field.IsHidden

        match = None
        for schedulable in combined_schedulables.values():
            if get_schedulable_field_name(schedulable) == field_name:
                match = schedulable
                break

        if match:
            field_map[field_name] = (match, is_hidden)
        else:
            skipped_fields.append(field_name)

    # Create the new schedule
    with revit.Transaction("Duplicate Schedule to Plumbing"):
        new_schedule = DB.ViewSchedule.CreateSchedule(doc, target_category.Id)
        new_schedule.Name = "{}{}".format(new_prefix, source_schedule.Name)
        new_def = new_schedule.Definition

        added_field_ids = []
        for field_id in source_def.GetFieldOrder():
            source_field = source_def.GetField(field_id)
            if not source_field:
                continue
            field_name = source_field.GetName()
            if field_name in field_map:
                schedulable_field, is_hidden = field_map[field_name]
                try:
                    new_field = new_def.AddField(schedulable_field)
                    new_def.GetField(new_field.FieldId).IsHidden = is_hidden
                    added_field_ids.append(new_field.FieldId)
                except:
                    skipped_fields.append(field_name)

        new_def.SetFieldOrder(added_field_ids)

        # Copy filters
        for filt in source_def.GetFilters():
            try:
                new_def.InsertFilter(filt, len(new_def.GetFilters()))
            except:
                output.print_md("- ⚠️ Could not copy filter: `{}`".format(filt))

        # Copy sort/group settings
        for i in range(source_def.GetSortGroupFieldCount()):
            try:
                sort_data = source_def.GetSortGroupField(i)
                new_def.SetSortGroupField(i, sort_data)
            except:
                pass

    return new_schedule, skipped_fields



def apply_find_replace_to_header(header_text, replacements, uppercase=True, strip_suffixes=True):
    """
    Process a schedule header string.

    Args:
        header_text (str): Original column heading text.
        replacements (dict): Mapping of {find: replace}.
        uppercase (bool): If True, convert to uppercase.
        strip_suffixes (bool): If True, strip known suffixes/prefixes (_CED, CKT_, etc.).

    Returns:
        str: The cleaned/modified header text.
    """
    new_text = header_text or ""

    if uppercase:
        new_text = new_text.upper()

    if strip_suffixes:
        new_text = re.sub(r'(_CEDT|_CEDR|_CED|CKT_)', '', new_text)

    if replacements:
        for find_str, repl_str in replacements.items():
            new_text = new_text.replace(find_str, repl_str)

    return new_text.strip()


def batch_update_schedule_headers(schedules, replacements=None, uppercase=True, strip_suffixes=True):
    """
    Update headers for all fields in the given schedules.

    Args:
        schedules (list[DB.ViewSchedule]): Schedules to modify.
        replacements (dict): {find: replace} mapping applied to header text.
        uppercase (bool): If True, force uppercase.
        strip_suffixes (bool): If True, strip known suffixes/prefixes.
    """
    if not schedules:
        forms.alert("No schedules to update.")
        return

    with revit.Transaction("Update Schedule Headers"):
        for schedule in schedules:
            definition = schedule.Definition
            field_count = definition.GetFieldCount()
            updated = []

            for idx in range(field_count):
                field = definition.GetField(idx)
                old_heading = field.ColumnHeading
                new_heading = apply_find_replace_to_header(
                    old_heading, replacements or {}, uppercase, strip_suffixes
                )
                if new_heading != old_heading:
                    field.ColumnHeading = new_heading
                    updated.append((old_heading, new_heading))

            if updated:
                output.print_md("### ✏ `{}`".format(schedule.Name))
                for old, new in updated:
                    output.print_md("- '{}' → '{}'".format(old, new))



def main():
    schedules = get_selected_schedules(multiselect=True)
    if not schedules:
        forms.alert("No schedules selected.")
        return

    # with revit.Transaction("Add Fields", doc=doc):
    #     for sched in schedules:
    #         convert_to_new_category(sched)
    #
    with DB.TransactionGroup(doc, "Modify Schedule Fields") as tg:
        tg.Start()

        for schedule in schedules:
            definition = schedule.Definition
            output.print_md("### ✏ Modifying `{}`".format(schedule.Name))



            # --- Add desired fields ---
            schedulable_map = get_schedulable_instance_fields(schedule)
            fields_to_add = []
            for name in FIELDS_TO_ADD:
                for sf in schedulable_map.values():
                    if get_schedulable_field_name(sf) == name:
                        fields_to_add.append(sf)
                        break

            with revit.Transaction("Add Fields", doc=doc):
                add_schedulable_fields_to_schedule(schedule, fields_to_add)

            with revit.Transaction("Unhide Fields", doc=doc):
                unhide_fields(schedule, FIELDS_TO_UNHIDE)

            # --- Reorder ---
            with revit.Transaction("Reorder Fields", doc=doc):
                reorder_fields(definition, FIELDS_TO_ADD, position="after", after_name="Interlock Identity Mark")

            # --- Remove unwanted fields ---
            with revit.Transaction("Remove Fields", doc=doc):
                remove_schedulable_fields_from_schedule(schedule, FIELDS_TO_REMOVE)

             # Example 2: With find/replace rules
            find_replace_rules = {"Circuit 1 Description_CEDT":"Circuit Description",
                                    "Circuit 1 Voltage_CED":  "Voltage",
                                    "Circuit 1 Phase_CED":  "Phase",
                                    "Circuit 1 Number of Poles_CED":  "Number of Poles",
                                    "Circuit 1 FLA_CED":  "FLA",
                                    "Circuit 1 MCA_CED":  "MCA",
                                    "Circuit 1 MOCP_CED":  "MOCP",
                                    "Circuit 1 Apparent Load_CED":  "Apparent Load",
                                    "Circuit 1 Remarks_CEDT":  "Circuit Remarks"
            }

            batch_update_schedule_headers(schedules, replacements=find_replace_rules,uppercase=False, strip_suffixes=False)
        tg.Assimilate()


if __name__ == '__main__':
    main()


"""Remove/Add/Reorder Example

 schedules = get_selected_schedules()
    if not schedules:
        forms.alert("No schedules selected.")
        return

    with DB.TransactionGroup(doc, "Modify Schedule Fields") as tg:
        tg.Start()

        for schedule in schedules:
            definition = schedule.Definition
            output.print_md("### ✏ Modifying `{}`".format(schedule.Name))

            # --- Remove unwanted fields ---
            with revit.Transaction("Remove Fields", doc=doc):
                remove_schedulable_fields_from_schedule(schedule, FIELDS_TO_REMOVE)

            # --- Add desired fields ---
            schedulable_map = get_schedulable_instance_fields(schedule)
            fields_to_add = []
            for name in FIELDS_TO_ADD:
                for sf in schedulable_map.values():
                    if get_schedulable_field_name(sf) == name:
                        fields_to_add.append(sf)
                        break

            with revit.Transaction("Add Fields", doc=doc):
                add_schedulable_fields_to_schedule(schedule, fields_to_add)

            # --- Reorder ---
            with revit.Transaction("Reorder Fields", doc=doc):
                reorder_fields(definition, FIELDS_TO_ADD, position="after",after_name="Gas Regulator Inlet")

        tg.Assimilate()

    forms.alert("Field modifications complete.")
    """


"""Find and replace example
    schedules = get_selected_schedules()
    if not schedules:
        return

    # Example 1: Just strip suffixes/prefixes + uppercase
    batch_update_schedule_headers(schedules)

    # Example 2: With find/replace rules
    find_replace_rules = {"CIRCUIT 1 ":""

    }
    batch_update_schedule_headers(schedules, replacements=find_replace_rules)
    """