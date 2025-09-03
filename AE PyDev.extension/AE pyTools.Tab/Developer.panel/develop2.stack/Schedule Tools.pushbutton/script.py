# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()
COMMON_FIELDS = [
    "Appears in Schedule",
    "Schedule Filter",
    "Schedule Sort Order",
    "Family",
    "Type",
    "Product Type",
    "Identity Type Mark",
    "Identity Label Seperator",
    "Identity Mark",
    "Description",
    "Schedule Description",
    "Installed Location",
    "Space: Name",
    "Space: Number",
    "Schedule Notes"

]

REMOVE_FIELDS = ["Equipment Type_CEDT",
                 "Equipment Type ID_CEDT",
                 "Equipment ID_CEDT",
                 "Equipment Remarks_CEDT"]

field_replacement_map = {
    "Equipment Type_CEDT": "Product Type",
    "Equipment Type ID_CEDT": "Identity Type Mark",
    "Equipment ID_CEDT": "Identity Mark",
    "Equipment Remarks_CEDT": "Schedule Notes",
    "Area Served_CEDT": "Area Served"
}


# ---------------------------------------------------
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

    reorder_fields_by_name(definition, added_names)
    output.print_md("- Updated `{}` with: {}".format(schedule.Name, ", ".join(added_names)))


def remove_schedulable_fields_from_schedule(schedule, schedulable_fields_to_remove):
    definition = schedule.Definition
    current_field_ids = definition.GetFieldOrder()
    removed_names = []

    for schedulable_field in schedulable_fields_to_remove:
        parameter_name = get_schedulable_field_name(schedulable_field)
        for field_id in current_field_ids:
            field = definition.GetField(field_id)
            if field and field.GetName() == parameter_name:
                definition.RemoveField(field_id)
                removed_names.append(parameter_name)
                break  # remove only once per match

    output.print_md("- Removed from `{}`: {}".format(schedule.Name, ", ".join(removed_names)))


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

            reorder_fields_by_name(definition, updated_order)


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
def reorder_fields_by_name(definition, prioritized_names):
    current_fields = definition.GetFieldOrder()
    prioritized_fields = []
    remaining_fields = []

    for name in prioritized_names:
        for field_id in current_fields:
            field_name = definition.GetField(field_id).GetName()
            if field_name == name and field_id not in prioritized_fields:
                prioritized_fields.append(field_id)
                break

    for field_id in current_fields:
        if field_id not in prioritized_fields:
            remaining_fields.append(field_id)

    definition.SetFieldOrder(prioritized_fields + remaining_fields)


def hide_fields_in_schedule(schedule, field_names_to_hide):
    definition = schedule.Definition
    hidden = []
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        if field and field.GetName() in field_names_to_hide:
            if not field.IsHidden:
                field.IsHidden = True
                hidden.append(field.GetName())
    if hidden:
        output.print_md("- 👻 Hidden fields in `{}`: {}".format(schedule.Name, ", ".join(hidden)))


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
    field_map = {}
    for schedulable_field in schedule.Definition.GetSchedulableFields():
        if schedulable_field.FieldType == DB.ScheduleFieldType.Instance:
            parameter_id = schedulable_field.ParameterId
            if parameter_id and parameter_id != DB.ElementId.InvalidElementId:
                parameter_def = doc.GetElement(parameter_id)
                if parameter_def and hasattr(parameter_def, 'Name'):
                    field_map[parameter_def.Name] = schedulable_field
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
            reorder_fields_by_name(definition, added)
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

def main():
    source_schedule = forms.select_schedules(title="Select Source Schedule (Mechanical)", multiple=False)
    if not source_schedule:
        forms.alert("No source schedule selected.")
        return

    plumbing_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_PlumbingEquipment)
    if not plumbing_category:
        forms.alert("Plumbing Equipment category not found.")
        return

    source_def = source_schedule.Definition
    source_fields = source_def.GetFieldOrder()
    plumbing_schedulable = get_schedulable_instance_fields_from_category(plumbing_category)

    # Also include space-based fields
    space_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MEPSpaces)
    space_schedulable = get_schedulable_instance_fields_from_category(space_category)

    combined_schedulables = plumbing_schedulable.copy()
    combined_schedulables.update(space_schedulable)

    field_map = {}  # {field name: (schedulable field, is_hidden)}
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

    with revit.Transaction("Duplicate Schedule to Plumbing"):
        new_schedule = DB.ViewSchedule.CreateSchedule(doc, plumbing_category.Id)
        new_schedule.Name = "PE_{}".format(source_schedule.Name)
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

        for filt in source_def.GetFilters():
            try:
                new_def.InsertFilter(filt, len(new_def.GetFilters()))
            except:
                output.print_md("- ⚠️ Could not copy filter: `{}`".format(filt))

        for i in range(source_def.GetSortGroupFieldCount()):
            try:
                sort_data = source_def.GetSortGroupField(i)
                new_def.SetSortGroupField(i, sort_data)
            except:
                pass

    output.print_md("# ✅ Plumbing Schedule Created: `{}`".format(new_schedule.Name))
    if skipped_fields:
        output.print_md("## ⚠️ Fields Skipped (Not Available in Plumbing or Space Category):")
        for name in skipped_fields:
            output.print_md("- `{}`".format(name))

if __name__ == '__main__':
    main()
