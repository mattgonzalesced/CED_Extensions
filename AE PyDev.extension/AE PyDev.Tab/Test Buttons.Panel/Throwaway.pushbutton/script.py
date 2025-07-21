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
    "Equipment Type_CEDT",
    "Equipment Type ID_CEDT",
    "Equipment ID_CEDT",
    "Equipment Remarks_CEDT"
]


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


def main():
    # Mapping: Filter value → Schedule name
    filter_map = {
        "Air Curtain": "Air Curtain Schedule",
        "Destratification": "Destratification Fan Schedule",
        "Exhaust": "Exhaust Fan Schedule",
        "Gravity Ventilator": "Gravity Intake Ventilator Schedule",
        "Hood": "Hood Schedule",
        "Makeup Air": "Makeup Air Unit Schedule",
        "Outside Air": "Outside Air Unit Schedule",
        "RTU": "Rooftop Unit Schedule",
        "Condensing Unit": "Split System Condensing Unit Schedule",
        "Fan Coil": "Split System Fan Coil Unit Schedule"
    }

    schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
    schedule_by_name = {schedule.Name: schedule for schedule in schedules}

    filter_field_name = "Schedule Filter"
    filter_type = DB.ScheduleFilterType.Equal
    insert_at_top = True

    with revit.Transaction("Apply Filters to Unit Schedules"):
        for filter_value, schedule_name in filter_map.items():
            schedule = schedule_by_name.get(schedule_name)
            if not schedule:
                output.print_md("- ⚠️ Schedule not found: `{}`".format(schedule_name))
                continue

            definition = schedule.Definition
            existing_field_names = [definition.GetField(fid).GetName() for fid in definition.GetFieldOrder()]

            # Ensure the Schedule Filter field exists
            if filter_field_name not in existing_field_names:
                schedulable_map = get_schedulable_instance_fields(schedule)
                found = False
                for schedulable in schedulable_map.values():
                    if get_schedulable_field_name(schedulable) == filter_field_name:
                        definition.AddField(schedulable)
                        output.print_md("- 🔧 Added missing field `{}` to `{}`".format(filter_field_name, schedule_name))
                        found = True
                    break
                if not found:
                    output.print_md("- ❌ Could not find schedulable field `{}` for `{}`".format(filter_field_name, schedule_name))
                    continue

            # Add the filter
            add_filter_to_schedule(schedule, filter_field_name, filter_type, filter_value, insert_at_top)



if __name__ == '__main__':
    main()
