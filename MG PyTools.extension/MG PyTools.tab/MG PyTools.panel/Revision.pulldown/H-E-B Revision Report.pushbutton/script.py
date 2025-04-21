# -*- coding: utf-8 -*-
__title__ = "Revision Report"
__doc__ = """Version = 1.1
Date    = 12.05.2024
________________________________________________________________

________________________________________________________________
Author: AEvelina"""

from pyrevit import script, forms
from pyrevit import coreutils
from pyrevit import revit, DB
from pyrevit.revit import query
from pyrevit import HOST_APP
from os.path import dirname, join

# Initialize console
console = script.get_output()
console.close_others()
console.set_height(800)

script_dir = dirname(__file__)
logo_path = join(script_dir, 'CED_Logo_H.png')
doc = revit.doc


def validate_additional_parameters(param_names):
    """Validate that parameters exist in the current project."""
    sample_cloud = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
        .WhereElementIsNotElementType() \
        .FirstElement()

    if not sample_cloud:
        return []  # No revision clouds found, skip validation

    invalid_params = []
    for param_name in param_names:
        if not sample_cloud.LookupParameter(param_name):
            invalid_params.append(param_name)
    return invalid_params


def clean_param_name(param_name):
    """Remove suffixes like '_CED' or '_CEDT' from parameter names."""
    suffixes = ["_CED", "_CEDT"]
    for suffix in suffixes:
        if param_name.endswith(suffix):
            return param_name[:-len(suffix)]
    return param_name


def get_param_value_by_name(element, param_name):
    """Fetch parameter value by name."""
    param = element.LookupParameter(param_name)
    value = query.get_param_value(param) if param else None
    return value if value is not None else ""  # Replace None with an empty string


def get_revision_data_by_sheet(param_names):
    """Group revision clouds by revisions using ViewSheet.GetAllRevisionCloudIds()."""

    revision_data = {}

    # Collect all sheets in the document
    all_sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)

    for sheet in all_sheets:
        sheet_number = query.get_param_value(sheet.LookupParameter("Sheet Number"))

        # Get all revision clouds on this sheet
        revision_cloud_ids = sheet.GetAllRevisionCloudIds()
        revision_clouds = [doc.GetElement(cloud_id) for cloud_id in revision_cloud_ids]

        for cloud in revision_clouds:
            revision_id = cloud.RevisionId.IntegerValue

            # Ensure data structure for this revision exists
            if revision_id not in revision_data:
                revision_data[revision_id] = []

            # Collect cloud data
            comment = query.get_param_value(cloud.LookupParameter("Comments"))
            rfi_number = query.get_param_value(cloud.LookupParameter("RFI Number_CEDT"))
            if not comment:  # Skip clouds without comments
                continue

            # Get additional parameter values
            additional_data = {}
            for param_name in param_names:
                additional_data[param_name] = get_param_value_by_name(cloud, param_name)

            # Combine all data manually
            cloud_data = {
                "RFI Number": rfi_number,
                "Sheet Number": sheet_number,
                "Comments": comment,
            }
            cloud_data.update(additional_data)
            revision_data[revision_id].append(cloud_data)

    return revision_data


def get_revision_data_from_cloud(clouds, selected_revisions, param_names):
    """Group revision clouds by selected revisions for Revit versions below 2024."""
    revision_data = {rev.Id.IntegerValue: [] for rev in selected_revisions}

    for cloud in clouds:
        rev_id = cloud.RevisionId.IntegerValue
        if rev_id not in revision_data:
            continue

        # Get sheets associated with the revision cloud
        sheet_ids = cloud.GetSheetIds()  # This method gets all sheets a cloud is associated with

        # Iterate through associated sheets
        for sheet_id in sheet_ids:
            sheet = doc.GetElement(sheet_id)
            sheet_number = query.get_param_value(sheet.LookupParameter("Sheet Number"))

            # Get the comment parameter value
            comment = query.get_param_value(cloud.LookupParameter("Comments"))
            rfi_number = query.get_param_value(cloud.LookupParameter("RFI Number_CEDT"))
            if not comment:  # Skip clouds without comments
                continue

            # Get additional parameter values
            additional_data = {}
            for param_name in param_names:
                additional_data[param_name] = get_param_value_by_name(cloud, param_name)

            # Combine data for this cloud and sheet
            cloud_data = {
                "RFI Number": rfi_number,
                "Sheet Number": sheet_number,
                "Comments": comment,
            }
            cloud_data.update(additional_data)

            # Append cloud data to the revision entry
            revision_data[rev_id].append(cloud_data)

    return revision_data


def deduplicate_clouds(cloud_data):
    """Remove duplicate comments for the same sheet."""
    seen_sheets_comments = set()
    deduplicated_clouds = []
    for cloud in cloud_data:
        sheet_number = cloud.get("Sheet Number", "N/A")
        comment = cloud.get("Comments", "").strip()  # Ensure comment is cleaned

        # Skip if the comment is empty or already seen
        if not comment or (sheet_number, comment) in seen_sheets_comments:
            continue

        deduplicated_clouds.append(cloud)
        seen_sheets_comments.add((sheet_number, comment))
    return deduplicated_clouds


def print_project_metadata():
    """Print project metadata like project number, client, and report date."""
    report_date = coreutils.current_date()
    project_info = revit.query.get_project_info()
    project_name = project_info.name
    project_number = project_info.number
    client_name = project_info.client_name

    console.print_html(
        "<img src='{}' width='150px' style='margin: 0; padding: 0;' />".format(logo_path)
    )

    console.print_md("**Coolsys Energy Design**")
    console.print_md("## Project Revision Summary")
    console.print_md("---")
    console.print_md("Project Number: **{}**".format(project_number))
    console.print_md("Client: **{}**".format(client_name))
    console.print_md("Project Name: **{}**".format(project_name))
    console.print_md("Report Date: **{}**".format(report_date))
    console.print_md("---")


def print_revision_report(revisions, revision_data, param_names):
    """Print the revision report."""
    for rev in revisions:
        revision_id = rev.Id.IntegerValue

        # Retrieve revision details
        revision_number = query.get_param_value(rev.LookupParameter("Revision Number"))
        revision_date = query.get_param_value(rev.LookupParameter("Revision Date"))
        revision_desc = query.get_param_value(rev.LookupParameter("Revision Description"))

        # Print the revision header
        console.print_md(
            "### Revision Number: {0} | Date: {1} | Description: {2}".format(revision_number, revision_date,
                                                                             revision_desc)
        )

        # Handle cases where no clouds are associated with this revision
        if revision_id not in revision_data or not revision_data[revision_id]:
            console.print_md("No revision clouds found for this revision.")
            console.insert_divider()
            continue

        # Sort and deduplicate clouds
        rev_clouds = sorted(revision_data[revision_id], key=lambda x: x["Sheet Number"] or "")
        deduplicated_clouds = deduplicate_clouds(rev_clouds)

        # Prepare table columns
        columns = ["RFI Number", "Sheet #[s]", "Description of Change"] + [clean_param_name(param) for param in param_names]

        table_data = []
        for cloud in deduplicated_clouds:
            row = [cloud.get("RFI Number", ""),
                   cloud.get("Sheet Number", ""),
                   cloud.get("Comments", "")]
            # Add custom parameter values
            for param_name in param_names:
                row.append(cloud.get(param_name, ""))  # Use original parameter name here
            table_data.append(row)

        if table_data:
            console.print_table(table_data, columns=columns)

        else:
            console.print_md("No revision clouds found for this revision.")
        console.insert_divider()


def main():
    # Get the script configuration
    config = script.get_config("revision_parameters_config")
    param_names_raw = getattr(config, "selected_param_names", "")  # Use `getattr` for safety
    param_names = param_names_raw.split(",") if param_names_raw else []

    # Validate additional parameters
    invalid_params = validate_additional_parameters(param_names)
    if invalid_params:
        # Alert user and reset to default
        forms.alert(
            "The following parameters are not assigned to Revision Clouds in this project:\n" +
            "\n".join(invalid_params) +
            "\n\nDefault output will be used. Exit and Shift+Click button to configure custom parameters for this project."
        )
        param_names = []

        # Update the config file to defaults
        config.selected_param_names = ""  # Reset config to default (no additional parameters)
        script.save_config()

    # Select revisions
    revisions = forms.select_revisions(button_name="Select Revision", multiple=True)
    if not revisions:
        script.exit()

    # Determine Revit version and use appropriate logic
    revit_version = int(HOST_APP.version)
    if revit_version >= 2024:
        revision_data = get_revision_data_by_sheet(param_names)
    else:
        all_clouds = DB.FilteredElementCollector(revit.doc) \
            .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
            .WhereElementIsNotElementType() \
            .ToElements()
        revision_data = get_revision_data_from_cloud(all_clouds, revisions, param_names)

    # Print metadata and report
    print_project_metadata()
    print_revision_report(revisions, revision_data, param_names)


# Execute main function
if __name__ == "__main__":
    main()
