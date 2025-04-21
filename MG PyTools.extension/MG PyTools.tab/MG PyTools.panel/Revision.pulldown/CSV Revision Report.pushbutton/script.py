from pyrevit import script, forms
from pyrevit import coreutils
from pyrevit import revit, DB
from pyrevit.revit import query
from pyrevit import HOST_APP
from os.path import dirname, join
import csv
import tempfile
import os

# Initialize console
console = script.get_output()
console.close_others()
console.set_height(800)

doc = revit.doc

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


def display_csv_in_popup(revision_data, param_names, revisions):
    """Display the revision data in a popup as a CSV."""
    temp_csv_path = tempfile.mktemp(suffix=".csv")
    with open(temp_csv_path, mode="w") as temp_csv:  # Removed newline=""
        writer = csv.writer(temp_csv)

        # Iterate over revisions to group data
        for rev in revisions:
            revision_id = rev.Id.IntegerValue
            revision_number = query.get_param_value(rev.LookupParameter("Revision Number"))
            revision_date = query.get_param_value(rev.LookupParameter("Revision Date"))
            revision_desc = query.get_param_value(rev.LookupParameter("Revision Description"))

            # Write revision header
            writer.writerow(["Revision Number: {}".format(revision_number),
                             "Date: {}".format(revision_date),
                             "Description: {}".format(revision_desc)])

            # Write RFI headers
            writer.writerow(["RFI Number", "Sheet Number", "Comments"] + param_names)

            # Write RFI data for the current revision
            if revision_id in revision_data:
                for cloud in revision_data[revision_id]:
                    row = [
                        cloud.get("RFI Number", ""),
                        cloud.get("Sheet Number", ""),
                        cloud.get("Comments", "")
                    ] + [cloud.get(param, "") for param in param_names]
                    writer.writerow(row)

    # Display the file in a popup
    os.startfile(temp_csv_path)


def main():
    # Get the script configuration
    config = script.get_config("revision_parameters_config")
    param_names_raw = getattr(config, "selected_param_names", "")  # Use `getattr` for safety
    param_names = param_names_raw.split(",") if param_names_raw else []

    # Select revisions
    revisions = forms.select_revisions(button_name="Select Revision", multiple=True)
    if not revisions:
        script.exit()

    # Get revision data
    revit_version = int(HOST_APP.version)
    if revit_version >= 2024:
        revision_data = get_revision_data_by_sheet(param_names)
    else:
        all_clouds = DB.FilteredElementCollector(revit.doc) \
            .OfCategory(DB.BuiltInCategory.OST_RevisionClouds) \
            .WhereElementIsNotElementType() \
            .ToElements()
        revision_data = get_revision_data_from_cloud(all_clouds, revisions, param_names)

    # Display CSV in popup
    display_csv_in_popup(revision_data, param_names, revisions)


# Execute main function
if __name__ == "__main__":
    main()
