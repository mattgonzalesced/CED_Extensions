# -*- coding: utf-8 -*-
__title__   = "Revision Report"
__doc__     = """Version = 1.0
Date    = 09.18.2024
________________________________________________________________
Description:

This script generates a detailed report for the selected revisions.

________________________________________________________________
How-To:
1. Select the desired revisions from the list.
2. The report will provide sheet numbers, names, and comments for each revision cloud.

________________________________________________________________
Author: AEvelina"""

from pyrevit import script, forms
from pyrevit import coreutils
from pyrevit import revit, DB
from pyrevit.revit import query
from os.path import dirname,join


# Initialize console
console = script.get_output()
console.close_others()
console.set_height(800)


# Get the directory of the current script
script_dir = dirname(__file__)

# Construct the relative path to your logo
logo_path = join(script_dir, 'CED_Logo_H.png')

# Enhanced helper function to retrieve and cache project metadata
def get_project_metadata():
    """Retrieve and return project metadata."""
    project_info = revit.query.get_project_info()
    return {
        "project_name": project_info.name,
        "project_number": project_info.number,
        "client_name": project_info.client_name,
        "report_date": coreutils.current_date()
    }

# Enhanced project metadata printing with error handling
def print_project_metadata(metadata):
    """Print project metadata like project number, client, and report date."""
    if metadata:
        console.print_html("<img src='{}' width='150px' style='margin: 0; padding: 0;' />".format(logo_path))
        console.print_md("**Coolsys Energy Design**")
        console.print_md("## Project Revision Summary")
        console.print_md("---")
        console.print_md("Project Number: **{}**".format(metadata.get("project_number", "N/A")))
        console.print_md("Client: **{}**".format(metadata.get("client_name", "N/A")))
        console.print_md("Project Name: **{}**".format(metadata.get("project_name", "N/A")))
        console.print_md("Report Date: **{}**".format(metadata.get("report_date", "N/A")))
        console.print_md("---")

def get_sheet_info(view_id):
    """Retrieve sheet number and name for a given view ID."""
    sheet = revit.doc.GetElement(view_id)
    if sheet and sheet.LookupParameter("Sheet Number"):
        sheet_number = query.get_param_value(sheet.LookupParameter("Sheet Number"))
        sheet_name = query.get_param_value(sheet.LookupParameter("Sheet Name"))
        return sheet_number, sheet_name
    return None, None

def get_revision_data(clouds, selected_revisions):
    """Group revision clouds by selected revisions."""
    revision_data = {rev.Id.IntegerValue: [] for rev in selected_revisions}
    for cloud in clouds:
        rev_id = cloud.RevisionId.IntegerValue
        if rev_id in revision_data:
            sheet_number, sheet_name = get_sheet_info(cloud.OwnerViewId)
            comment = query.get_param_value(cloud.LookupParameter("Comments"))
            rfi_number = query.get_param_value(cloud.LookupParameter("RFI Number_CEDT"))
            revision_data[rev_id].append({
                "Sheet Number": sheet_number,
                "Sheet Name": sheet_name,
                "Comments": comment,
                "RFI Number": rfi_number
            })
    return revision_data

def deduplicate_clouds(cloud_data):
    """Remove duplicate comments for the same sheet in the revision clouds."""
    seen_sheets_comments = set()
    deduplicated_clouds = []
    for cloud in cloud_data:
        sheet_number = cloud["Sheet Number"] or "N/A"
        comment = cloud["Comments"] or None
        rfi_number = cloud["RFI Number"] or "N/A"
        if not comment or (sheet_number, comment, rfi_number) in seen_sheets_comments:
            continue
        deduplicated_clouds.append(cloud)
        seen_sheets_comments.add((sheet_number, comment, rfi_number))
    return deduplicated_clouds

def print_revision_report(revisions, revision_data):
    """Print the revision report."""
    for rev in revisions:
        revision_number = query.get_param_value(rev.LookupParameter("Revision Number"))
        revision_date = query.get_param_value(rev.LookupParameter("Revision Date"))
        revision_desc = query.get_param_value(rev.LookupParameter("Revision Description"))
        rev_clouds = sorted(revision_data[rev.Id.IntegerValue], key=lambda x: x["Sheet Number"] or "")

        deduplicated_clouds = deduplicate_clouds(rev_clouds)

        console.print_md(
            "### Revision Number: {0} | Date: {1} | Description: {2}".format(revision_number, revision_date, revision_desc)
        )

        table_data = [[cloud["Sheet Number"] or "N/A", cloud["Sheet Name"] or "N/A", cloud["Comments"] or "N/A", cloud["RFI Number"] or "N/A"] for cloud in deduplicated_clouds]

        if table_data:
            console.print_table(table_data, columns=["Sheet Number", "Sheet Name", "Comments", "RFI Number"])
        else:
            console.print_md("No revision clouds with comments found for this revision.")
        console.insert_divider()


# Main logic
all_clouds = DB.FilteredElementCollector(revit.doc)\
    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Select revisions
revisions = forms.select_revisions(button_name='Select Revision', multiple=True)

# Exit if no revisions are selected
if not revisions:
    script.exit()

# Now that revisions are selected, print the header
metadata = get_project_metadata()
print_project_metadata(metadata)

# Collect and print revision data
revision_data = get_revision_data(all_clouds, revisions)
print_revision_report(revisions, revision_data)
