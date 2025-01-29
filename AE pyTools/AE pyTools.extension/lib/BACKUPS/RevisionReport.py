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

# Report metadata
def print_project_metadata():
    """Print project metadata like project number, client, and report date."""
    report_date = coreutils.current_date()
    report_project = revit.query.get_project_info().name
    report_projectno = revit.query.get_project_info().number
    report_client = revit.query.get_project_info().client_name

    # Display the logo with HTML
    console.print_html(
        "<img src='{}' width='150px' style='margin: 0; padding: 0;' />".format(logo_path)
    )

    # Use Markdown to print "Coolsys Energy Design" and "Project Revision Report"
    console.print_md("**Coolsys Energy Design**")
    console.print_md("## Project Revision Summary")

    console.print_md("---")
    console.print_md("Project Number: **{}**".format(report_projectno))
    console.print_md("Client: **{}**".format(report_client))
    console.print_md("Project Name: **{}**".format(report_project))
    console.print_md("Report Date: **{}**".format(report_date))
    console.print_md("---")
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
            revision_data[rev_id].append({
                "Sheet Number": sheet_number,
                "Sheet Name": sheet_name,
                "Comments": comment
            })
    return revision_data

def deduplicate_clouds(cloud_data):
    """Remove duplicate comments for the same sheet in the revision clouds."""
    seen_sheets_comments = set()
    deduplicated_clouds = []
    for cloud in cloud_data:
        sheet_number = cloud["Sheet Number"] or "N/A"
        comment = cloud["Comments"] or None
        if not comment or (sheet_number, comment) in seen_sheets_comments:
            continue
        deduplicated_clouds.append(cloud)
        seen_sheets_comments.add((sheet_number, comment))
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

        table_data = [[cloud["Sheet Number"] or "N/A", cloud["Sheet Name"] or "N/A", cloud["Comments"] or "N/A"] for cloud in deduplicated_clouds]

        if table_data:
            console.print_table(table_data, columns=["Sheet Number", "Sheet Name", "Comments"])
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
print_project_metadata()

# Collect and print revision data
revision_data = get_revision_data(all_clouds, revisions)
print_revision_report(revisions, revision_data)
