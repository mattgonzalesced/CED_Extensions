# -*- coding: utf-8 -*-
__title__ = "Revision Cloud Issues Report"
__doc__ = """Version = 1.1
Date = 12.05.2024
________________________________________________________________
Description:

This script generates a report on revision clouds with issues, including:
1. Missing comments.
2. Clouds located inside views.
3. Clouds located inside views that are on sheets.

Last Updates:
- [12.05.2024] v1.1 Change Description
________________________________________________________________
Author: AEvelina"""

from pyrevit import script, revit, DB
from pyrevit.revit import query

# Close any currently open output windows
output = script.get_output()
output.close_others()

# Initialize console
console = script.get_output()
console.set_height(800)

def get_revision_cloud_issues(clouds):
    """Identify issues for revision clouds and generate a structured report."""
    issues = []
    doc = revit.doc

    # Process each cloud
    for cloud in clouds:
        comment = query.get_param_value(cloud.LookupParameter("Comments"))
        parent_view = doc.GetElement(cloud.OwnerViewId)
        view_name = query.get_name(parent_view) if parent_view else "Unknown View"

        # Check associated sheets
        sheet_ids = cloud.GetSheetIds()
        sheets = [doc.GetElement(sheet_id) for sheet_id in sheet_ids]

        # Collect sheet numbers and names
        sheet_info = [(query.get_param_value(sheet.LookupParameter("Sheet Number")),
                       query.get_param_value(sheet.LookupParameter("Sheet Name")))
                      for sheet in sheets]

        # Get revision info
        revision = doc.GetElement(cloud.RevisionId)
        revision_number = query.get_param_value(revision.LookupParameter("Revision Number"))
        sequence_number = revision.SequenceNumber

        # Error flags
        missing_comment = not comment
        cloud_in_view = parent_view and not isinstance(parent_view, DB.ViewSheet)
        view_not_on_sheet = parent_view and not sheet_ids

        # Avoid duplicates by processing each cloud only once
        if not sheet_info:
            sheet_info = [(None, None)]

        for sheet_number, sheet_name in sheet_info:
            issues.append({
                "Revision Cloud": cloud.Id,
                "Revision Number": revision_number or "N/A",
                "Sequence Number": sequence_number,
                "Sheet Number": sheet_number or "N/A",
                "Sheet Name": sheet_name or "N/A",
                "View Name": view_name if view_not_on_sheet or cloud_in_view else None,
                "Missing Comment": "X" if missing_comment else "",
                "Cloud in View": "X" if cloud_in_view else "",
                "Not on Sheet": "X" if view_not_on_sheet else ""
            })

    # Sort issues by sequence number and then by sheet number
    return sorted(issues, key=lambda x: (x["Sequence Number"], x["Sheet Number"]))


def print_issues_report(issues):
    """Print the report with structured columns for revision cloud issues."""
    console.print_md("# Revision Cloud Issues Report")
    console.print_md("### Explanation of Error Checks:")
    console.print_md(" **An 'X' in a column indicates that the issue applies to the respective revision cloud.**")
    console.print_md(" - **Missing Comment**: The 'Comments' parameter for the cloud is blank.")
    console.print_md(" - **Cloud in View**: The revision cloud is located in a view, but not directly on a sheet.")
    console.print_md(" - <span style='color:red;'>**Not on Sheet**: The revision cloud is in a view, and that view is not placed on a sheet.</span>")

    # Prepare table data
    table_data = [
        [console.linkify(issue["Revision Cloud"]),
         issue["Revision Number"],
         issue["Sheet Number"],
         issue["Sheet Name"],
         issue["View Name"] or "",
         issue["Missing Comment"],
         issue["Cloud in View"],
         issue["Not on Sheet"]]
        for issue in issues
    ]

    # Print the table
    if table_data:
        console.print_table(
            table_data,
            columns=["Revision Cloud", "Revision Number", "Sheet Number", "Sheet Name", "View Name",
                     "Missing Comment", "Cloud in View", "Not on Sheet"]
        )
    else:
        console.print_md("No issues found with revision clouds.")


# Main logic
all_clouds = DB.FilteredElementCollector(revit.doc)\
    .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Generate the issues report
issues = get_revision_cloud_issues(all_clouds)
print_issues_report(issues)