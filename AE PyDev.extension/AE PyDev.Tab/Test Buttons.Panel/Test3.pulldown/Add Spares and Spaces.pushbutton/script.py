# -*- coding: utf-8 -*-
# Open local Desktop Connector path for active cloud project's "Project Files" folder.

import os

from pyrevit import revit, DB, script, forms

logger = script.get_logger()

def main():
    doc = revit.doc
    model_path = doc.GetWorksharingCentralModelPath()

    if not model_path or model_path.Empty:
        logger.debug("No central model path found (maybe not a workshared model).")
        return

    if not model_path.CloudPath:
        user_visible_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
        logger.debug("Not a cloud path:\n{}".format(user_visible_path))
        return

    # Get user-visible path
    user_visible_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
    logger.debug("User-visible path:\n{}".format(user_visible_path))

    if not user_visible_path.startswith("Autodesk Docs://"):
        logger.debug("Unexpected user-visible path format:\n{}".format(user_visible_path))
        return

    # Remove prefix and split
    cleaned_path = user_visible_path.replace("Autodesk Docs://", "")
    parts = cleaned_path.split("/")

    if len(parts) < 2:
        logger.debug("Unable to extract project name from path:\n{}".format(user_visible_path))
        return

    # Extract project name
    project_name = parts[0]
    logger.debug("Project Name: {}".format(project_name))

    # Build local Desktop Connector path
    user_folder = os.path.expanduser('~')
    base_folder = r"DC\ACCDocs\CoolSys"
    local_project_folder = os.path.join(user_folder, base_folder, project_name, "Project Files")
    logger.debug("Local Project Files path:\n{}".format(local_project_folder))

    # Check existence and open or alert
    if os.path.exists(local_project_folder):
        script.show_folder_in_explorer(local_project_folder)
    else:
        forms.alert(
            "Local folder for project \"{}\" does not exist.\n"
            "Please sync this project with Desktop Connector.".format(project_name),
            title="Desktop Connector Project Missing"
        )

main()
