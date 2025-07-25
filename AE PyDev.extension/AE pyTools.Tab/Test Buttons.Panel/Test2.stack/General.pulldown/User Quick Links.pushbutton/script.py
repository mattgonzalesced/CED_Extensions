# -*- coding: utf-8 -*-
import os

from pyrevit import coreutils
from pyrevit import forms, script

import config  # pulls links from config.py
from config import get_display_groups

logger = script.get_logger()
output = script.get_output()

# Helpers from pyrevit.script module
show_file_in_explorer = script.show_file_in_explorer
show_folder_in_explorer = script.show_folder_in_explorer
open_url = script.open_url



def handle_link(entry):
    path = entry.path.strip()

    if coreutils.is_url_valid(path):
        open_url(path)
    elif os.path.exists(path):
        if os.path.isfile(path):
            show_file_in_explorer(path)
        elif os.path.isdir(path):
            show_folder_in_explorer(path)
        else:
            forms.alert("Path exists but is not a file or folder:\n{}".format(path))
    else:
        forms.alert("This entry doesn’t look valid:\n{}".format(path))


# MAIN

# MAIN
grouped = config.get_grouped_links()

if not grouped["All Links"]:
    forms.alert("No saved links found. Shift-click the button to add some.", exitscript=True)

group_dict = get_display_groups(label_all=True)

selected = forms.SelectFromList.show(group_dict,
                                     title="Pick Link to Open",
                                     multiselect=False,
                                     name_attr=None,
                                     group_selector_title="Link Groups",
                                     button_name="Open",
                                     width=1200
                                     )

if selected:
    handle_link(selected)