# -*- coding: utf-8 -*-
import json
import os

from pyrevit import forms, script
from pyrevit.coreutils import is_url_valid

LINKS_PATH = os.path.join(os.path.dirname(__file__), "links.json")


class LinkWrapper(object):
    def __init__(self, path, label_group=None):
        self.path = path
        self.label_group = label_group

    def __str__(self):
        if self.label_group and self.label_group != "All Links":
            label = self.label_group.rstrip("s")  # singularize
            return "[{}]\t{}".format(label, self.path)
        return self.path


def load_links():
    if os.path.exists(LINKS_PATH):
        with open(LINKS_PATH, 'r') as f:
            return json.load(f)
    return []

def save_links(links):
    with open(LINKS_PATH, 'w') as f:
        json.dump(links, f, indent=2)

def add_link_to_list(entry):
    links = load_links()
    if entry not in links:
        links.append(entry)
        save_links(links)


def post_add_prompt(added_link):
    return forms.alert(
        "Link added:\n\n{}\n\nAdd another?".format(added_link),
        options=["Add Folder", "Add File", "Add URL", "Done"]
    )

def add_folder():
    folder = forms.pick_folder(title="Select Folder to Add")
    if folder:
        add_link_to_list(folder)
        return folder

def add_file():
    file_path = forms.pick_file(title="Select File to Add")
    if file_path:
        add_link_to_list(file_path)
        return file_path

def add_url():
    url = forms.ask_for_string(
        prompt="Enter full URL (must start with http:// or https://)",
        title="Add URL"
    )
    if url:
        url = url.strip()
        if is_url_valid(url):
            add_link_to_list(url)
            return url
        else:
            forms.alert("Invalid URL format. Must start with http:// or https://")

def remove_links():
    group_dict = get_display_groups(label_all=True)

    to_remove = forms.SelectFromList.show(
        group_dict,
        title="Select links to remove",
        multiselect=True,
        name_attr=None,
        group_selector_title="Link Groups",
        button_name="Remove Selected"
    )

    if to_remove:
        selected_paths = [item.path for item in to_remove]
        updated = [l for l in load_links() if l not in selected_paths]

        save_links(updated)

        if len(selected_paths) == 1:
            forms.alert("1 selected link removed.")
        elif len(selected_paths) > 1:
            forms.alert("{} selected links removed.".format(len(selected_paths)))
    else:
        script.exit()


def link_add_loop(initial_action):
    next_action = initial_action

    while next_action != "Done":
        added = None

        if next_action == "Add Folder":
            added = add_folder()
        elif next_action == "Add File":
            added = add_file()
        elif next_action == "Add URL":
            added = add_url()

        if added:
            next_action = post_add_prompt(added)
        else:
            break




def main():
    action = forms.CommandSwitchWindow.show(
        ["Add Folder", "Add File", "Add URL", "Remove Link"],
        message="What would you like to do?"
    )

    if action in ["Add Folder", "Add File", "Add URL"]:
        link_add_loop(action)
    elif action == "Remove Link":
        remove_links()


def get_grouped_links():
    links = load_links()
    grouped = {
        "All Links": links[:],
        "Files": [],
        "Folders": [],
        "URLs": []
    }

    for l in links:
        l = l.strip()
        if is_url_valid(l):
            grouped["URLs"].append(l)
        elif os.path.isfile(l):
            grouped["Files"].append(l)
        elif os.path.isdir(l):
            grouped["Folders"].append(l)

    return grouped

def get_display_groups(label_all=False):
    grouped = get_grouped_links()

    group_dict = {
        "All Links": [],
        "Folders": [],
        "Files": [],
        "URLs": []
    }

    for group_name in ["Folders", "Files", "URLs"]:
        for item in grouped[group_name]:
            wrapper = LinkWrapper(item, label_group=(group_name if label_all else None))
            group_dict[group_name].append(wrapper)
            if label_all:
                group_dict["All Links"].append(wrapper)
            else:
                group_dict["All Links"].append(LinkWrapper(item))  # no prefix here either

    return group_dict


if __name__ == "__main__":
    main()
