# -*- coding: utf-8 -*-
__title__ = "About"
__doc__ = "Display toolbar version and author information."

import os
from pyrevit import forms


def read_about_yaml(yaml_path):
    version = None
    authors = []

    with open(yaml_path, "r") as yaml_file:
        for raw_line in yaml_file:
            stripped = raw_line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("toolbar_version:"):
                version = stripped.split(":", 1)[1].strip()
                continue
            if stripped.startswith("- "):
                authors.append(stripped[2:].strip())

    return version, authors


def main():
    yaml_path = os.path.join(os.path.dirname(__file__), "about.yaml")

    if not os.path.exists(yaml_path):
        forms.alert("Could not find about.yaml.", title="About", ok=True)
        return

    try:
        version, authors = read_about_yaml(yaml_path)
    except Exception as err:
        forms.alert("Failed to read about.yaml:\n{}".format(err), title="About", ok=True)
        return

    if not version:
        version = "Unknown"

    if not authors:
        authors = ["No authors listed"]

    author_lines = "\n".join("- {}".format(name) for name in authors)
    message = "The version of the toolbar is {}\n\nAuthors:\n{}".format(version, author_lines)
    forms.alert(message, title="About", ok=True)


if __name__ == "__main__":
    main()
