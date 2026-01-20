# -*- coding: utf-8 -*-
"""
Place Single Profile (Dockable Pane)
------------------------------------
Opens the dockable pane for placing a single equipment definition.
"""

from pyrevit import forms

from PlaceSingleProfilePanel import ensure_panel_visible  # noqa: E402


def main():
    try:
        ensure_panel_visible()
    except Exception as exc:
        forms.alert("Failed to open Place Single Profile pane:\n\n{}".format(exc), title="Place Single Profile")


if __name__ == "__main__":
    main()
