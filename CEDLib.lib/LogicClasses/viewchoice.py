# -*- coding: utf-8 -*-
"""
ViewChoice helper for tag view selection UI.
"""


class ViewChoice(object):
    def __init__(self, view):
        self.view = view
        view_type = getattr(view, "ViewType", None)
        self.label = u"{} ({})".format(view.Name, view_type)

    def __repr__(self):
        return self.label
