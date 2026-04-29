# -*- coding: utf-8 -*-
"""
Raw WPF helper using pythonnet.

``pyrevit.forms`` is IronPython-only; this module loads XAML directly
through ``System.Windows.Markup.XamlReader`` so windows work in
CPython 3 + pythonnet (and IronPython 2.7 too, if we ever fall back).

Convention for stage 1: every UI is *modal*. Modeless windows in
Revit need ``ExternalEvent`` plumbing for any Revit-API call from an
event handler — we'll add that selectively later.
"""

import io
import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.IO import StringReader  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Xml import XmlReader  # noqa: E402


def load_xaml(xaml_path_or_text):
    """Load XAML from a file path or a string. Returns the root WPF object.

    The argument is treated as a file path if it points at an existing
    file, otherwise as inline XAML text.
    """
    if os.path.isfile(xaml_path_or_text):
        with io.open(xaml_path_or_text, "r", encoding="utf-8") as f:
            xaml_text = f.read()
    else:
        xaml_text = xaml_path_or_text
    reader = XmlReader.Create(StringReader(xaml_text))
    return XamlReader.Load(reader)


class WpfWindow(object):
    """Base class for stage 1 modal windows.

    Subclasses load a XAML file in ``__init__`` and bind events via
    ``self.find(name)``. Call ``show_modal()`` to display.
    """

    def __init__(self, xaml_path):
        self.window = load_xaml(xaml_path)
        self.result = None  # subclasses populate before closing

    def find(self, name):
        return self.window.FindName(name)

    def show_modal(self):
        self.window.ShowDialog()
        return self.result

    def close(self, result=None):
        if result is not None:
            self.result = result
        self.window.Close()
