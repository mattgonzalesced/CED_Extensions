# -*- coding: utf-8 -*-
"""
Minimal file-dialog and message-box helpers that work in both pyRevit
engines.

This module is intentionally Windows.Forms-only — it does NOT load WPF.
That keeps Import/Export light, and any module-load failure in WPF
won't propagate to scripts that just need file dialogs and alerts.

WPF-based prompts (single-line text input, list picker) live in
``wpf_dialogs.py``; import that explicitly from a script when needed.
"""

import os

import clr  # noqa: F401  -- pythonnet bridge

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import (  # noqa: E402
    DialogResult,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
    OpenFileDialog,
    SaveFileDialog,
)


def _build_filter(file_ext):
    if not file_ext:
        return "All files (*.*)|*.*"
    ext = file_ext.lstrip(".")
    return "{0} files (*.{0})|*.{0}|All files (*.*)|*.*".format(ext)


def pick_file(file_ext=None, title=None, init_dir=None):
    """Show an open-file dialog. Returns the selected path or ``None``."""
    dlg = OpenFileDialog()
    if title:
        dlg.Title = title
    dlg.Filter = _build_filter(file_ext)
    if init_dir and os.path.isdir(init_dir):
        dlg.InitialDirectory = init_dir
    dlg.Multiselect = False
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None


def save_file(file_ext=None, title=None, default_name=None, init_dir=None):
    """Show a save-file dialog. Returns the selected path or ``None``."""
    dlg = SaveFileDialog()
    if title:
        dlg.Title = title
    dlg.Filter = _build_filter(file_ext)
    if init_dir and os.path.isdir(init_dir):
        dlg.InitialDirectory = init_dir
    if default_name:
        dlg.FileName = default_name
        if file_ext:
            dlg.DefaultExt = file_ext.lstrip(".")
            dlg.AddExtension = True
    dlg.OverwritePrompt = True
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None


def alert(message, title=None, **_ignored):
    """Show an informational message box.

    Extra keyword arguments (``exitscript``, etc.) accepted by
    ``pyrevit.forms.alert`` are accepted and ignored so caller code can
    use either implementation interchangeably.
    """
    MessageBox.Show(
        str(message) if message is not None else "",
        str(title) if title else "",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information,
    )


def confirm(message, title=None):
    """Yes/No prompt. Returns True on Yes."""
    result = MessageBox.Show(
        str(message) if message is not None else "",
        str(title) if title else "",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question,
    )
    return result == DialogResult.Yes
