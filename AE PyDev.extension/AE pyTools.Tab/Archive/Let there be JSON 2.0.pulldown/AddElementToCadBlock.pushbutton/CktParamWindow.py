# -*- coding: utf-8 -*-
"""
Simple WPF window to capture all CKT_* parameters in one form.
"""

from pyrevit import forms


class CktParamWindow(forms.WPFWindow):
    def __init__(self, xaml_path, title_text):
        forms.WPFWindow.__init__(self, xaml_path)
        self.TitleBlock.Text = title_text

    def OkButton_Click(self, sender, args):
        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()

    def get_values(self):
        return {
            "CKT_Panel_CEDT": (self.PanelBox.Text or u"").strip(),
            "CKT_Circuit Number_CEDT": (self.CircuitBox.Text or u"").strip(),
            "CKT_Rating_CED": (self.RatingBox.Text or u"").strip(),
            "CKT_Load Name_CEDT": (self.LoadNameBox.Text or u"").strip(),
            "CKT_Schedule Notes_CEDT": (self.NotesBox.Text or u"").strip(),
        }
