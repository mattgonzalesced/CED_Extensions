# -*- coding: utf-8 -*-
from pyrevit import script, forms, DB
from pyrevit.framework import Color, SolidColorBrush


WIRE_TYPE_CONFIG_KEY = "wire_type_config"
SETTINGS_XAML = "settings.xaml"
WIRING_TYPES = ["Arc", "Chamfer"]
DEFAULT_HOME_RUN_LENGTH = 4.0


def collect_wire_type_names(doc):
    wire_types = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.WireType).ToElements()
    names = []
    for wire_type in wire_types:
        name_param = wire_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        name = name_param.AsString() if name_param else None
        if name:
            names.append(name)
    return sorted(set(names))


class SettingsWindow(forms.WPFWindow):
    def __init__(self, doc, config):
        xaml_path = script.get_bundle_file(SETTINGS_XAML)
        forms.WPFWindow.__init__(self, xaml_path)

        self.doc = doc
        self.config = config
        self.saved = False
        self._wire_types = collect_wire_type_names(doc)

        self.TitleText.Foreground = SolidColorBrush(Color.FromRgb(0x2A, 0x6F, 0x97))
        self._bind_controls()
        self._load_values()

    def _bind_controls(self):
        self.WireTypeCombo.ItemsSource = self._wire_types
        self.BranchWiringCombo.ItemsSource = WIRING_TYPES
        self.HomeRunWiringCombo.ItemsSource = WIRING_TYPES

    def _safe_select_combo_item(self, combo, value, fallback):
        target = value if value in list(combo.ItemsSource) else fallback
        combo.SelectedItem = target

    def _load_values(self):
        saved_wire_type = getattr(self.config, "default_wire_type", None)
        if self._wire_types:
            fallback_wire = saved_wire_type if saved_wire_type in self._wire_types else self._wire_types[0]
            self.WireTypeCombo.SelectedItem = fallback_wire

        self._safe_select_combo_item(
            self.BranchWiringCombo,
            getattr(self.config, "branch_wiring_type", "Chamfer"),
            "Chamfer"
        )
        self._safe_select_combo_item(
            self.HomeRunWiringCombo,
            getattr(self.config, "homerun_wiring_type", "Arc"),
            "Arc"
        )

        length_value = getattr(self.config, "homerun_length", DEFAULT_HOME_RUN_LENGTH)
        self.HomeRunLengthText.Text = str(length_value)

    def _parse_homerun_length(self):
        text = (self.HomeRunLengthText.Text or "").strip()
        if not text:
            return DEFAULT_HOME_RUN_LENGTH
        try:
            value = float(text)
        except Exception:
            return None
        if value <= 0:
            return None
        return value

    def on_save(self, sender, args):
        homerun_length = self._parse_homerun_length()
        if homerun_length is None:
            forms.alert("Home Run Length must be a positive number.", title="Wire Settings", warn_icon=True)
            return

        if self.WireTypeCombo.SelectedItem:
            self.config.default_wire_type = self.WireTypeCombo.SelectedItem
        if self.BranchWiringCombo.SelectedItem:
            self.config.branch_wiring_type = self.BranchWiringCombo.SelectedItem
        if self.HomeRunWiringCombo.SelectedItem:
            self.config.homerun_wiring_type = self.HomeRunWiringCombo.SelectedItem
        self.config.homerun_length = homerun_length

        script.save_config()
        self.saved = True
        self.Close()

    def on_cancel(self, sender, args):
        self.saved = False
        self.Close()


def main():
    doc = __revit__.ActiveUIDocument.Document
    config = script.get_config(WIRE_TYPE_CONFIG_KEY)
    window = SettingsWindow(doc, config)
    window.ShowDialog()


if __name__ == "__main__":
    main()
