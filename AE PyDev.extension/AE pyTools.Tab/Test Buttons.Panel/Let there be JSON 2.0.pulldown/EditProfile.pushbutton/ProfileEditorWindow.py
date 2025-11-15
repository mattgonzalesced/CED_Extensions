# -*- coding: utf-8 -*-
"""
ProfileEditorWindow
-------------------
WPF window for editing CadBlockProfile TypeConfigs at runtime:

- Select a profile (CAD block name)
- Select a TypeConfig (label)
- Edit:
    - Offsets (x/y/z in inches, rotation in degrees)
    - Instance parameter VALUES (names are fixed, user edits values only)
    - Tags (one label per line)

Changes are written back into the existing InstanceConfig objects in memory.
"""

from pyrevit import forms
from Element_Linker import OffsetConfig, TagConfig

try:
    basestring
except NameError:
    basestring = str

# WPF controls for building parameter rows dynamically
from System.Windows.Controls import StackPanel, TextBlock, TextBox, Orientation
from System.Windows import Thickness


class ProfileEditorWindow(forms.WPFWindow):
    def __init__(self, xaml_path, cad_block_profiles):
        self._profiles = cad_block_profiles
        self._current_profile = None
        self._current_typecfg = None

        # cache of (param_name, TextBox) for current type
        self._param_rows = []

        forms.WPFWindow.__init__(self, xaml_path)

        # Fill profile combo
        profile_names = sorted(self._profiles.keys())
        for name in profile_names:
            self.ProfileCombo.Items.Add(name)

        if self.ProfileCombo.Items.Count > 0:
            self.ProfileCombo.SelectedIndex = 0

    # ------------------------------------------------------------------ #
    #  Event handlers
    # ------------------------------------------------------------------ #
    def ProfileCombo_SelectionChanged(self, sender, args):
        """When user picks a profile, populate the Type (label) combo."""
        self.TypeCombo.Items.Clear()
        self._current_profile = None
        self._current_typecfg = None
        self._clear_fields()

        profile_name = self.ProfileCombo.SelectedItem
        if not profile_name:
            return

        profile = self._profiles.get(profile_name)
        if not profile:
            return

        self._current_profile = profile

        # Discover TypeConfig objects by introspecting the profile
        type_list = self._discover_type_configs(profile)

        for tc in type_list:
            lbl = getattr(tc, "label", None)
            if lbl:
                self.TypeCombo.Items.Add(lbl)

        if self.TypeCombo.Items.Count > 0:
            self.TypeCombo.SelectedIndex = 0
        else:
            # Optional tiny hint so you know this profile really has no types
            # (remove this if it gets annoying)
            forms.alert(
                "No TypeConfigs found for profile:\n\n{}".format(profile_name),
                title="Element Linker Profile Editor"
            )

    def TypeCombo_SelectionChanged(self, sender, args):
        """When user picks a type label, load its data into the editor."""
        self._clear_fields()

        if not self._current_profile:
            return

        label = self.TypeCombo.SelectedItem
        if not label:
            return

        # Prefer profile.find_type_by_label if it exists
        type_cfg = None
        if hasattr(self._current_profile, "find_type_by_label"):
            type_cfg = self._current_profile.find_type_by_label(label)

        # Fallback: search discovered type list by label
        if type_cfg is None:
            for tc in self._discover_type_configs(self._current_profile):
                if getattr(tc, "label", None) == label:
                    type_cfg = tc
                    break

        self._current_typecfg = type_cfg

        if not type_cfg:
            return

        inst_cfg = type_cfg.instance_config

        # --- Offsets (use first offset) ---
        offset = inst_cfg.get_offset(0)
        self.OffsetXBox.Text = self._fmt_float(offset.x_inches)
        self.OffsetYBox.Text = self._fmt_float(offset.y_inches)
        self.OffsetZBox.Text = self._fmt_float(offset.z_inches)
        self.OffsetRotBox.Text = self._fmt_float(offset.rotation_deg)

        # --- Parameters (name + editable value rows) ---
        self._param_rows = []
        self.ParamList.Items.Clear()

        params = {}
        # Try method first
        if hasattr(inst_cfg, "get_parameters"):
            params = inst_cfg.get_parameters() or {}
        # Fallback: direct attribute
        if not params and hasattr(inst_cfg, "parameters"):
            params = inst_cfg.parameters or {}

        for name, val in params.items():
            row_panel = StackPanel()
            row_panel.Orientation = Orientation.Horizontal
            row_panel.Margin = Thickness(0, 0, 0, 4)

            name_block = TextBlock()
            name_block.Text = name
            name_block.Width = 200
            name_block.Margin = Thickness(0, 0, 8, 0)

            value_box = TextBox()
            value_box.Text = u"{}".format(val if val is not None else u"")
            value_box.Width = 200

            row_panel.Children.Add(name_block)
            row_panel.Children.Add(value_box)

            self.ParamList.Items.Add(row_panel)
            self._param_rows.append((name, value_box))

        # --- Tags (just labels per line for now) ---
        tags = []
        if hasattr(inst_cfg, "get_tags"):
            tags = inst_cfg.get_tags() or []
        elif hasattr(inst_cfg, "tags"):
            tags = inst_cfg.tags or []

        tag_lines = []
        for tg in tags:
            label_attr = getattr(tg, "label", None)
            if label_attr:
                tag_lines.append(label_attr)
        self.TagTextBox.Text = u"\n".join(tag_lines)

    def OkButton_Click(self, sender, args):
        """Apply edits back into the current TypeConfig's InstanceConfig."""
        if not self._current_profile or not self._current_typecfg:
            self.DialogResult = False
            self.Close()
            return

        inst_cfg = self._current_typecfg.instance_config

        # --- Offsets ---
        try:
            x_in = float(self.OffsetXBox.Text.strip() or "0")
        except Exception:
            x_in = 0.0
        try:
            y_in = float(self.OffsetYBox.Text.strip() or "0")
        except Exception:
            y_in = 0.0
        try:
            z_in = float(self.OffsetZBox.Text.strip() or "0")
        except Exception:
            z_in = 0.0
        try:
            rot_deg = float(self.OffsetRotBox.Text.strip() or "0")
        except Exception:
            rot_deg = 0.0

        offsets = list(getattr(inst_cfg, "offsets", []))
        if not offsets:
            offsets = [OffsetConfig()]

        offsets[0] = OffsetConfig(
            x_inches=x_in,
            y_inches=y_in,
            z_inches=z_in,
            rotation_deg=rot_deg,
        )
        inst_cfg.offsets = offsets

        # --- Parameters: read back values from the UI rows ---
        new_params = {}
        for name, value_box in self._param_rows:
            val_text = (value_box.Text or u"").strip()
            new_params[name] = val_text

        # Write back via attribute or setter
        if hasattr(inst_cfg, "parameters"):
            inst_cfg.parameters = new_params
        elif hasattr(inst_cfg, "set_parameters"):
            inst_cfg.set_parameters(new_params)

        # --- Tags ---
        new_tags = []
        raw_tag_text = self.TagTextBox.Text or u""
        for line in raw_tag_text.splitlines():
            line = line.strip()
            if not line:
                continue
            new_tags.append(TagConfig(label=line))

        if hasattr(inst_cfg, "tags"):
            inst_cfg.tags = new_tags
        elif hasattr(inst_cfg, "set_tags"):
            inst_cfg.set_tags(new_tags)

        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()

    # ------------------------------------------------------------------ #
    #  Helpers
    # ------------------------------------------------------------------ #
    def _clear_fields(self):
        self.OffsetXBox.Text = ""
        self.OffsetYBox.Text = ""
        self.OffsetZBox.Text = ""
        self.OffsetRotBox.Text = ""
        self.ParamList.Items.Clear()
        self._param_rows = []
        self.TagTextBox.Text = ""

    def _fmt_float(self, val):
        try:
            f = float(val)
            return str(round(f, 4))
        except Exception:
            return "0"

    def _discover_type_configs(self, profile):
        """
        Very generic introspection:
        - First look at explicit names (type_configs / types)
        - Then scan profile.__dict__ (including private attrs) for:
            * lists whose items have .label
            * dicts whose values have .label
        """
        # 1) Explicit common names
        for attr_name in ("type_configs", "types"):
            if hasattr(profile, attr_name):
                attr = getattr(profile, attr_name)
                if isinstance(attr, list) and attr:
                    return attr
                if isinstance(attr, dict) and attr:
                    return list(attr.values())

        # 2) Scan __dict__ (includes private attrs)
        type_list = []

        pdict = getattr(profile, "__dict__", {}) or {}
        for attr_name, attr_val in pdict.items():
            if not attr_val:
                continue

            # Lists of TypeConfig
            if isinstance(attr_val, list) and attr_val:
                first = attr_val[0]
                if hasattr(first, "label"):
                    type_list = attr_val
                    break

            # Dicts of TypeConfig
            if isinstance(attr_val, dict) and attr_val:
                vals = list(attr_val.values())
                first = vals[0]
                if hasattr(first, "label"):
                    type_list = vals
                    break

        return type_list
