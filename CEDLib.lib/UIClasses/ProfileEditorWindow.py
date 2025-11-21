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
        self._type_lookup = {}

        # cache of (param_name, TextBox) for current type
        self._param_rows = []
        self._tag_rows = []

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
        self._type_lookup = {}
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

        label_totals = {}
        for tc in type_list:
            lbl = getattr(tc, "label", None) or "<Unnamed>"
            label_totals[lbl] = label_totals.get(lbl, 0) + 1

        label_indices = {}
        for tc in type_list:
            lbl = getattr(tc, "label", None) or "<Unnamed>"
            label_indices[lbl] = label_indices.get(lbl, 0) + 1
            display = lbl
            if label_totals.get(lbl, 0) > 1:
                display = u"{} [#{}]".format(lbl, label_indices[lbl])
            self.TypeCombo.Items.Add(display)
            self._type_lookup[display] = tc

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

        display_label = self.TypeCombo.SelectedItem
        if not display_label:
            return

        type_cfg = self._type_lookup.get(display_label)
        label = getattr(type_cfg, "label", None) if type_cfg else None

        if type_cfg is None and display_label:
            label = display_label

        if type_cfg is None and label:
            if hasattr(self._current_profile, "find_type_by_label"):
                type_cfg = self._current_profile.find_type_by_label(label)

        if type_cfg is None and label:
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

        # --- Tags: build editable rows ---
        tags = []
        if hasattr(inst_cfg, "get_tags"):
            tags = inst_cfg.get_tags() or []
        elif hasattr(inst_cfg, "tags"):
            tags = inst_cfg.tags or []
        self.TagList.Items.Clear()
        self._tag_rows = []
        if tags:
            for tg in tags:
                self._add_tag_row(tg)
        else:
            self._add_tag_row()

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
        for row in self._tag_rows:
            _, family_box, type_box, x_box, y_box, z_box, rot_box = row
            family = (family_box.Text or u"").strip()
            type_name = (type_box.Text or u"").strip() or None
            if not family and not type_name:
                continue
            tag_offset = OffsetConfig(
                x_inches=self._parse_float(x_box.Text),
                y_inches=self._parse_float(y_box.Text),
                z_inches=self._parse_float(z_box.Text),
                rotation_deg=self._parse_float(rot_box.Text),
            )
            new_tags.append(
                TagConfig(
                    category_name='Annotation Symbols',
                    family_name=family,
                    type_name=type_name,
                    offsets=tag_offset,
                )
            )

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
        if hasattr(self, "TagList"):
            self.TagList.Items.Clear()
        self._tag_rows = []

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

    # ------------------------------------------------------------------ #
    #  Tag helpers / events
    # ------------------------------------------------------------------ #
    def _add_tag_row(self, tag=None):
        if not hasattr(self, "TagList"):
            return
        panel = StackPanel(Orientation=Orientation.Horizontal, Margin=Thickness(0, 0, 0, 5))

        def _make_field(label_text, width):
            container = StackPanel(Margin=Thickness(0, 0, 5, 0))
            container.Width = width
            lbl = TextBlock(Text=label_text, Margin=Thickness(0, 0, 0, 2))
            box = TextBox()
            container.Children.Add(lbl)
            container.Children.Add(box)
            panel.Children.Add(container)
            return box

        family_box = _make_field("Family", 150.0)
        type_box = _make_field("Type", 140.0)
        x_box = _make_field("X (in)", 80.0)
        y_box = _make_field("Y (in)", 80.0)
        z_box = _make_field("Z (in)", 80.0)
        rot_box = _make_field("Rot (deg)", 90.0)

        fam, typ, offsets = self._resolve_tag_parts(tag)
        if fam:
            family_box.Text = fam
        if typ:
            type_box.Text = typ
        if offsets:
            x_box.Text = self._fmt_float(offsets[0])
            y_box.Text = self._fmt_float(offsets[1])
            z_box.Text = self._fmt_float(offsets[2])
            rot_box.Text = self._fmt_float(offsets[3])

        self.TagList.Items.Add(panel)
        self._tag_rows.append((panel, family_box, type_box, x_box, y_box, z_box, rot_box))

    def _resolve_tag_parts(self, tag):
        if not tag:
            return (None, None, (0.0, 0.0, 0.0, 0.0))

        if isinstance(tag, dict):
            family = tag.get("family_name") or tag.get("family")
            typ = tag.get("type_name") or tag.get("type")
            offsets = tag.get("offsets") or {}
            return (
                family,
                typ,
                (
                    offsets.get("x_inches", 0.0),
                    offsets.get("y_inches", 0.0),
                    offsets.get("z_inches", 0.0),
                    offsets.get("rotation_deg", 0.0),
                ),
            )

        family = getattr(tag, "family_name", None) or getattr(tag, "family", None)
        typ = getattr(tag, "type_name", None) or getattr(tag, "type", None)
        offsets_obj = getattr(tag, "offsets", None)
        if offsets_obj is None:
            return (family, typ, (0.0, 0.0, 0.0, 0.0))
        return (
            family,
            typ,
            (
                getattr(offsets_obj, "x_inches", 0.0),
                getattr(offsets_obj, "y_inches", 0.0),
                getattr(offsets_obj, "z_inches", 0.0),
                getattr(offsets_obj, "rotation_deg", 0.0),
            ),
        )

    def AddTagButton_Click(self, sender, args):
        self._add_tag_row()

    def RemoveTagButton_Click(self, sender, args):
        selected = self.TagList.SelectedItem
        if not selected:
            return
        for idx, row in enumerate(list(self._tag_rows)):
            panel = row[0]
            if panel == selected:
                self._tag_rows.pop(idx)
                break
        self.TagList.Items.Remove(selected)

    def _parse_float(self, text_val):
        try:
            return float((text_val or "0").strip())
        except Exception:
            return 0.0
