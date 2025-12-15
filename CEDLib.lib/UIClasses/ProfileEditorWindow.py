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

import copy

from pyrevit import forms
from LogicClasses.Element_Linker import OffsetConfig, TagConfig

try:
    basestring
except NameError:
    basestring = str

# WPF controls for building parameter rows dynamically
from System.Windows.Controls import StackPanel, TextBlock, TextBox, Orientation, ListBoxItem
from System.Windows import Thickness


class ProfileEditorWindow(forms.WPFWindow):
    def __init__(self, xaml_path, cad_block_profiles, relations=None, truth_groups=None, child_to_root=None, delete_callback=None):
        self._profiles = cad_block_profiles
        self._relations = relations or {}
        self._truth_groups = truth_groups or {}
        self._child_to_root = child_to_root or {}
        self._delete_callback = delete_callback
        self._current_profile = None
        self._current_profile_name = None
        self._current_typecfg = None
        self._type_lookup = {}
        self._active_root_key = None

        # cache of (param_name, TextBox) for current type
        self._param_rows = []
        self._tag_rows = []
        self._in_edit_mode = False
        self._force_read_only = False
        self._profile_filter = u""
        self._header_entries = {}
        self._child_entries = {}
        self._group_order = []
        self._display_entries = []

        self._normalize_truth_groups()
        self._rebuild_profile_items()

        forms.WPFWindow.__init__(self, xaml_path)

        self._apply_profile_filter(u"")
        if not getattr(self.ProfileList, "Items", None) or not self.ProfileList.Items.Count:
            self._clear_fields()
        self._set_edit_mode(False)

    # ------------------------------------------------------------------ #
    #  Event handlers
    # ------------------------------------------------------------------ #
    def ProfileList_SelectionChanged(self, sender, args):
        """When user picks a profile, populate the Type list."""
        self._set_edit_mode(False)
        self.TypeList.Items.Clear()
        if hasattr(self, "ChildrenList"):
            self.ChildrenList.Items.Clear()
        if hasattr(self, "ShowChildrenButton"):
            self.ShowChildrenButton.IsEnabled = False
        self._current_profile = None
        self._current_profile_name = None
        self._current_typecfg = None
        self._type_lookup = {}
        self._clear_fields()

        selected_item = getattr(self.ProfileList, "SelectedItem", None)
        metadata = getattr(selected_item, "Tag", None)
        profile_name = metadata.get("profile_name") if metadata else None
        root_key = metadata.get("root_key") if metadata else None
        is_read_only = metadata.get("read_only") if metadata else False
        if not profile_name:
            self.ParentInfoText.Text = "Parent: (none)"
            self._active_root_key = None
            self._force_read_only = False
            self._update_rename_button_state()
            return

        if not root_key:
            root_key = self._child_to_root.get(profile_name, profile_name)

        self._active_root_key = root_key
        self._force_read_only = bool(is_read_only)
        if self._force_read_only:
            self._in_edit_mode = False

        profile = self._profiles.get(profile_name)
        if not profile:
            self.ParentInfoText.Text = "Parent: (none)"
            return

        self._current_profile = profile
        self._current_profile_name = profile_name
        self._update_parent_display()

        # Discover TypeConfig objects by introspecting the profile
        type_list = self._discover_type_configs(profile)

        label_totals = {}
        for tc in type_list:
            lbl = getattr(tc, "label", None) or "<Unnamed>"
            label_totals[lbl] = label_totals.get(lbl, 0) + 1

        label_indices = {}
        for tc in type_list:
            lbl = getattr(tc, "label", None) or "<Unnamed>"
            led_id = getattr(tc, "led_id", None)
            label_indices[lbl] = label_indices.get(lbl, 0) + 1
            display = lbl
            if led_id:
                display = u"{} [{}]".format(lbl, led_id)
            elif label_totals.get(lbl, 0) > 1:
                display = u"{} [#{}]".format(lbl, label_indices[lbl])
            self.TypeList.Items.Add(display)
            self._type_lookup[display] = tc

        if self.TypeList.Items.Count > 0:
            self.TypeList.SelectedIndex = 0
        else:
            forms.alert(
                "No TypeConfigs found for profile:\n\n{}".format(profile_name),
                title="Element Linker Profile Editor"
            )
        self._apply_read_only_state()
        self._refresh_param_buttons()
        self._update_edit_button_state()
        self._update_rename_button_state()

    def TypeList_SelectionChanged(self, sender, args):
        """When user picks a type label, load its data into the editor."""
        self._clear_fields()
        self._set_edit_mode(False)

        if not self._current_profile:
            return

        display_label = self.TypeList.SelectedItem
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
            if hasattr(self, "ShowChildrenButton"):
                self.ShowChildrenButton.IsEnabled = False
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
        if hasattr(inst_cfg, "get_parameters"):
            params = inst_cfg.get_parameters() or {}
        if not params and hasattr(inst_cfg, "parameters"):
            params = inst_cfg.parameters or {}

        for name, val in params.items():
            self._add_param_row(name, val)

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

        if hasattr(self, "ShowChildrenButton"):
            has_children = bool(self._get_children_for_current_profile())
            self.ShowChildrenButton.IsEnabled = bool(has_children and self._current_typecfg)
        if hasattr(self, "ChildrenList"):
            self.ChildrenList.Items.Clear()
        self._apply_read_only_state()

    def EditButton_Click(self, sender, args):
        if self._force_read_only:
            root_key = self._active_root_key or self._child_to_root.get(self._current_profile_name, self._current_profile_name)
            root_display = self._root_display_name(root_key) if root_key else None
            if root_display:
                msg = "Select the non-indented '{}' entry to edit merged profiles.".format(root_display)
            else:
                msg = "Select a source profile to edit."
            forms.alert(msg, title="Element Linker Profile Editor")
            return
        if not self._current_typecfg:
            forms.alert("Select a type before editing.", title="Element Linker Profile Editor")
            return
        if self._in_edit_mode:
            if not self._save_current_typecfg():
                return
            self._mirror_group_profiles(self._active_root_key or self._child_to_root.get(self._current_profile_name))
            self._set_edit_mode(False)
        else:
            self._set_edit_mode(True)

    def ShowChildrenButton_Click(self, sender, args):
        if not hasattr(self, "ChildrenList"):
            return
        if not self._current_typecfg:
            forms.alert("Select a type before showing child equipment.", title="Element Linker Profile Editor")
            return
        self.ChildrenList.Items.Clear()
        children = self._get_children_for_current_profile()
        if not children:
            self.ChildrenList.Items.Add("No child equipment definitions linked to this type.")
            return
        for child in children:
            cid = child.get("id") or ""
            cname = child.get("name") or ""
            if cid and cname:
                self.ChildrenList.Items.Add("{} ({})".format(cname, cid))
            elif cid:
                self.ChildrenList.Items.Add(cid)
            elif cname:
                self.ChildrenList.Items.Add(cname)

    def DeleteTypesButton_Click(self, sender, args):
        """Invoke delete flow from hosting script and refresh lists."""
        if not self._delete_callback:
            forms.alert("Delete logic is not available in this context.", title="Element Linker Profile Editor")
            return
        selection = {
            "profile_name": self._current_profile_name,
            "type_label": getattr(self._current_typecfg, "label", None) if self._current_typecfg else None,
            "type_id": getattr(self._current_typecfg, "element_def_id", None) if self._current_typecfg else None,
            "root_key": self._active_root_key or self._child_to_root.get(self._current_profile_name, self._current_profile_name),
        }
        result = self._delete_callback(selection)
        if not result:
            return
        self._profiles = result.get("profiles", self._profiles)
        self._relations = result.get("relations", self._relations)
        self._truth_groups = result.get("truth_groups", self._truth_groups)
        self._child_to_root = result.get("child_to_root", self._child_to_root)
        self._normalize_truth_groups()
        self._rebuild_profile_items()
        self._apply_profile_filter(self._profile_filter)

    def RenameButton_Click(self, sender, args):
        if self._force_read_only or not self._in_edit_mode:
            forms.alert("Click Edit on a source profile before renaming.", title="Element Linker Profile Editor")
            return
        root_key = self._active_root_key
        if not root_key:
            forms.alert("Select a source profile to rename.", title="Element Linker Profile Editor")
            return
        current_label = self._root_display_name(root_key)
        new_name = forms.ask_for_string(
            prompt="New name for '{}'".format(current_label or ""),
            title="Rename Profile",
            default=current_label or "",
        )
        if new_name is None:
            return
        new_name = (new_name or "").strip()
        if not new_name:
            forms.alert("Profile name cannot be empty.", title="Rename Profile")
            return
        lower = new_name.lower()
        for key, data in self._truth_groups.items():
            if key == root_key:
                continue
            existing = (data.get("display_name") or "").strip().lower()
            if existing == lower:
                forms.alert("A profile named '{}' already exists.".format(new_name), title="Rename Profile")
                return
        group = self._truth_groups.get(root_key)
        if not group:
            return
        group["display_name"] = new_name
        self._group_order = sorted(
            self._truth_groups.keys(),
            key=lambda key: (self._truth_groups[key].get("display_name") or key).lower()
        )
        self._rebuild_profile_items()
        self._apply_profile_filter(self._profile_filter)
        self._update_rename_button_state()

    def OkButton_Click(self, sender, args):
        """Apply edits back into the current TypeConfig's InstanceConfig."""
        if not self._save_current_typecfg():
            return
        if not self._force_read_only:
            self._mirror_group_profiles(self._active_root_key or self._child_to_root.get(self._current_profile_name))
        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()

    # ------------------------------------------------------------------ #
    #  Helpers
    # ------------------------------------------------------------------ #
    def _relation_entry(self):
        if not self._relations:
            return {}
        keys = []
        if self._current_profile_name:
            keys.append(self._current_profile_name)
        if self._current_profile and hasattr(self._current_profile, "cad_name"):
            keys.append(getattr(self._current_profile, "cad_name", ""))
        for key in keys:
            if not key:
                continue
            relation = self._relations.get(key)
            if relation:
                return relation
        return {}

    def _update_parent_display(self):
        relation = self._relation_entry()
        parent_id = (relation.get("parent_id") or "").strip()
        parent_name = (relation.get("parent_name") or "").strip()
        if parent_name and parent_id:
            text = "Parent: {} ({})".format(parent_name, parent_id)
        elif parent_name:
            text = "Parent: {}".format(parent_name)
        elif parent_id:
            text = "Parent: {}".format(parent_id)
        else:
            text = "Parent: (none)"
        if hasattr(self, "ParentInfoText"):
            self.ParentInfoText.Text = text

    def _get_children_for_current_profile(self):
        if not self._current_profile_name:
            return []
        relation = self._relation_entry()
        children = relation.get("children") or []
        anchor_led = None
        if self._current_typecfg:
            anchor_led = getattr(self._current_typecfg, "led_id", None)
            if anchor_led:
                anchor_led = anchor_led.strip()
        cleaned = []
        for child in children:
            if not isinstance(child, dict):
                continue
            cid = (child.get("id") or "").strip()
            cname = (child.get("name") or "").strip()
            child_anchor = (child.get("anchor_led_id") or "").strip()
            if anchor_led and child_anchor and child_anchor != anchor_led:
                continue
            if cid or cname:
                cleaned.append({"id": cid, "name": cname})
        return cleaned

    def _set_edit_mode(self, enabled):
        self._in_edit_mode = bool(enabled) and not self._force_read_only
        self._update_edit_button_state()
        self._update_rename_button_state()
        self._apply_read_only_state()
        self._refresh_param_buttons()

    def _update_edit_button_state(self):
        if not hasattr(self, "EditButton"):
            return
        if self._force_read_only:
            self.EditButton.Content = "Edit"
            self.EditButton.IsEnabled = False
        else:
            self.EditButton.IsEnabled = True
            self.EditButton.Content = "Done" if self._in_edit_mode else "Edit"

    def _update_rename_button_state(self):
        if not hasattr(self, "RenameButton"):
            return
        can_rename = bool(self._in_edit_mode and not self._force_read_only and self._active_root_key)
        self.RenameButton.IsEnabled = can_rename

    def _apply_read_only_state(self):
        read_only = (not self._in_edit_mode) or self._force_read_only
        for textbox in (self.OffsetXBox, self.OffsetYBox, self.OffsetZBox, self.OffsetRotBox):
            textbox.IsReadOnly = read_only
        for entry in self._param_rows:
            entry["value_box"].IsReadOnly = read_only
        for row in self._tag_rows:
            for textbox in row[1:]:
                textbox.IsReadOnly = read_only
        if hasattr(self, "AddTagButton"):
            self.AddTagButton.IsEnabled = not read_only
        if hasattr(self, "RemoveTagButton"):
            self.RemoveTagButton.IsEnabled = not read_only
        self._refresh_param_buttons()
        self._update_rename_button_state()

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
        if hasattr(self, "ChildrenList"):
            self.ChildrenList.Items.Clear()
        if hasattr(self, "ShowChildrenButton"):
            self.ShowChildrenButton.IsEnabled = False
        if hasattr(self, "ParentInfoText") and not self._current_profile_name:
            self.ParentInfoText.Text = "Parent: (none)"
        self._apply_read_only_state()
        self._refresh_param_buttons()

    def _root_group(self, root_key):
        if not root_key:
            return {}
        return self._truth_groups.get(root_key) or {}

    def _root_display_name(self, root_key):
        group = self._root_group(root_key)
        return group.get("display_name") or group.get("source_profile_name") or root_key

    def _root_source_profile(self, root_key):
        group = self._root_group(root_key)
        return group.get("source_profile_name") or root_key

    def _root_members(self, root_key):
        group = self._root_group(root_key)
        members = list(group.get("members") or [])
        source = self._root_source_profile(root_key)
        if source and source not in members:
            members.insert(0, source)
        return members

    def _normalize_truth_groups(self):
        if not self._truth_groups:
            self._truth_groups = {}
        normalized = {}
        child_map = {}
        for root_key, data in self._truth_groups.items():
            display = (data.get("display_name") or data.get("source_profile_name") or root_key or "").strip()
            source_profile = (data.get("source_profile_name") or root_key or "").strip()
            members = data.get("members") or []
            seen = set()
            cleaned = []
            for entry in members:
                name = (entry or "").strip()
                if not name or name in seen:
                    continue
                seen.add(name)
                cleaned.append(name)
            if source_profile and source_profile not in seen:
                cleaned.insert(0, source_profile)
                seen.add(source_profile)
            if not display:
                display = source_profile or root_key
            remainder = [entry for entry in cleaned if entry != source_profile]
            remainder.sort(key=lambda val: val.lower())
            ordered = []
            if source_profile:
                ordered.append(source_profile)
            ordered.extend(remainder)
            normalized[root_key] = {
                "display_name": display,
                "source_profile_name": source_profile or root_key,
                "source_id": data.get("source_id") or root_key,
                "members": ordered,
            }
            for name in ordered:
                if name:
                    child_map[name] = root_key
        for name in sorted(self._profiles.keys()):
            if name not in child_map:
                normalized[name] = {
                    "display_name": name,
                    "source_profile_name": name,
                    "source_id": name,
                    "members": [name],
                }
                child_map[name] = name
        self._truth_groups.clear()
        self._truth_groups.update(normalized)
        self._child_to_root.clear()
        self._child_to_root.update(child_map)
        self._group_order = sorted(
            self._truth_groups.keys(),
            key=lambda key: (self._truth_groups[key].get("display_name") or key).lower()
        )

    def _rebuild_profile_items(self):
        self._header_entries = {}
        self._child_entries = {}
        entries = []
        for root_key in self._group_order:
            group = self._truth_groups.get(root_key) or {}
            display = group.get("display_name") or group.get("source_profile_name") or root_key
            profile_name = group.get("source_profile_name") or root_key
            header = {
                "display": display,
                "profile_name": profile_name,
                "root_key": root_key,
                "display_name": display,
                "read_only": False,
                "is_header": True,
                "key": ("header", root_key),
            }
            self._header_entries[root_key] = header
            entries.append(header)
            members = list(group.get("members") or []) or [profile_name]
            child_list = []
            for idx, member in enumerate(members):
                display_child = u"    - {}".format(member)
                child_entry = {
                    "display": display_child,
                    "profile_name": member,
                    "root_key": root_key,
                    "display_name": member,
                    "read_only": True,
                    "is_header": False,
                    "key": ("child", root_key, member, idx),
                }
                child_list.append(child_entry)
                entries.append(child_entry)
            self._child_entries[root_key] = child_list
        self._display_entries = list(entries)

    def _mirror_group_profiles(self, root_key):
        if not root_key:
            return
        group = self._truth_groups.get(root_key) or {}
        members = list(group.get("members") or [])
        source_name = group.get("source_profile_name") or root_key
        if len(members) <= 1 or not source_name:
            return
        source_profile = self._profiles.get(source_name)
        if not source_profile:
            return
        for member in members:
            if member == source_name:
                continue
            cloned = self._clone_profile_shim(source_profile, member)
            if cloned:
                self._profiles[member] = cloned

    def _clone_profile_shim(self, source_profile, target_name):
        if source_profile is None:
            return None
        try:
            cloned = copy.deepcopy(source_profile)
        except Exception:
            return None
        try:
            cloned.cad_name = target_name
        except Exception:
            pass
        return cloned

    def _populate_profile_list(self, entries, preferred=None):
        if not hasattr(self, "ProfileList"):
            return
        previous = preferred
        if previous is None:
            selected = getattr(self.ProfileList, "SelectedItem", None)
            if selected is not None:
                previous = getattr(selected, "Tag", None)
        self.ProfileList.Items.Clear()
        selected_item = None
        for entry in entries:
            lb_item = ListBoxItem()
            lb_item.Content = entry["display"]
            lb_item.Tag = entry
            self.ProfileList.Items.Add(lb_item)
            if previous and entry.get("key") == previous.get("key"):
                selected_item = lb_item
        if not entries:
            self.ProfileList.SelectedIndex = -1
            return
        if selected_item is None:
            selected_item = self.ProfileList.Items[0]
        self.ProfileList.SelectedItem = selected_item

    def _apply_profile_filter(self, search_text):
        if not hasattr(self, "ProfileList"):
            return
        normalized = (search_text or u"").strip().lower()
        self._profile_filter = normalized
        if not normalized:
            filtered = list(self._display_entries)
        else:
            filtered = []
            for root_key in self._group_order:
                header_entry = self._header_entries.get(root_key)
                child_entries = self._child_entries.get(root_key, [])
                header_label = (header_entry.get("display") if header_entry else self._root_display_name(root_key) or "")
                header_matches = normalized in header_label.lower()
                matching_children = [
                    entry for entry in child_entries
                    if normalized in (entry.get("profile_name") or "").lower()
                ]
                if not header_matches and not matching_children:
                    continue
                if header_entry:
                    filtered.append(header_entry)
                if header_matches:
                    filtered.extend(child_entries)
                else:
                    filtered.extend(matching_children)
        preferred = None
        current_selected = getattr(self.ProfileList, "SelectedItem", None)
        if current_selected is not None:
            tag = getattr(current_selected, "Tag", None)
            if tag:
                preferred = tag
        elif isinstance(self._current_profile_name, basestring):
            root_key = self._child_to_root.get(self._current_profile_name, self._current_profile_name)
            preferred = self._header_entries.get(root_key)
        self._populate_profile_list(filtered, preferred=preferred)
        if not filtered:
            self._current_profile = None
            self._current_profile_name = None
            self._current_typecfg = None
            self._active_root_key = None
            self._force_read_only = False
            self._clear_fields()

    def ProfileSearchBox_TextChanged(self, sender, args):
        text = u""
        if sender is not None:
            text = getattr(sender, "Text", u"") or u""
        self._apply_profile_filter(text)

    def _fmt_float(self, val):
        try:
            f = float(val)
            return str(round(f, 4))
        except Exception:
            return "0"

    def _refresh_param_buttons(self):
        add_enabled = self._in_edit_mode and not self._force_read_only and bool(self._current_typecfg)
        if hasattr(self, "AddParamButton"):
            self.AddParamButton.IsEnabled = add_enabled
        delete_enabled = self._in_edit_mode and not self._force_read_only and bool(getattr(self, "ParamList", None) and self.ParamList.SelectedItem)
        if hasattr(self, "DeleteParamButton"):
            self.DeleteParamButton.IsEnabled = delete_enabled

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
            box.IsReadOnly = not self._in_edit_mode
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

    def AddParamButton_Click(self, sender, args):
        if not self._current_typecfg:
            forms.alert("Select a type before adding parameters.", title="Element Linker Profile Editor")
            return
        if not self._in_edit_mode:
            self._set_edit_mode(True)
        name = forms.ask_for_string(prompt="Parameter name", title="Add Parameter")
        if not name:
            return
        for entry in self._param_rows:
            existing = entry["name"]
            if existing.strip().lower() == name.strip().lower():
                forms.alert("Parameter '{}' already exists.".format(name), title="Add Parameter")
                return
        value = forms.ask_for_string(prompt="Value for '{}'".format(name), title="Add Parameter") or ""
        self._add_param_row(name, value)
        self._apply_read_only_state()

    def _parse_float(self, text_val):
        try:
            return float((text_val or "0").strip())
        except Exception:
            return 0.0

    def _save_current_typecfg(self):
        if not self._current_profile or not self._current_typecfg:
            forms.alert("Select a type before saving.", title="Element Linker Profile Editor")
            return False

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

        # --- Parameters ---
        new_params = {}
        for entry in self._param_rows:
            name = entry["name"]
            value_box = entry["value_box"]
            val_text = (value_box.Text or u"").strip()
            new_params[name] = val_text

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

        return True
    def _add_param_row(self, name, value):
        row_panel = StackPanel()
        row_panel.Orientation = Orientation.Horizontal
        row_panel.Margin = Thickness(0, 0, 0, 4)

        name_block = TextBlock()
        name_block.Text = name
        name_block.Width = 200
        name_block.Margin = Thickness(0, 0, 8, 0)

        value_box = TextBox()
        value_box.Text = u"{}".format(value if value is not None else u"")
        value_box.Width = 200

        row_panel.Children.Add(name_block)
        row_panel.Children.Add(value_box)

        self.ParamList.Items.Add(row_panel)
        self._param_rows.append({
            "name": name,
            "value_box": value_box,
            "panel": row_panel,
        })
        self._apply_read_only_state()
        self.ParamList.SelectedItem = row_panel
        self._refresh_param_buttons()

    def DeleteParamButton_Click(self, sender, args):
        if not self._in_edit_mode:
            forms.alert("Click Edit before deleting parameters.", title="Element Linker Profile Editor")
            return
        if not hasattr(self, "ParamList"):
            return
        selected = self.ParamList.SelectedItem
        if not selected:
            forms.alert("Select a parameter row to delete.", title="Element Linker Profile Editor")
            return
        removed = False
        for idx, entry in enumerate(list(self._param_rows)):
            if entry["panel"] == selected:
                self._param_rows.pop(idx)
                removed = True
                break
        if removed:
            self.ParamList.Items.Remove(selected)
            self.ParamList.SelectedItem = None
        else:
            forms.alert("Could not determine which parameter to delete.", title="Element Linker Profile Editor")
        self._refresh_param_buttons()

    def ParamList_SelectionChanged(self, sender, args):
        self._refresh_param_buttons()
