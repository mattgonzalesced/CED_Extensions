# -*- coding: utf-8 -*-
"""
Manage Profiles editor (option B scope: view + delete + structural edits).

Mutations:
    * Rename a profile
    * Edit parent_filter triple + the three boolean flags
    * Add / remove LEDs (an Add LED prompts the caller to re-run capture
      on a child element — so this just exposes the entry point)
    * Edit a LED's parameters (key/value grid)
    * Edit the first offset entry numerically
    * Delete a whole profile

The window mutates the supplied ``profile_data`` dict in place; the
caller is responsible for committing it back via
``active_yaml.save_active_payload``.
"""

import copy
import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    TreeViewItem,
)

import wpf as _wpf
import profile_model
import truth_groups
import merge_workflow
import wpf_dialogs


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "ManageProfilesWindow.xaml"
)


class _ProfileItem(object):
    """ListBox display wrapper."""

    def __init__(self, profile_dict):
        self.data = profile_dict

    @property
    def DisplayName(self):
        return "{}  ({})".format(
            self.data.get("name") or "(unnamed)",
            self.data.get("id") or "??",
        )


class _ParamRow(object):

    def __init__(self, name="", value=""):
        self.Name = name
        self.Value = value


def _coerce_float(text, default=0.0):
    try:
        return float(str(text).strip())
    except (ValueError, TypeError):
        return default


class ManageProfilesController(object):
    """Modal editor. Pass the profile-document dict; mutations land in it."""

    def __init__(self, profile_data, doc=None):
        self.profile_data = profile_data
        self.doc = doc  # active Revit doc; used by Change Type to enumerate FamilySymbols
        self.dirty = False
        # If the user clicks "Add LED..." we close the modal and signal
        # the caller to run the append workflow on this profile.
        self.requested_add_to_profile_id = None
        # Cross-profile fixture clipboard (deep-copied LED dict).
        self._fixture_clipboard = None
        self.window = _wpf.load_xaml(_XAML_PATH)

        self._profile_items = ObservableCollection[_NetObject]()
        self._reload_profile_list()

        self._lookup_controls()
        self._wire_events()
        self._set_status("Ready")
        self._load_profile(None)

    # ---- bootstrapping ------------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.profile_list = f("ProfileList")
        self.profile_list.ItemsSource = self._profile_items
        self.filter_box = f("FilterBox")
        self.delete_btn = f("DeleteProfileButton")
        self.id_label = f("IdLabel")
        self.name_box = f("NameBox")
        self.cat_box = f("CategoryBox")
        self.fam_box = f("FamilyBox")
        self.type_box = f("TypeBox")
        self.allow_parentless = f("AllowParentlessCheck")
        self.allow_unmatched = f("AllowUnmatchedCheck")
        self.prompt_mismatch = f("PromptMismatchCheck")
        self.led_tree = f("LedTree")
        self.add_led_btn = f("AddLedButton")
        self.remove_led_btn = f("RemoveLedButton")
        self.change_type_btn = f("ChangeTypeButton")
        self.copy_fixture_btn = f("CopyFixtureButton")
        self.paste_fixture_btn = f("PasteFixtureButton")
        self.relationship_box = f("RelationshipBox")
        self.profile_meta_grid = f("ProfileMetaGrid")
        self.add_profile_meta_btn = f("AddProfileMetaButton")
        self.remove_profile_meta_btn = f("RemoveProfileMetaButton")
        self.save_profile_meta_btn = f("SaveProfileMetaButton")
        self.alias_list = f("AliasList")
        self.alias_info_label = f("AliasInfoLabel")
        self.add_alias_btn = f("AddAliasButton")
        self.remove_alias_btn = f("RemoveAliasButton")
        self.save_aliases_btn = f("SaveAliasesButton")
        self.parameter_grid = f("ParameterGrid")
        self.offset_x = f("OffsetXBox")
        self.offset_y = f("OffsetYBox")
        self.offset_z = f("OffsetZBox")
        self.offset_rot = f("OffsetRotBox")
        self.save_btn = f("SaveButton")
        self.save_selected_btn = f("SaveSelectedButton")
        self.selected_header_label = f("SelectedHeaderLabel")
        self.close_btn = f("CloseButton")
        self.status = f("StatusLabel")

    def _wire_events(self):
        self.profile_list.SelectionChanged += self._on_profile_selected
        self.filter_box.TextChanged += self._on_filter_changed
        self.delete_btn.Click += self._on_delete_profile
        self.add_led_btn.Click += self._on_add_led_clicked
        self.remove_led_btn.Click += self._on_remove_led
        self.change_type_btn.Click += self._on_change_type
        self.copy_fixture_btn.Click += self._on_copy_fixture
        self.paste_fixture_btn.Click += self._on_paste_fixture
        self.led_tree.SelectedItemChanged += self._on_led_selected
        self.save_btn.Click += self._on_save
        self.save_selected_btn.Click += self._on_save_selected
        self.save_profile_meta_btn.Click += self._on_save_profile_metadata
        self.add_profile_meta_btn.Click += self._on_add_profile_metadata
        self.remove_profile_meta_btn.Click += self._on_remove_profile_metadata
        self.add_alias_btn.Click += self._on_add_alias
        self.remove_alias_btn.Click += self._on_remove_alias
        self.save_aliases_btn.Click += self._on_save_aliases
        self.close_btn.Click += self._on_close

    # ---- list -------------------------------------------------------

    def _reload_profile_list(self, filter_text=""):
        self._profile_items.Clear()
        defs = self.profile_data.get("equipment_definitions") or []
        text = (filter_text or "").strip().lower()
        # Display alphabetically (case-insensitive) by profile name. The
        # underlying ``equipment_definitions`` list order is not changed,
        # so import / export order stays stable.
        sorted_defs = sorted(
            (e for e in defs if isinstance(e, dict)),
            key=lambda e: (e.get("name") or "").lower(),
        )
        for entry in sorted_defs:
            if text and text not in (entry.get("name") or "").lower():
                continue
            self._profile_items.Add(_ProfileItem(entry))

    # ---- selection / load -------------------------------------------

    def _selected_profile(self):
        item = self.profile_list.SelectedItem
        if item is None:
            return None
        return item.data

    def _load_profile(self, profile):
        self._current_profile = profile
        self._current_set = None
        self._current_led = None
        if profile is None:
            self.id_label.Text = ""
            self.name_box.Text = ""
            self.cat_box.Text = ""
            self.fam_box.Text = ""
            self.type_box.Text = ""
            self.allow_parentless.IsChecked = False
            self.allow_unmatched.IsChecked = False
            self.prompt_mismatch.IsChecked = False
            self.profile_meta_grid.ItemsSource = ObservableCollection[_NetObject]()
            self._populate_aliases(None)
            self.led_tree.Items.Clear()
            self._load_led(None)
            return

        wrapper = profile_model.Profile(profile)
        self.id_label.Text = wrapper.id or ""
        self.name_box.Text = wrapper.name or ""
        pf = wrapper.parent_filter
        self.cat_box.Text = pf.category or ""
        self.fam_box.Text = pf.family_name_pattern or ""
        self.type_box.Text = pf.type_name_pattern or ""
        self.allow_parentless.IsChecked = bool(wrapper.allow_parentless)
        self.allow_unmatched.IsChecked = bool(wrapper.allow_unmatched_parents)
        self.prompt_mismatch.IsChecked = bool(wrapper.prompt_on_parent_mismatch)
        self._populate_profile_metadata(profile)
        self._populate_aliases(profile)
        self._populate_led_tree(profile)
        self._load_led(None)

    def _populate_led_tree(self, profile):
        self.led_tree.Items.Clear()
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            set_node = TreeViewItem()
            set_node.Header = "{}  ({})".format(set_dict.get("name") or "set",
                                                set_dict.get("id") or "")
            set_node.Tag = ("set", set_dict)
            set_node.IsExpanded = True
            for led in set_dict.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                led_node = TreeViewItem()
                led_node.Header = "{}  ({})".format(led.get("label") or "led",
                                                    led.get("id") or "")
                led_node.Tag = ("led", set_dict, led)
                led_node.IsExpanded = True
                # Nest annotations (new schema) AND legacy tags/keynotes/notes.
                for ann in self._iter_led_annotations(led):
                    ann_node = TreeViewItem()
                    label = ann.get("label") or ann.get("type_name") or ann.get("kind") or "annotation"
                    kind = ann.get("kind") or ""
                    ann_node.Header = "[{}] {}  ({})".format(
                        kind or "ann",
                        label,
                        ann.get("id") or "ANN-?",
                    )
                    ann_node.Tag = ("ann", set_dict, led, ann)
                    led_node.Items.Add(ann_node)
                set_node.Items.Add(led_node)
            self.led_tree.Items.Add(set_node)

    def _iter_led_annotations(self, led):
        """Yield annotation dicts in tree-display order. Prefers the new
        ``annotations`` list; falls back to legacy peer lists if absent."""
        anns = led.get("annotations")
        if isinstance(anns, list) and anns:
            for a in anns:
                if isinstance(a, dict):
                    yield a
            return
        for kind, key in (("tag", "tags"), ("keynote", "keynotes"),
                          ("text_note", "text_notes")):
            for entry in (led.get(key) or []):
                if not isinstance(entry, dict):
                    continue
                clone = dict(entry)
                clone.setdefault("kind", kind)
                yield clone

    def _load_led(self, payload):
        # Reset state first.
        self._current_set = None
        self._current_led = None
        self._current_ann = None
        self.parameter_grid.ItemsSource = ObservableCollection[_NetObject]()
        self.offset_x.Text = ""
        self.offset_y.Text = ""
        self.offset_z.Text = ""
        self.offset_rot.Text = ""
        self._update_selected_header(None)

        if not payload:
            return

        kind = payload[0]
        if kind == "set":
            self._current_set = payload[1]
            self._update_selected_header(
                "Set: {}".format(payload[1].get("id") or "?")
            )
            return

        if kind == "led":
            self._current_set = payload[1]
            self._current_led = payload[2]
            self._populate_param_grid(self._current_led.setdefault("parameters", {}))
            offsets = self._current_led.setdefault("offsets", [])
            if not offsets:
                offsets.append({"x_inches": 0.0, "y_inches": 0.0,
                                "z_inches": 0.0, "rotation_deg": 0.0})
            self._populate_offset_inputs(offsets[0])
            self._update_selected_header(
                "LED: {} ({})  -  offsets relative to PARENT".format(
                    self._current_led.get("label") or "(no label)",
                    self._current_led.get("id") or "?",
                )
            )
            return

        if kind == "ann":
            self._current_set = payload[1]
            self._current_led = payload[2]
            self._current_ann = payload[3]
            self._populate_param_grid(self._current_ann.setdefault("parameters", {}))
            offset_dict = self._current_ann.setdefault(
                "offsets",
                {"x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0},
            )
            # Legacy tag entries kept their offsets as a dict; new schema
            # also uses a dict — but we tolerate a one-element list.
            if isinstance(offset_dict, list):
                offset_dict = offset_dict[0] if offset_dict else {}
                self._current_ann["offsets"] = offset_dict
            self._populate_offset_inputs(offset_dict)
            self._update_selected_header(
                "Annotation: {} [{}] ({})  -  offsets relative to FIXTURE".format(
                    self._current_ann.get("label") or self._current_ann.get("kind") or "annotation",
                    self._current_ann.get("kind") or "?",
                    self._current_ann.get("id") or "?",
                )
            )
            return

    def _update_selected_header(self, label):
        if hasattr(self, "selected_header_label") and self.selected_header_label is not None:
            self.selected_header_label.Text = label or "Selected LED / annotation"
        self._populate_relationship_box()

    def _populate_relationship_box(self):
        """Show the profile's parent_filter + any active directives on the
        currently loaded LED / ANN."""
        if not hasattr(self, "relationship_box") or self.relationship_box is None:
            return
        lines = []
        profile = self._current_profile
        if profile is not None:
            pf = profile.get("parent_filter") or {}
            lines.append("PROFILE PARENT FILTER")
            lines.append("  category: {}".format(pf.get("category") or "(any)"))
            lines.append("  family:   {}".format(pf.get("family_name_pattern") or "(any)"))
            lines.append("  type:     {}".format(pf.get("type_name_pattern") or "(any)"))

        if self._current_led is not None:
            lines.append("")
            lines.append("LED  {}  ({})".format(
                self._current_led.get("label") or "(no label)",
                self._current_led.get("id") or "?",
            ))
            params = self._current_led.get("parameters") or {}
            directive_lines = self._format_directive_lines(params)
            if directive_lines:
                lines.append("  directives:")
                lines.extend("    " + ln for ln in directive_lines)
            else:
                lines.append("  directives: (none — all parameters static)")

        if self._current_ann is not None:
            lines.append("")
            lines.append("ANNOTATION  {}  ({})".format(
                self._current_ann.get("label")
                or self._current_ann.get("kind") or "annotation",
                self._current_ann.get("id") or "?",
            ))
            lines.append("  kind:     {}".format(self._current_ann.get("kind") or "?"))
            lines.append("  attached to LED {}".format(
                self._current_led.get("id") if self._current_led else "?"
            ))
            lines.append("  offsets relative to the host fixture, NOT the parent")

        self.relationship_box.Text = "\n".join(lines) if lines else ""

    def _format_directive_lines(self, params):
        out = []
        for name, value in (params or {}).items():
            if not isinstance(value, dict):
                continue
            if "parent_parameter" in value:
                out.append("{}  ->  parent.{}".format(
                    name, value.get("parent_parameter") or "?"
                ))
            elif "sibling_parameter" in value:
                out.append("{}  ->  sibling {}".format(
                    name, value.get("sibling_parameter") or "?"
                ))
        return out

    def _populate_param_grid(self, params_dict):
        rows = ObservableCollection[_NetObject]()
        for k, v in (params_dict or {}).items():
            rows.Add(_ParamRow(str(k), "" if v is None else str(v)))
        self.parameter_grid.ItemsSource = rows

    def _populate_offset_inputs(self, offset_dict):
        self.offset_x.Text = str((offset_dict or {}).get("x_inches") or 0.0)
        self.offset_y.Text = str((offset_dict or {}).get("y_inches") or 0.0)
        self.offset_z.Text = str((offset_dict or {}).get("z_inches") or 0.0)
        self.offset_rot.Text = str((offset_dict or {}).get("rotation_deg") or 0.0)

    # ---- event handlers ---------------------------------------------

    def _on_filter_changed(self, sender, e):
        self._reload_profile_list(self.filter_box.Text)

    def _on_profile_selected(self, sender, e):
        profile = self._selected_profile()
        self._load_profile(profile)

    def _on_led_selected(self, sender, e):
        node = self.led_tree.SelectedItem
        if node is None:
            self._load_led(None)
            return
        self._load_led(node.Tag)

    def _on_delete_profile(self, sender, e):
        profile = self._selected_profile()
        if profile is None:
            return
        defs = self.profile_data.get("equipment_definitions") or []
        try:
            defs.remove(profile)
        except ValueError:
            return
        self.dirty = True
        self._reload_profile_list(self.filter_box.Text)
        self._load_profile(None)
        self._set_status("Profile deleted (unsaved)")

    def _on_add_led_clicked(self, sender, e):
        if self._current_profile is None:
            self._set_status("Pick a profile first")
            return
        profile_id = self._current_profile.get("id")
        if not profile_id:
            self._set_status("Selected profile has no id; cannot add to it")
            return
        # Persist any in-progress edits before handing control back.
        self._save_profile_metadata_to_data()
        self._save_selected_to_data()
        self.dirty = True
        # Tell the caller to run append_workflow on this profile, then close.
        self.requested_add_to_profile_id = profile_id
        self.window.Close()

    def _on_remove_led(self, sender, e):
        # If an annotation is currently selected, remove it from its LED.
        if self._current_ann is not None and self._current_led is not None:
            anns = self._current_led.get("annotations")
            if isinstance(anns, list):
                try:
                    anns.remove(self._current_ann)
                    self.dirty = True
                    self._populate_led_tree(self._current_profile)
                    self._load_led(None)
                    self._set_status("Annotation removed (unsaved)")
                    return
                except ValueError:
                    pass
            # Legacy schema fallback: try peer lists.
            for key in ("tags", "keynotes", "text_notes"):
                lst = self._current_led.get(key)
                if isinstance(lst, list) and self._current_ann in lst:
                    lst.remove(self._current_ann)
                    self.dirty = True
                    self._populate_led_tree(self._current_profile)
                    self._load_led(None)
                    self._set_status("Annotation removed (unsaved)")
                    return
            self._set_status("Could not locate annotation to remove")
            return
        # Otherwise remove the LED itself.
        if self._current_led is None or self._current_set is None:
            return
        leds = self._current_set.get("linked_element_definitions") or []
        try:
            leds.remove(self._current_led)
        except ValueError:
            return
        self.dirty = True
        self._populate_led_tree(self._current_profile)
        self._load_led(None)
        self._set_status("LED removed (unsaved)")

    # ---- change type / copy / paste ---------------------------------

    def _on_change_type(self, sender, e):
        """Pick a new ``Family : Type`` from a dropdown of types loaded in
        the active project, filtered to the same category as the
        currently selected LED / annotation.

        For text-note annotations the dropdown shows TextNoteType names
        (no family).
        """
        if self.doc is None:
            self._set_status(
                "No active Revit document — cannot enumerate types. "
                "Run Manage Profiles from the panel to enable this."
            )
            return

        if self._current_ann is not None:
            kind = self._current_ann.get("kind") or ""
            if kind == "text_note":
                options = self._enumerate_text_note_types()
                if not options:
                    self._set_status("No TextNoteType is loaded in the project.")
                    return
                chosen = wpf_dialogs.pick_from_list(
                    options,
                    title="Change Type",
                    prompt="Pick a TextNoteType:",
                )
                if not chosen:
                    return
                self._current_ann["type_name"] = chosen
                # text-note label is the text content, not Family:Type;
                # leave label / family_name alone.
            else:
                options = self._enumerate_all_family_types()
                if not options:
                    self._set_status("No FamilySymbols loaded in the project.")
                    return
                chosen = wpf_dialogs.pick_from_list(
                    options,
                    title="Change Type",
                    prompt="Pick a Family : Type for the annotation:",
                )
                if not chosen:
                    return
                if " : " in chosen:
                    fam, typ = chosen.split(" : ", 1)
                    self._current_ann["family_name"] = fam.strip()
                    self._current_ann["type_name"] = typ.strip()
                else:
                    self._current_ann["family_name"] = chosen
                    self._current_ann["type_name"] = ""
                self._current_ann["label"] = chosen
            self.dirty = True
            if self._current_profile is not None:
                self._populate_led_tree(self._current_profile)
            self._set_status("Annotation type changed (unsaved)")
            return

        if self._current_led is not None:
            options = self._enumerate_all_family_types()
            if not options:
                self._set_status("No FamilySymbols loaded in the project.")
                return
            chosen = wpf_dialogs.pick_from_list(
                options,
                title="Change Type",
                prompt="Pick a Family : Type for the LED:",
            )
            if not chosen:
                return
            self._current_led["label"] = chosen
            self.dirty = True
            if self._current_profile is not None:
                self._populate_led_tree(self._current_profile)
            self._set_status("LED type changed (unsaved)")
            return

        self._set_status("Pick a LED or annotation first")

    def _enumerate_all_family_types(self):
        """Return a sorted list of every loaded FamilySymbol as a
        ``"Family : Type"`` string. No category filter — the user
        explicitly wants the full menu so cross-category re-typing
        (e.g. swapping a fixture for a tag) is possible from one place.
        """
        if self.doc is None:
            return []
        from Autodesk.Revit.DB import FamilySymbol, FilteredElementCollector
        out = set()
        for sym in FilteredElementCollector(self.doc).OfClass(FamilySymbol):
            family = getattr(sym, "Family", None)
            if family is None:
                continue
            fam_name = getattr(family, "Name", "") or ""
            type_name = getattr(sym, "Name", "") or ""
            if not fam_name or not type_name:
                continue
            out.add("{} : {}".format(fam_name, type_name))
        return sorted(out, key=lambda s: s.lower())

    def _enumerate_text_note_types(self):
        if self.doc is None:
            return []
        from Autodesk.Revit.DB import TextNoteType, FilteredElementCollector
        out = set()
        for t in FilteredElementCollector(self.doc).OfClass(TextNoteType):
            name = getattr(t, "Name", "") or ""
            if name:
                out.add(name)
        return sorted(out, key=lambda s: s.lower())

    def _on_copy_fixture(self, sender, e):
        if self._current_led is None:
            self._set_status("Pick a fixture (LED) to copy")
            return
        # Persist any pending edits to the parameter grid / offsets first.
        self._save_selected_to_data()
        self._fixture_clipboard = copy.deepcopy(self._current_led)
        label = self._fixture_clipboard.get("label") or "?"
        n_anns = len(self._fixture_clipboard.get("annotations") or [])
        self._set_status(
            "Copied fixture {!r} (with {} annotation(s)). "
            "Switch to the destination profile and click Paste fixture.".format(
                label, n_anns
            )
        )

    def _on_paste_fixture(self, sender, e):
        if self._fixture_clipboard is None:
            self._set_status("Nothing to paste — copy a fixture first")
            return
        if self._current_profile is None:
            self._set_status("Pick a destination profile first")
            return

        sets = self._current_profile.get("linked_sets") or []
        if not sets:
            new_set = {
                "id": self._next_set_id(),
                "name": "Set",
                "linked_element_definitions": [],
            }
            self._current_profile.setdefault("linked_sets", []).append(new_set)
            target_set = new_set
        elif len(sets) == 1:
            target_set = sets[0]
        else:
            target_set = wpf_dialogs.pick_from_list(
                sets,
                title="Paste fixture",
                prompt="Pick the destination linked_set:",
                display_func=lambda s: "{}  ({})".format(
                    s.get("name") or "set", s.get("id") or "?"
                ),
            )
            if target_set is None:
                return

        new_led = copy.deepcopy(self._fixture_clipboard)

        # Generate a unique LED id within the target set.
        target_set_id = target_set.get("id") or "SET-???"
        existing = {
            (l.get("id") or "")
            for l in target_set.get("linked_element_definitions") or []
        }
        idx = 1
        while True:
            led_id = "{}-LED-{:03d}".format(target_set_id, idx)
            if led_id not in existing:
                break
            idx += 1
        new_led["id"] = led_id

        # Reset offsets to zero (paste at origin, no relative pose).
        new_led["offsets"] = [{
            "x_inches": 0.0,
            "y_inches": 0.0,
            "z_inches": 0.0,
            "rotation_deg": 0.0,
        }]

        # Strip parent-relationship directives + stale parent / element
        # references from the LED's parameters.
        new_led["parameters"] = self._strip_parent_directives(
            new_led.get("parameters") or {}
        )

        # Renumber annotations under the pasted LED.
        new_anns = []
        for i, ann in enumerate(new_led.get("annotations") or []):
            if not isinstance(ann, dict):
                continue
            ann_copy = copy.deepcopy(ann)
            ann_copy["id"] = "{}-ANN-{:03d}".format(led_id, i + 1)
            new_anns.append(ann_copy)
        new_led["annotations"] = new_anns

        target_set.setdefault("linked_element_definitions", []).append(new_led)

        self.dirty = True
        self._populate_led_tree(self._current_profile)
        self._set_status(
            "Pasted fixture into set {!r} as {!r}".format(
                target_set.get("id") or "?", led_id
            )
        )

    @staticmethod
    def _strip_parent_directives(params):
        """Drop BYPARENT / BYSIBLING directive entries and stale
        parent / element references from a parameter dict so the pasted
        LED has no lingering link to its source profile's parent."""
        _STALE_KEYS = (
            "Parent ElementId", "Parent Element ID", "Parent ID",
            "Parent Rotation (deg)", "Parent_location",
            "Element_Linker Parameter",
            "ElementId", "Element Id", "Element ID",
            "Linked Element Definition ID",
            "Set Definition ID",
            "Location XYZ (ft)",
            "FacingOrientation",
            "LevelId",
        )
        out = {}
        for k, v in params.items():
            if isinstance(v, dict) and (
                "parent_parameter" in v or "sibling_parameter" in v
            ):
                continue  # parent / sibling directive
            if k in _STALE_KEYS:
                continue  # stale Element_Linker-derived value
            out[k] = v
        return out

    def _next_set_id(self):
        """Lowest unused ``SET-NNN`` across the whole document."""
        seen = set()
        for p in self.profile_data.get("equipment_definitions") or []:
            if not isinstance(p, dict):
                continue
            for s in p.get("linked_sets") or []:
                if isinstance(s, dict) and s.get("id"):
                    seen.add(s["id"])
        n = 1
        while True:
            candidate = "SET-{:03d}".format(n)
            if candidate not in seen:
                return candidate
            n += 1

    def _commit_grid_edits(self, grid=None):
        """WPF DataGrid edits aren't visible in the bound items until both
        the cell and the row are committed. CommitEdit() commits one
        level; calling it twice covers the cell + row.

        Defaults to the LED/ANN ``parameter_grid`` if ``grid`` is None.
        """
        target = grid if grid is not None else self.parameter_grid
        try:
            target.CommitEdit()
            target.CommitEdit()
        except Exception:
            pass

    def _populate_profile_metadata(self, profile):
        meta_raw = profile.get("equipment_properties")
        if not isinstance(meta_raw, dict):
            meta_raw = {}
        rows = ObservableCollection[_NetObject]()
        for k, v in meta_raw.items():
            rows.Add(_ParamRow(str(k), "" if v is None else str(v)))
        self.profile_meta_grid.ItemsSource = rows

    def _populate_aliases(self, profile):
        """Refresh the editable aliases ListBox + the legacy info label."""
        if not hasattr(self, "alias_list") or self.alias_list is None:
            return
        self.alias_list.Items.Clear()
        if profile is None:
            self.alias_info_label.Text = ""
            return
        for alias in merge_workflow.aliases(profile):
            self.alias_list.Items.Add(alias)

        # Surface a hint if this profile has legacy ced_truth_source markers
        # — they're inert under the new model but worth flagging until
        # the user runs the migration via Merge Profiles.
        info_bits = []
        if truth_groups.is_group_member(profile):
            info_bits.append(
                "Note: this profile still carries a legacy "
                "ced_truth_source_id pointing at {!r}. Run Merge Profiles "
                "to migrate these.".format(
                    truth_groups.truth_source_name(profile) or "?"
                )
            )
        all_profiles = self.profile_data.get("equipment_definitions") or []
        legacy_members = truth_groups.find_group_members(all_profiles, profile.get("id"))
        if legacy_members:
            info_bits.append(
                "Note: {} legacy member(s) point at this profile via "
                "ced_truth_source_id; run Merge Profiles to migrate.".format(
                    len(legacy_members)
                )
            )
        self.alias_info_label.Text = "  ".join(info_bits)

    def _read_profile_meta_grid(self):
        out = {}
        for row in (self.profile_meta_grid.ItemsSource or []):
            key = (getattr(row, "Name", "") or "").strip()
            if not key:
                continue
            out[key] = getattr(row, "Value", "")
        return out

    def _save_profile_metadata_to_data(self):
        if self._current_profile is None:
            return False
        self._commit_grid_edits(self.profile_meta_grid)
        self._current_profile["equipment_properties"] = self._read_profile_meta_grid()
        return True

    def _on_save_profile_metadata(self, sender, e):
        if self._current_profile is None:
            self._set_status("Pick a profile first")
            return
        self._save_profile_metadata_to_data()
        self.dirty = True
        self._set_status("Saved profile metadata (unsaved to project)")

    def _rebuild_profile_meta_rows(self, mutator):
        """Apply ``mutator`` to a snapshot of current rows, then reassign
        the grid's ItemsSource to a fresh ObservableCollection.

        Reassigning is the only refresh path that's reliable in
        CPython 3 + pythonnet — ObservableCollection.CollectionChanged
        events don't always propagate to the DataGrid when items are
        plain Python objects.
        """
        self._commit_grid_edits(self.profile_meta_grid)
        snapshot = []
        for row in (self.profile_meta_grid.ItemsSource or []):
            snapshot.append(row)
        mutator(snapshot)
        new_collection = ObservableCollection[_NetObject]()
        for r in snapshot:
            new_collection.Add(r)
        self.profile_meta_grid.ItemsSource = new_collection
        return new_collection

    def _on_add_profile_metadata(self, sender, e):
        if self._current_profile is None:
            self._set_status("Pick a profile first")
            return
        new_row = _ParamRow("", "")

        def add(rows):
            rows.append(new_row)

        self._rebuild_profile_meta_rows(add)
        # Scroll the new row into view + select it so the user can start typing.
        try:
            self.profile_meta_grid.ScrollIntoView(new_row)
            self.profile_meta_grid.SelectedItem = new_row
        except Exception:
            pass
        self._set_status("Added metadata row — type a Name and Value, then Save profile metadata")

    def _on_remove_profile_metadata(self, sender, e):
        selected = self.profile_meta_grid.SelectedItem
        if selected is None:
            self._set_status("Pick a metadata row to remove first")
            return

        removed = [False]

        def drop(rows):
            try:
                rows.remove(selected)
                removed[0] = True
            except ValueError:
                pass

        self._rebuild_profile_meta_rows(drop)
        if removed[0]:
            self._set_status("Removed metadata row — Save profile metadata to commit")
        else:
            self._set_status("Could not locate the selected row")

    # ---- alias editing ---------------------------------------------

    def _read_aliases_from_listbox(self):
        return [str(item) for item in self.alias_list.Items if item]

    def _save_aliases_to_data(self):
        """Persist the listbox aliases back onto the current profile."""
        if self._current_profile is None:
            return False
        new_aliases = self._read_aliases_from_listbox()
        # Replace the list in-place (preserves dict identity).
        self._current_profile[merge_workflow.MERGED_ALIASES_KEY] = new_aliases
        return True

    def _on_add_alias(self, sender, e):
        if self._current_profile is None:
            self._set_status("Pick a profile first")
            return
        text = wpf_dialogs.prompt_for_string(
            "Add alias to {!r}:".format(self._current_profile.get("name") or "?"),
            title="Add merged alias",
            default="",
        )
        if not text:
            return
        clean = text.strip()
        if not clean:
            return
        existing = {
            str(item).strip().lower() for item in self.alias_list.Items
        }
        if clean.lower() in existing:
            self._set_status("Alias already present (deduped)")
            return
        self.alias_list.Items.Add(clean)
        self._set_status("Alias added — click 'Save aliases' to commit")

    def _on_remove_alias(self, sender, e):
        selected = self.alias_list.SelectedItem
        if selected is None:
            self._set_status("Pick an alias to remove first")
            return
        self.alias_list.Items.Remove(selected)
        self._set_status("Alias removed — click 'Save aliases' to commit")

    def _on_save_aliases(self, sender, e):
        if self._current_profile is None:
            self._set_status("Pick a profile first")
            return
        if not self._save_aliases_to_data():
            self._set_status("Save failed — no current profile")
            return
        self.dirty = True
        self._set_status("Saved aliases (in memory) — Save changes to commit")

    def _read_param_grid(self):
        """Read parameter rows from the grid into a fresh ``{name: value}`` dict."""
        out = {}
        for row in (self.parameter_grid.ItemsSource or []):
            key = (getattr(row, "Name", "") or "").strip()
            if not key:
                continue
            out[key] = getattr(row, "Value", "")
        return out

    def _read_offset_inputs(self):
        return {
            "x_inches": _coerce_float(self.offset_x.Text),
            "y_inches": _coerce_float(self.offset_y.Text),
            "z_inches": _coerce_float(self.offset_z.Text),
            "rotation_deg": _coerce_float(self.offset_rot.Text),
        }

    def _save_selected_to_data(self):
        """Persist the parameter grid + offset inputs to whichever node is
        currently loaded (annotation > LED). Returns a status string."""
        self._commit_grid_edits()
        params = self._read_param_grid()
        offset = self._read_offset_inputs()
        if self._current_ann is not None:
            self._current_ann["parameters"] = params
            self._current_ann["offsets"] = offset
            return "Saved annotation edits (unsaved to project)"
        if self._current_led is not None:
            self._current_led["parameters"] = params
            offsets = self._current_led.setdefault("offsets", [{}])
            if not offsets:
                offsets.append({})
            offsets[0].update(offset)
            return "Saved LED edits (unsaved to project)"
        return "Nothing selected to save"

    def _on_save_selected(self, sender, e):
        if self._current_led is None and self._current_ann is None:
            self._set_status("Pick a LED or annotation first")
            return
        msg = self._save_selected_to_data()
        self.dirty = True
        # Re-render tree so updated labels (if any) reflect new state.
        if self._current_profile is not None:
            self._populate_led_tree(self._current_profile)
        self._set_status(msg)

    def _on_save(self, sender, e):
        profile = self._current_profile
        if profile is None:
            self._set_status("No profile selected")
            return
        profile["name"] = self.name_box.Text or profile.get("name") or ""
        pf = profile.setdefault("parent_filter", {})
        pf["category"] = self.cat_box.Text or ""
        pf["family_name_pattern"] = self.fam_box.Text or ""
        pf["type_name_pattern"] = self.type_box.Text or ""
        profile["allow_parentless"] = bool(self.allow_parentless.IsChecked)
        profile["allow_unmatched_parents"] = bool(self.allow_unmatched.IsChecked)
        profile["prompt_on_parent_mismatch"] = bool(self.prompt_mismatch.IsChecked)

        # Also persist the profile metadata grid + aliases + the currently
        # selected LED / ANN before the global "Save changes" returns
        # control to the caller (which writes the whole document back to
        # Extensible Storage).
        self._save_profile_metadata_to_data()
        self._save_aliases_to_data()
        self._save_selected_to_data()

        self.dirty = True
        self._reload_profile_list(self.filter_box.Text)
        self._set_status("Saved (in memory)")

    def _on_close(self, sender, e):
        self.window.Close()

    def _set_status(self, msg):
        self.status.Text = msg

    # ---- modal entry -------------------------------------------------

    def show(self):
        self.window.ShowDialog()
        return self.dirty


def show_modal(profile_data, doc=None):
    return ManageProfilesController(profile_data, doc=doc).show()
