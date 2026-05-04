# -*- coding: utf-8 -*-
"""
Modal editor for ``space_profiles[*]`` in the active YAML payload.

The editor exposes the structural fields the placement engine consumes:
profile name, owning bucket, and a flat per-LED table that captures
(linked-set name, label, category) plus the placement-rule fields
(anchor kind, inset, door offset).

Per-LED ``parameters`` / ``offsets`` / ``annotations`` are intentionally
NOT editable in this first iteration — the YAML round-trip through
Import/Export Profiles is the escape hatch. We can add deeper editors
in a follow-up batch when the placement engine surfaces the need.

The controller mutates ``profile_data["space_profiles"]`` in place; the
calling pushbutton script saves via ``active_yaml.save_active_data``
inside its own Revit transaction (mirroring Manage Profiles).
"""

import copy
import os
import uuid

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    Button as _WpfButton,
    SelectionChangedEventHandler,
)

import wpf as _wpf  # noqa: E402
import space_profile_model as _profile_model  # noqa: E402
import space_led_details_window as _led_details  # noqa: E402


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "ManageSpaceProfilesWindow.xaml",
)


# ---------------------------------------------------------------------
# Display-only wrappers
# ---------------------------------------------------------------------

class _ProfileItem(object):
    """ListBox display row."""

    def __init__(self, profile_dict):
        self.data = profile_dict

    @property
    def DisplayName(self):
        bucket = self.data.get("bucket_id") or "(no bucket)"
        return "{}  ({})  -> {}".format(
            self.data.get("name") or "(unnamed)",
            self.data.get("id") or "??",
            bucket,
        )


class _BucketComboItem(object):
    """ItemsSource entry for the bucket ComboBox."""

    def __init__(self, bucket_dict):
        self.Id = bucket_dict.get("id") or ""
        name = bucket_dict.get("name") or "(unnamed)"
        self.Display = "{}  ({})".format(name, self.Id or "??")


_KIND_OPTIONS = list(_profile_model.PLACEMENT_KINDS)


class _LedRow(object):
    """One row in the per-profile flat LED grid.

    Backed by two underlying YAML dicts: the linked-set dict and the
    LED dict. Edits via WPF land on this row first; the controller
    re-syncs to YAML in ``flush_to_yaml``.
    """

    def __init__(self, led_dict, set_dict):
        self._led = led_dict
        self._set = set_dict
        self.KindOptions = _KIND_OPTIONS

    # Set name (binds to set_dict, shared across rows from the same set)

    @property
    def SetName(self):
        return self._set.get("name") or ""

    @SetName.setter
    def SetName(self, value):
        # Empty set name OK; controller normalises on save.
        self._set["name"] = (value or "").strip()

    # LED display fields

    @property
    def LedId(self):
        return self._led.get("id") or ""

    @property
    def Label(self):
        return self._led.get("label") or ""

    @Label.setter
    def Label(self, value):
        self._led["label"] = (value or "").strip()

    @property
    def Category(self):
        return self._led.get("category") or ""

    @Category.setter
    def Category(self, value):
        self._led["category"] = (value or "").strip()

    # Placement-rule fields

    def _rule(self):
        rule = self._led.setdefault("placement_rule", {})
        if not isinstance(rule, dict):
            rule = {}
            self._led["placement_rule"] = rule
        return rule

    @property
    def Kind(self):
        return self._rule().get("kind") or _profile_model.KIND_CENTER

    @Kind.setter
    def Kind(self, value):
        if value in _profile_model.PLACEMENT_KINDS:
            self._rule()["kind"] = value

    @property
    def InsetText(self):
        v = self._rule().get("inset_inches")
        if v is None:
            return ""
        try:
            return _format_float(float(v))
        except (ValueError, TypeError):
            return ""

    @InsetText.setter
    def InsetText(self, value):
        v = _coerce_float(value, default=None)
        if v is None:
            self._rule().pop("inset_inches", None)
        else:
            self._rule()["inset_inches"] = v

    def _door_offset_dict(self):
        rule = self._rule()
        d = rule.setdefault("door_offset_inches", {})
        if not isinstance(d, dict):
            d = {}
            rule["door_offset_inches"] = d
        return d

    @property
    def DoorOffsetXText(self):
        v = self._door_offset_dict().get("x")
        return _format_float(float(v)) if v is not None else ""

    @DoorOffsetXText.setter
    def DoorOffsetXText(self, value):
        v = _coerce_float(value, default=None)
        if v is None:
            self._door_offset_dict().pop("x", None)
        else:
            self._door_offset_dict()["x"] = v

    @property
    def DoorOffsetYText(self):
        v = self._door_offset_dict().get("y")
        return _format_float(float(v)) if v is not None else ""

    @DoorOffsetYText.setter
    def DoorOffsetYText(self, value):
        v = _coerce_float(value, default=None)
        if v is None:
            self._door_offset_dict().pop("y", None)
        else:
            self._door_offset_dict()["y"] = v


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _format_float(v):
    if v == int(v):
        return str(int(v))
    return "{:.4f}".format(v).rstrip("0").rstrip(".")


def _coerce_float(text, default=None):
    if text is None:
        return default
    s = str(text).strip()
    if not s:
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _new_id(prefix):
    """Short stable id; not user-visible-critical so a uuid suffix is fine."""
    return "{}-{}".format(prefix, uuid.uuid4().hex[:8].upper())


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class ManageSpaceProfilesController(object):
    """Modal editor. Caller passes in the YAML dict; mutations land in it."""

    def __init__(self, profile_data, doc=None):
        self.profile_data = profile_data
        self.doc = doc
        self.dirty = False

        self.window = _wpf.load_xaml(_XAML_PATH)
        self._profile_items = ObservableCollection[_NetObject]()
        self._led_rows = ObservableCollection[_NetObject]()
        self._buckets = []  # [_BucketComboItem, ...]
        self._loading = False  # guards re-entrancy in field setters

        self._lookup_controls()
        self._wire_events()
        self._populate_bucket_combo()
        self._reload_profile_list()
        self._set_status("Ready.")
        self._select_profile(None)

    # ----- bootstrapping -------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.profile_list = f("ProfileList")
        self.profile_summary_label = f("ProfileSummaryLabel")
        self.detail_header = f("DetailHeader")
        self.name_box = f("NameBox")
        self.bucket_combo = f("BucketCombo")
        self.led_summary_label = f("LedSummaryLabel")
        self.led_grid = f("LedGrid")
        self.new_profile_btn = f("NewProfileButton")
        self.duplicate_profile_btn = f("DuplicateProfileButton")
        self.delete_profile_btn = f("DeleteProfileButton")
        self.add_led_btn = f("AddLedButton")
        self.delete_led_btn = f("DeleteLedButton")
        self.save_btn = f("SaveButton")
        self.close_btn = f("CloseButton")
        self.status_label = f("StatusLabel")

        self.profile_list.ItemsSource = self._profile_items
        self.led_grid.ItemsSource = self._led_rows

    def _wire_events(self):
        self._h_new = RoutedEventHandler(
            lambda s, e: self._safe(self._on_new_profile, "new-profile")
        )
        self._h_dup = RoutedEventHandler(
            lambda s, e: self._safe(self._on_duplicate_profile, "duplicate-profile")
        )
        self._h_del = RoutedEventHandler(
            lambda s, e: self._safe(self._on_delete_profile, "delete-profile")
        )
        self._h_add_led = RoutedEventHandler(
            lambda s, e: self._safe(self._on_add_led, "add-led")
        )
        self._h_del_led = RoutedEventHandler(
            lambda s, e: self._safe(self._on_delete_led, "delete-led")
        )
        self._h_save = RoutedEventHandler(
            lambda s, e: self._safe(self._on_save, "save")
        )
        self._h_close = RoutedEventHandler(
            lambda s, e: self.window.Close()
        )
        self._h_select = SelectionChangedEventHandler(
            lambda s, e: self._safe(self._on_profile_selection_changed, "profile-select")
        )
        self._h_name_lost = RoutedEventHandler(
            lambda s, e: self._safe(self._on_name_lost_focus, "name-edit")
        )
        self._h_bucket_change = SelectionChangedEventHandler(
            lambda s, e: self._safe(self._on_bucket_changed, "bucket-change")
        )

        self.new_profile_btn.Click += self._h_new
        self.duplicate_profile_btn.Click += self._h_dup
        self.delete_profile_btn.Click += self._h_del
        self.add_led_btn.Click += self._h_add_led
        self.delete_led_btn.Click += self._h_del_led
        self.save_btn.Click += self._h_save
        self.close_btn.Click += self._h_close
        self.profile_list.SelectionChanged += self._h_select
        self.name_box.LostFocus += self._h_name_lost
        self.bucket_combo.SelectionChanged += self._h_bucket_change

        # Bubbled Click handler for the per-row "Details..." button on
        # the LED grid. Tag filtering ensures non-row buttons (Save,
        # Close, etc.) are ignored cheaply.
        self._h_details_click = RoutedEventHandler(
            lambda s, e: self._safe_with(s, e, self._on_details_click, "led-details")
        )
        self.window.AddHandler(_WpfButton.ClickEvent, self._h_details_click)

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _safe_with(self, sender, e, fn, label):
        try:
            fn(sender, e)
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    # ----- list helpers --------------------------------------------

    def _profiles(self):
        raw = self.profile_data.setdefault("space_profiles", [])
        if not isinstance(raw, list):
            raw = []
            self.profile_data["space_profiles"] = raw
        return raw

    def _reload_profile_list(self, select_id=None):
        self._profile_items.Clear()
        for p in self._profiles():
            if isinstance(p, dict):
                self._profile_items.Add(_ProfileItem(p))
        self.profile_summary_label.Text = "{} profile(s)".format(
            self._profile_items.Count
        )
        if select_id:
            self._select_by_id(select_id)

    def _select_by_id(self, profile_id):
        for item in self._profile_items:
            if item.data.get("id") == profile_id:
                self.profile_list.SelectedItem = item
                return

    def _selected_profile_dict(self):
        item = self.profile_list.SelectedItem
        if item is None:
            return None
        return getattr(item, "data", None)

    def _populate_bucket_combo(self):
        self._buckets = []
        for b in (self.profile_data.get("space_buckets") or ()):
            if isinstance(b, dict):
                self._buckets.append(_BucketComboItem(b))
        self.bucket_combo.ItemsSource = self._buckets

    # ----- profile selection ---------------------------------------

    def _on_profile_selection_changed(self):
        self._select_profile(self._selected_profile_dict())

    def _select_profile(self, profile_dict):
        self._loading = True
        try:
            if profile_dict is None:
                self.detail_header.Text = "Select a profile..."
                self.name_box.Text = ""
                self.bucket_combo.SelectedValue = None
                self._led_rows.Clear()
                self.led_summary_label.Text = ""
                self._set_enabled(False)
                return
            self._set_enabled(True)
            pid = profile_dict.get("id") or "?"
            name = profile_dict.get("name") or "(unnamed)"
            self.detail_header.Text = "Profile: {}  [{}]".format(name, pid)
            self.name_box.Text = name
            self.bucket_combo.SelectedValue = profile_dict.get("bucket_id") or ""
            self._reload_led_rows(profile_dict)
        finally:
            self._loading = False

    def _set_enabled(self, enabled):
        self.name_box.IsEnabled = enabled
        self.bucket_combo.IsEnabled = enabled
        self.led_grid.IsEnabled = enabled
        self.add_led_btn.IsEnabled = enabled
        self.delete_led_btn.IsEnabled = enabled
        self.duplicate_profile_btn.IsEnabled = enabled
        self.delete_profile_btn.IsEnabled = enabled

    # ----- name + bucket edits -------------------------------------

    def _on_name_lost_focus(self):
        if self._loading:
            return
        prof = self._selected_profile_dict()
        if prof is None:
            return
        new_name = (self.name_box.Text or "").strip()
        if new_name and new_name != prof.get("name"):
            prof["name"] = new_name
            self.dirty = True
            self._refresh_profile_row(prof)
            self.detail_header.Text = "Profile: {}  [{}]".format(
                new_name, prof.get("id") or "?"
            )
            self._set_status("Renamed profile.")

    def _on_bucket_changed(self):
        if self._loading:
            return
        prof = self._selected_profile_dict()
        if prof is None:
            return
        new_bucket = self.bucket_combo.SelectedValue
        if new_bucket is None:
            new_bucket = ""
        if new_bucket != (prof.get("bucket_id") or ""):
            prof["bucket_id"] = new_bucket or None
            self.dirty = True
            self._refresh_profile_row(prof)
            self._set_status("Updated bucket.")

    def _refresh_profile_row(self, prof_dict):
        # Force the ListBox to redraw the affected row (DisplayName changed).
        for i, item in enumerate(self._profile_items):
            if item.data is prof_dict:
                # Trigger a refresh by reseating the item.
                self._profile_items[i] = _ProfileItem(prof_dict)
                self.profile_list.SelectedIndex = i
                return

    # ----- profile add / duplicate / delete ------------------------

    def _on_new_profile(self):
        new = {
            "id": _new_id("SP"),
            "name": "New Space Profile",
            "bucket_id": None,
            "linked_sets": [],
        }
        self._profiles().append(new)
        self.dirty = True
        self._reload_profile_list(select_id=new["id"])
        self._set_status("New profile created.")

    def _on_duplicate_profile(self):
        prof = self._selected_profile_dict()
        if prof is None:
            return
        clone = copy.deepcopy(prof)
        clone["id"] = _new_id("SP")
        clone["name"] = "{} (copy)".format(prof.get("name") or "Profile")
        # Reassign LED IDs so duplicates don't share LED ids with the
        # original — otherwise saved Element_Linker payloads on placed
        # elements would point at the same id in two profiles.
        for s in clone.get("linked_sets") or ():
            if not isinstance(s, dict):
                continue
            s["id"] = _new_id("SET")
            for led in s.get("linked_element_definitions") or ():
                if isinstance(led, dict):
                    led["id"] = _new_id("LED")
        self._profiles().append(clone)
        self.dirty = True
        self._reload_profile_list(select_id=clone["id"])
        self._set_status("Profile duplicated.")

    def _on_delete_profile(self):
        prof = self._selected_profile_dict()
        if prof is None:
            return
        self._profiles().remove(prof)
        self.dirty = True
        self._reload_profile_list()
        self._select_profile(None)
        self._set_status("Profile deleted.")

    # ----- LED rows ------------------------------------------------

    def _reload_led_rows(self, profile_dict):
        self._led_rows.Clear()
        sets = profile_dict.setdefault("linked_sets", [])
        if not isinstance(sets, list):
            sets = []
            profile_dict["linked_sets"] = sets
        for s in sets:
            if not isinstance(s, dict):
                continue
            for led in (s.setdefault("linked_element_definitions", []) or ()):
                if isinstance(led, dict):
                    led.setdefault("id", _new_id("LED"))
                    self._led_rows.Add(_LedRow(led, s))
        self._refresh_led_summary()

    def _refresh_led_summary(self):
        n_rows = self._led_rows.Count
        prof = self._selected_profile_dict() or {}
        n_sets = len(prof.get("linked_sets") or [])
        self.led_summary_label.Text = "{} LED(s) across {} set(s)".format(
            n_rows, n_sets
        )

    def _on_add_led(self):
        prof = self._selected_profile_dict()
        if prof is None:
            return
        # If a row is selected, drop the new LED into the same set so
        # the user can build the same set quickly.
        target_set = None
        sel = self.led_grid.SelectedItem
        if isinstance(sel, _LedRow):
            target_set = sel._set
        if target_set is None:
            sets = prof.setdefault("linked_sets", [])
            if not isinstance(sets, list):
                sets = []
                prof["linked_sets"] = sets
            if sets and isinstance(sets[-1], dict):
                target_set = sets[-1]
            else:
                target_set = {
                    "id": _new_id("SET"),
                    "name": "Default",
                    "linked_element_definitions": [],
                }
                sets.append(target_set)

        new_led = {
            "id": _new_id("LED"),
            "label": "",
            "category": "",
            "placement_rule": {
                "kind": _profile_model.KIND_CENTER,
                "inset_inches": 0.0,
            },
        }
        leds = target_set.setdefault("linked_element_definitions", [])
        if not isinstance(leds, list):
            leds = []
            target_set["linked_element_definitions"] = leds
        leds.append(new_led)
        self._led_rows.Add(_LedRow(new_led, target_set))
        self._refresh_led_summary()
        self.dirty = True
        self._set_status("Added LED to set '{}'.".format(target_set.get("name") or ""))

    def _on_details_click(self, sender, e):
        # Bubbled Button.Click — Tag is bound to the _LedRow if the
        # clicked button is the per-row "Details..." button.
        source = getattr(e, "Source", None) or getattr(e, "OriginalSource", None)
        tag = getattr(source, "Tag", None) if source is not None else None
        if not isinstance(tag, _LedRow):
            return
        row = tag
        led_label = row.Label or row.LedId or "(unnamed LED)"
        header = "Edit LED: {} [{}]".format(led_label, row.LedId or "?")
        ok = _led_details.show_modal(
            led_dict=row._led, header=header, owner=self.window,
        )
        if ok:
            self.dirty = True
            # The set name and label cells in the parent grid may have
            # become stale (label is the only one we let the sub-dialog
            # touch indirectly via parameters); a refresh keeps the
            # display honest.
            self.led_grid.Items.Refresh()
            self._set_status(
                "Updated details for {}.".format(led_label)
            )

    def _on_delete_led(self):
        sel = self.led_grid.SelectedItem
        if not isinstance(sel, _LedRow):
            self._set_status("Select an LED row to delete.")
            return
        # Remove the LED dict from its set.
        leds = sel._set.get("linked_element_definitions") or []
        try:
            leds.remove(sel._led)
        except ValueError:
            pass
        self._led_rows.Remove(sel)

        # If the set is now empty, drop it from the profile.
        prof = self._selected_profile_dict() or {}
        if not leds and isinstance(prof.get("linked_sets"), list):
            try:
                prof["linked_sets"].remove(sel._set)
            except ValueError:
                pass

        self._refresh_led_summary()
        self.dirty = True
        self._set_status("LED removed.")

    # ----- save ----------------------------------------------------

    def _on_save(self):
        # The DataGrid CellEditEnding fires LostFocus already on Tab/Enter,
        # but a click on Save while focus is in a cell can leave one edit
        # in flight. Force the commit chain.
        try:
            self.led_grid.CommitEdit()
            self.led_grid.CommitEdit()
        except Exception:
            pass
        # Tidy: drop empty linked-sets that lost their LEDs.
        for prof in self._profiles():
            if not isinstance(prof, dict):
                continue
            sets = prof.get("linked_sets") or []
            prof["linked_sets"] = [
                s for s in sets
                if isinstance(s, dict)
                and (s.get("linked_element_definitions") or [])
            ]
        self.dirty = True
        self._set_status("Edits flushed. Click Close to save & dismiss.")

    # ----- entry point --------------------------------------------

    def show(self):
        self.window.ShowDialog()


# ---------------------------------------------------------------------
# Convenience entry point
# ---------------------------------------------------------------------

def show_modal(profile_data, doc=None):
    controller = ManageSpaceProfilesController(profile_data=profile_data, doc=doc)
    controller.show()
    return controller
