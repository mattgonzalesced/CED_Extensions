# -*- coding: utf-8 -*-
"""
General Proximity Check
Compare two selected category/family groups and report nearby instances.
"""

__title__ = "General Proximity\nCheck"
__doc__ = (
    "Find instance pairs between two selected category family/type groups "
    "that are within a user-entered distance."
)

import math

import clr
from pyrevit import DB, forms, revit, script

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Size
from System.Windows.Forms import (
    AnchorStyles,
    Button,
    CheckBox,
    CheckedListBox,
    CheckState,
    ColumnStyle,
    ComboBox,
    ComboBoxStyle,
    DialogResult,
    DockStyle,
    FlowDirection,
    FlowLayoutPanel,
    Form,
    FormBorderStyle,
    FormStartPosition,
    Label,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
    Padding,
    RowStyle,
    SizeType,
    TableLayoutPanel,
    TextBox,
)

TITLE = "General Proximity Check"
DISTANCE_MODE_3D = "3d"
DISTANCE_MODE_XY = "xy"
DISTANCE_MODE_OPTIONS = (
    ("3D", DISTANCE_MODE_3D),
    ("Plan (XY only)", DISTANCE_MODE_XY),
)


def _bic(name):
    return getattr(DB.BuiltInCategory, name, None)


def _bip(name):
    return getattr(DB.BuiltInParameter, name, None)


SYSTEM_CATEGORY_BICS = tuple(
    bic
    for bic in (
        _bic("OST_PipeCurves"),
        _bic("OST_FlexPipeCurves"),
        _bic("OST_DuctCurves"),
        _bic("OST_FlexDuctCurves"),
        _bic("OST_Conduit"),
        _bic("OST_CableTray"),
        _bic("OST_FabricationPipework"),
        _bic("OST_FabricationDuctwork"),
        _bic("OST_FabricationContainment"),
    )
    if bic is not None
)

PIPE_SIZE_BIPS = tuple(
    bip
    for bip in (
        _bip("RBS_PIPE_OUTER_DIAMETER"),
        _bip("RBS_PIPE_DIAMETER_PARAM"),
    )
    if bip is not None
)
DUCT_DIAMETER_BIPS = tuple(
    bip
    for bip in (
        _bip("RBS_CURVE_DIAMETER_PARAM"),
    )
    if bip is not None
)
DUCT_RECT_BIPS = tuple(
    bip
    for bip in (
        _bip("RBS_CURVE_WIDTH_PARAM"),
        _bip("RBS_CURVE_HEIGHT_PARAM"),
    )
    if bip is not None
)
CONDUIT_SIZE_BIPS = tuple(
    bip
    for bip in (
        _bip("RBS_CONDUIT_DIAMETER_PARAM"),
    )
    if bip is not None
)
CABLETRAY_SIZE_BIPS = tuple(
    bip
    for bip in (
        _bip("RBS_CABLETRAY_WIDTH_PARAM"),
        _bip("RBS_CABLETRAY_HEIGHT_PARAM"),
    )
    if bip is not None
)


class _CategoryChoice(object):
    def __init__(self, data):
        self.data = data

    def __str__(self):
        option_count = len(self.data.get("selectors") or [])
        return "{} ({} options)".format(self.data.get("name") or "<Unknown>", option_count)

    def ToString(self):
        return self.__str__()


class _GeneralProximityDialog(Form):
    def __init__(self, categories):
        Form.__init__(self)
        self._categories = categories or []
        self._is_syncing_checks = False
        self.result = None

        self._combo_a = None
        self._combo_b = None
        self._families_a = None
        self._families_b = None
        self._all_a = None
        self._all_b = None
        self._threshold_box = None
        self._mode_combo = None

        self._build_ui()
        self._load_categories()

    def _build_ui(self):
        self.Text = TITLE
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(1080, 620)
        self.Size = Size(1240, 760)

        root = TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.Padding = Padding(12)
        root.ColumnCount = 1
        root.RowCount = 2
        root.RowStyles.Add(RowStyle(SizeType.Percent, 100.0))
        root.RowStyles.Add(RowStyle(SizeType.AutoSize))
        self.Controls.Add(root)

        content = TableLayoutPanel()
        content.Dock = DockStyle.Fill
        content.ColumnCount = 3
        content.RowCount = 1
        content.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 45.0))
        content.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 10.0))
        content.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 45.0))
        root.Controls.Add(content, 0, 0)

        left_panel = self._build_side_panel("Category 1", side_key="a")
        center_panel = self._build_center_panel()
        right_panel = self._build_side_panel("Category 2", side_key="b")

        content.Controls.Add(left_panel, 0, 0)
        content.Controls.Add(center_panel, 1, 0)
        content.Controls.Add(right_panel, 2, 0)

        button_row = FlowLayoutPanel()
        button_row.Dock = DockStyle.Fill
        button_row.AutoSize = True
        button_row.FlowDirection = FlowDirection.RightToLeft
        button_row.WrapContents = False
        root.Controls.Add(button_row, 0, 1)

        run_btn = Button()
        run_btn.Text = "Run Check"
        run_btn.AutoSize = True
        run_btn.Click += self._on_run
        button_row.Controls.Add(run_btn)

        cancel_btn = Button()
        cancel_btn.Text = "Cancel"
        cancel_btn.AutoSize = True
        cancel_btn.Click += self._on_cancel
        button_row.Controls.Add(cancel_btn)

        self.AcceptButton = run_btn
        self.CancelButton = cancel_btn

    def _build_side_panel(self, title_text, side_key):
        panel = TableLayoutPanel()
        panel.Dock = DockStyle.Fill
        panel.Padding = Padding(8, 0, 8, 0)
        panel.ColumnCount = 1
        panel.RowCount = 4
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.Percent, 100.0))

        title = Label()
        title.Text = title_text
        title.AutoSize = True
        title.Margin = Padding(0, 0, 0, 6)
        panel.Controls.Add(title, 0, 0)

        combo = ComboBox()
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        combo.Dock = DockStyle.Top
        combo.Margin = Padding(0, 0, 0, 8)
        panel.Controls.Add(combo, 0, 1)

        all_box = CheckBox()
        all_box.Text = "All Families / Types"
        all_box.AutoSize = True
        all_box.Margin = Padding(0, 0, 0, 6)
        panel.Controls.Add(all_box, 0, 2)

        family_list = CheckedListBox()
        family_list.Dock = DockStyle.Fill
        family_list.CheckOnClick = True
        panel.Controls.Add(family_list, 0, 3)

        if side_key == "a":
            self._combo_a = combo
            self._families_a = family_list
            self._all_a = all_box
            combo.SelectedIndexChanged += self._on_category_a_changed
            all_box.CheckedChanged += self._on_all_a_changed
            family_list.ItemCheck += self._on_family_item_check
        else:
            self._combo_b = combo
            self._families_b = family_list
            self._all_b = all_box
            combo.SelectedIndexChanged += self._on_category_b_changed
            all_box.CheckedChanged += self._on_all_b_changed
            family_list.ItemCheck += self._on_family_item_check

        return panel

    def _build_center_panel(self):
        panel = TableLayoutPanel()
        panel.Dock = DockStyle.Fill
        panel.ColumnCount = 1
        panel.RowCount = 9
        panel.RowStyles.Add(RowStyle(SizeType.Percent, 25.0))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(RowStyle(SizeType.Percent, 75.0))

        mode_label = Label()
        mode_label.Text = "Distance Mode"
        mode_label.AutoSize = True
        mode_label.Anchor = getattr(AnchorStyles, "None")
        mode_label.Margin = Padding(0, 0, 0, 6)
        panel.Controls.Add(mode_label, 0, 1)

        mode_combo = ComboBox()
        mode_combo.DropDownStyle = ComboBoxStyle.DropDownList
        mode_combo.Width = 110
        mode_combo.Anchor = getattr(AnchorStyles, "None")
        for item in DISTANCE_MODE_OPTIONS:
            mode_combo.Items.Add(item[0])
        mode_combo.SelectedIndex = 0
        panel.Controls.Add(mode_combo, 0, 2)
        self._mode_combo = mode_combo

        lbl_title = Label()
        lbl_title.Text = "Threshold"
        lbl_title.AutoSize = True
        lbl_title.Anchor = getattr(AnchorStyles, "None")
        lbl_title.Margin = Padding(0, 14, 0, 6)
        panel.Controls.Add(lbl_title, 0, 4)

        threshold_box = TextBox()
        threshold_box.Text = "18"
        threshold_box.Width = 72
        threshold_box.Anchor = getattr(AnchorStyles, "None")
        panel.Controls.Add(threshold_box, 0, 5)
        self._threshold_box = threshold_box

        lbl_units = Label()
        lbl_units.Text = "inches"
        lbl_units.AutoSize = True
        lbl_units.Anchor = getattr(AnchorStyles, "None")
        lbl_units.Margin = Padding(0, 6, 0, 0)
        panel.Controls.Add(lbl_units, 0, 6)

        return panel

    def _load_categories(self):
        choices = [_CategoryChoice(cat) for cat in self._categories]
        for choice in choices:
            self._combo_a.Items.Add(choice)
            self._combo_b.Items.Add(choice)

        if self._combo_a.Items.Count > 0:
            self._combo_a.SelectedIndex = 0
        if self._combo_b.Items.Count > 1:
            self._combo_b.SelectedIndex = 1
        elif self._combo_b.Items.Count > 0:
            self._combo_b.SelectedIndex = 0

    def _populate_family_list(self, combo, family_list, all_box):
        choice = combo.SelectedItem
        selectors = []
        if choice is not None:
            selectors = sorted(choice.data.get("selectors") or [], key=lambda x: x.lower())

        self._is_syncing_checks = True
        family_list.BeginUpdate()
        family_list.Items.Clear()
        for selector in selectors:
            family_list.Items.Add(selector, True)
        family_list.EndUpdate()
        all_box.Checked = family_list.Items.Count > 0
        self._is_syncing_checks = False

    def _set_all_families_checked(self, family_list, checked):
        self._is_syncing_checks = True
        family_list.BeginUpdate()
        for idx in range(family_list.Items.Count):
            family_list.SetItemChecked(idx, checked)
        family_list.EndUpdate()
        self._is_syncing_checks = False

    def _sync_all_checkbox_state(self, family_list):
        all_box = self._all_a if family_list is self._families_a else self._all_b
        total = family_list.Items.Count
        checked_total = family_list.CheckedItems.Count
        self._is_syncing_checks = True
        all_box.Checked = (total > 0 and checked_total == total)
        self._is_syncing_checks = False

    def _on_category_a_changed(self, sender, args):
        self._populate_family_list(self._combo_a, self._families_a, self._all_a)

    def _on_category_b_changed(self, sender, args):
        self._populate_family_list(self._combo_b, self._families_b, self._all_b)

    def _on_all_a_changed(self, sender, args):
        if self._is_syncing_checks:
            return
        self._set_all_families_checked(self._families_a, self._all_a.Checked)

    def _on_all_b_changed(self, sender, args):
        if self._is_syncing_checks:
            return
        self._set_all_families_checked(self._families_b, self._all_b.Checked)

    def _on_family_item_check(self, sender, args):
        if self._is_syncing_checks:
            return
        total = sender.Items.Count
        if total == 0:
            return
        checked_total = sender.CheckedItems.Count
        if args.NewValue == CheckState.Checked and args.CurrentValue != CheckState.Checked:
            checked_total += 1
        elif args.NewValue != CheckState.Checked and args.CurrentValue == CheckState.Checked:
            checked_total -= 1
        all_box = self._all_a if sender is self._families_a else self._all_b
        self._is_syncing_checks = True
        all_box.Checked = (checked_total == total)
        self._is_syncing_checks = False

    def _checked_families(self, family_list):
        selected = set()
        for item in family_list.CheckedItems:
            selected.add(str(item))
        return selected

    def _show_validation(self, message):
        MessageBox.Show(self, message, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)

    def _on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def _on_run(self, sender, args):
        choice_a = self._combo_a.SelectedItem
        choice_b = self._combo_b.SelectedItem
        if choice_a is None or choice_b is None:
            self._show_validation("Select both categories before running.")
            return

        families_a = self._checked_families(self._families_a)
        families_b = self._checked_families(self._families_b)
        if not families_a:
            self._show_validation("Select at least one family/type in Category 1.")
            return
        if not families_b:
            self._show_validation("Select at least one family/type in Category 2.")
            return

        raw_threshold = (self._threshold_box.Text or "").strip()
        try:
            threshold_inches = float(raw_threshold)
        except Exception:
            self._show_validation("Threshold distance must be numeric (inches).")
            return
        if threshold_inches < 0:
            self._show_validation("Threshold distance must be zero or greater.")
            return
        mode_idx = self._mode_combo.SelectedIndex if self._mode_combo is not None else 0
        if mode_idx < 0 or mode_idx >= len(DISTANCE_MODE_OPTIONS):
            mode_idx = 0
        distance_mode = DISTANCE_MODE_OPTIONS[mode_idx][1]

        self.result = {
            "cat_a": choice_a.data,
            "cat_b": choice_b.data,
            "families_a": families_a,
            "families_b": families_b,
            "threshold_inches": threshold_inches,
            "distance_mode": distance_mode,
        }
        self.DialogResult = DialogResult.OK
        self.Close()


def _family_name(elem):
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        name = getattr(family, "Name", None) if family else None
        if name:
            return name
    except Exception:
        pass
    return "<Unknown Family>"


def _param_text(elem, bip):
    try:
        param = elem.get_Parameter(bip)
    except Exception:
        param = None
    if param is None:
        return None
    try:
        text = param.AsString()
        if text:
            return text
    except Exception:
        pass
    try:
        text = param.AsValueString()
        if text:
            return text
    except Exception:
        pass
    return None


def _param_double(elem, bip):
    try:
        param = elem.get_Parameter(bip)
    except Exception:
        param = None
    if param is None:
        return None
    try:
        if not param.HasValue:
            return None
    except Exception:
        pass
    try:
        value = param.AsDouble()
        if value is not None and value > 0:
            return float(value)
    except Exception:
        return None
    return None


def _first_positive_param(elem, bips):
    for bip in bips or ():
        value = _param_double(elem, bip)
        if value is not None and value > 0:
            return value
    return None


def _type_name(elem, doc):
    # Instance-level type display name (works for many system elements).
    type_text = _param_text(elem, DB.BuiltInParameter.ELEM_TYPE_PARAM)
    if type_text:
        return type_text

    # Family instances can read the type name directly from Symbol.
    if isinstance(elem, DB.FamilyInstance):
        try:
            symbol = getattr(elem, "Symbol", None)
            name = getattr(symbol, "Name", None) if symbol else None
            if name:
                return name
        except Exception:
            pass

    # System elements (e.g. pipes) expose the type via GetTypeId.
    try:
        type_id = elem.GetTypeId()
    except Exception:
        type_id = None
    if type_id is not None and type_id != DB.ElementId.InvalidElementId:
        try:
            elem_type = doc.GetElement(type_id)
            if elem_type is not None:
                name = getattr(elem_type, "Name", None)
                if name:
                    return name
                # Fallbacks for type elements where Name is not exposed cleanly.
                name = _param_text(elem_type, DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if name:
                    return name
                name = _param_text(elem_type, DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if name:
                    return name
        except Exception:
            pass
    return "<Unknown Type>"


def _family_type_label(elem, doc):
    fam = _family_name(elem)
    typ = _type_name(elem, doc)
    if typ:
        return "{} : {}".format(fam, typ)
    return fam


def _iter_supported_elements(doc):
    # Loadable-component instances (families)
    try:
        for elem in (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilyInstance)
            .WhereElementIsNotElementType()
        ):
            yield elem
    except Exception:
        pass

    # System-curve categories (pipes, ducts, conduit, trays)
    for bic in SYSTEM_CATEGORY_BICS:
        try:
            collector = (
                DB.FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for elem in collector:
            yield elem


def _selection_key(elem, doc):
    if isinstance(elem, DB.FamilyInstance):
        key = _family_name(elem)
    else:
        key = _type_name(elem, doc)
        if not key or str(key).startswith("<Unknown"):
            # Keep system categories visible even when type text cannot be resolved.
            try:
                type_id = elem.GetTypeId()
                if type_id is not None and type_id != DB.ElementId.InvalidElementId:
                    return "TypeId {}".format(type_id.IntegerValue)
            except Exception:
                pass
            return None
    if not key:
        return None
    return key


def _element_label(elem, doc):
    if isinstance(elem, DB.FamilyInstance):
        return _family_type_label(elem, doc)
    cat_name = getattr(getattr(elem, "Category", None), "Name", None) or "<Category>"
    typ_name = _type_name(elem, doc)
    if typ_name and not str(typ_name).startswith("<Unknown"):
        return "{} : {}".format(cat_name, typ_name)
    return cat_name


def _bbox_distance(bbox_a, bbox_b, distance_mode):
    if bbox_a is None or bbox_b is None:
        return None

    def axis_distance(a_min, a_max, b_min, b_max):
        if a_max < b_min:
            return b_min - a_max
        if b_max < a_min:
            return a_min - b_max
        return 0.0

    dx = axis_distance(bbox_a.Min.X, bbox_a.Max.X, bbox_b.Min.X, bbox_b.Max.X)
    dy = axis_distance(bbox_a.Min.Y, bbox_a.Max.Y, bbox_b.Min.Y, bbox_b.Max.Y)
    dz = 0.0 if distance_mode == DISTANCE_MODE_XY else axis_distance(
        bbox_a.Min.Z, bbox_a.Max.Z, bbox_b.Min.Z, bbox_b.Max.Z
    )
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def _build_bbox(min_pt, max_pt):
    bb = DB.BoundingBoxXYZ()
    bb.Min = min_pt
    bb.Max = max_pt
    return bb


def _expanded_bbox(min_pt, max_pt, pad):
    if pad is None or pad <= 0:
        return _build_bbox(min_pt, max_pt)
    min_pad = DB.XYZ(min_pt.X - pad, min_pt.Y - pad, min_pt.Z - pad)
    max_pad = DB.XYZ(max_pt.X + pad, max_pt.Y + pad, max_pt.Z + pad)
    return _build_bbox(min_pad, max_pad)


def _curve_half_size_ft(elem):
    cat_name = getattr(getattr(elem, "Category", None), "Name", "") or ""
    name_low = cat_name.lower()

    if "pipe" in name_low:
        dia = _first_positive_param(elem, PIPE_SIZE_BIPS)
        if dia:
            return dia * 0.5
    if "duct" in name_low:
        dia = _first_positive_param(elem, DUCT_DIAMETER_BIPS)
        if dia:
            return dia * 0.5
        dims = [_param_double(elem, bip) for bip in DUCT_RECT_BIPS]
        dims = [d for d in dims if d is not None and d > 0]
        if dims:
            return max(dims) * 0.5
    if "conduit" in name_low:
        dia = _first_positive_param(elem, CONDUIT_SIZE_BIPS)
        if dia:
            return dia * 0.5
    if "cable tray" in name_low:
        dims = [_param_double(elem, bip) for bip in CABLETRAY_SIZE_BIPS]
        dims = [d for d in dims if d is not None and d > 0]
        if dims:
            return max(dims) * 0.5
    return 0.0


def _get_element_bbox(elem):
    size_pad = _curve_half_size_ft(elem)

    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is not None:
        if size_pad and size_pad > 0:
            return _expanded_bbox(bbox.Min, bbox.Max, size_pad)
        return bbox

    # Fallback for elements that expose only location geometry.
    try:
        loc = getattr(elem, "Location", None)
    except Exception:
        loc = None
    if loc is None:
        return None

    try:
        curve = getattr(loc, "Curve", None)
    except Exception:
        curve = None
    if curve is not None:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            min_pt = DB.XYZ(min(p0.X, p1.X), min(p0.Y, p1.Y), min(p0.Z, p1.Z))
            max_pt = DB.XYZ(max(p0.X, p1.X), max(p0.Y, p1.Y), max(p0.Z, p1.Z))
            return _expanded_bbox(min_pt, max_pt, size_pad)
        except Exception:
            pass

    try:
        point = getattr(loc, "Point", None)
    except Exception:
        point = None
    if point is not None:
        try:
            return _expanded_bbox(point, point, size_pad)
        except Exception:
            pass
    return None


def _collect_category_family_index(doc):
    category_data = {}
    for elem in _iter_supported_elements(doc):
        cat = getattr(elem, "Category", None)
        if cat is None:
            continue
        cat_name = getattr(cat, "Name", None)
        cat_id = getattr(cat, "Id", None)
        if not cat_name or cat_id is None:
            continue
        selector = _selection_key(elem, doc)
        if not selector:
            continue
        key = cat_id.IntegerValue
        if key not in category_data:
            category_data[key] = {
                "name": cat_name,
                "id": cat_id,
                "selectors": set(),
            }
        category_data[key]["selectors"].add(selector)
    return category_data


def _select_configuration(category_data):
    categories = sorted(category_data.values(), key=lambda x: (x.get("name") or "").lower())
    dialog = _GeneralProximityDialog(categories)
    try:
        result = dialog.ShowDialog()
        if result != DialogResult.OK:
            return None
        return dialog.result
    finally:
        dialog.Dispose()


def _collect_group_instances(doc, category_id, selected_selectors):
    items = []
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategoryId(category_id)
        .WhereElementIsNotElementType()
    )
    for elem in collector:
        selector = _selection_key(elem, doc)
        if not selector or selector not in selected_selectors:
            continue
        bbox = _get_element_bbox(elem)
        if bbox is None:
            continue
        items.append(
            {
                "id": elem.Id,
                "selector": selector,
                "label": _element_label(elem, doc),
                "bbox": bbox,
            }
        )
    return items


def _find_hits(instances_a, instances_b, threshold_feet, distance_mode):
    hits = []
    seen_pairs = set()
    for item_a in instances_a:
        aid = item_a["id"].IntegerValue
        bbox_a = item_a["bbox"]
        for item_b in instances_b:
            bid = item_b["id"].IntegerValue
            if aid == bid:
                continue
            pair_key = (aid, bid) if aid < bid else (bid, aid)
            if pair_key in seen_pairs:
                continue
            dist_ft = _bbox_distance(bbox_a, item_b["bbox"], distance_mode)
            if dist_ft is None or dist_ft > threshold_feet:
                continue
            seen_pairs.add(pair_key)
            hits.append(
                {
                    "a_id": item_a["id"],
                    "a_label": item_a["label"],
                    "b_id": item_b["id"],
                    "b_label": item_b["label"],
                    "distance_ft": dist_ft,
                }
            )
    return hits


def _mode_label(distance_mode):
    return "Plan (XY only)" if distance_mode == DISTANCE_MODE_XY else "3D"


def _report_results(hits, cat_a_name, cat_b_name, fam_a_count, fam_b_count, threshold_inches, distance_mode):
    output = script.get_output()
    output.set_width(1200)
    output.print_md("# General Proximity Check")
    output.print_md(
        "**Category A:** {} ({} selected families/types)  \n"
        "**Category B:** {} ({} selected families/types)  \n"
        "**Distance Mode:** {}  \n"
        "**Threshold:** {:.2f} in".format(
            cat_a_name,
            fam_a_count,
            cat_b_name,
            fam_b_count,
            _mode_label(distance_mode),
            threshold_inches,
        )
    )

    if not hits:
        output.print_md("### No issues found.")
        forms.alert(
            "No qualifying instances found within {:.2f} inches.".format(threshold_inches),
            title=TITLE,
        )
        return

    rows = []
    for hit in hits:
        rows.append(
            [
                output.linkify(hit["a_id"]),
                hit["a_label"],
                output.linkify(hit["b_id"]),
                hit["b_label"],
                "{:.2f}".format(hit["distance_ft"] * 12.0),
            ]
        )
    output.print_table(
        rows,
        columns=[
            "{} ID".format(cat_a_name),
            "{} Family/Type".format(cat_a_name),
            "{} ID".format(cat_b_name),
            "{} Family/Type".format(cat_b_name),
            "Distance (in)",
        ],
    )
    forms.alert(
        "Found {} instance pair(s) within {:.2f} inches.\n\nSee the output panel for details.".format(
            len(hits), threshold_inches
        ),
        title=TITLE,
    )


def main():
    doc = getattr(revit, "doc", None)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        forms.alert("Open a project model before running this check.", title=TITLE)
        return

    category_data = _collect_category_family_index(doc)
    if not category_data:
        forms.alert("No supported placed elements were found in this model.", title=TITLE)
        return

    config = _select_configuration(category_data)
    if config is None:
        return

    cat_a = config.get("cat_a")
    cat_b = config.get("cat_b")
    selected_fams_a = config.get("families_a") or set()
    selected_fams_b = config.get("families_b") or set()
    threshold_inches = config.get("threshold_inches")
    distance_mode = config.get("distance_mode") or DISTANCE_MODE_3D
    threshold_feet = threshold_inches / 12.0

    instances_a = _collect_group_instances(doc, cat_a["id"], selected_fams_a)
    instances_b = _collect_group_instances(doc, cat_b["id"], selected_fams_b)

    if not instances_a:
        forms.alert("No valid instances found for the first category selection.", title=TITLE)
        return
    if not instances_b:
        forms.alert("No valid instances found for the second category selection.", title=TITLE)
        return

    hits = _find_hits(instances_a, instances_b, threshold_feet, distance_mode)
    _report_results(
        hits,
        cat_a["name"],
        cat_b["name"],
        len(selected_fams_a),
        len(selected_fams_b),
        threshold_inches,
        distance_mode,
    )


if __name__ == "__main__":
    main()
