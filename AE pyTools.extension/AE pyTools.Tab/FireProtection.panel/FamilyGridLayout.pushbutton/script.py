# -*- coding: utf-8 -*-
"""Place a selected family type on a grid inside one selected Space/Room."""

__title__ = "FamilyGridLayout"
__doc__ = "Place selected family type in a grid within one selected space."

import clr
import re

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Drawing import Point, Size
from System.Windows.Forms import (
    Button,
    ComboBox,
    ComboBoxStyle,
    DialogResult,
    Form,
    FormBorderStyle,
    FormStartPosition,
    Label,
    MessageBox,
    TextBox,
)

from pyrevit import DB, forms, revit, script


doc = revit.doc
logger = script.get_logger()


try:
    basestring
except NameError:
    basestring = str


class _Option(object):
    def __init__(self, element, label):
        self.element = element
        self.label = label

    def __str__(self):
        return self.label


class _SpatialTarget(object):
    def __init__(self, spatial_elem, source_doc, link_instance=None):
        self.spatial_elem = spatial_elem
        self.source_doc = source_doc
        self.link_instance = link_instance
        if link_instance is not None:
            try:
                self.to_host = link_instance.GetTotalTransform()
            except Exception:
                self.to_host = None
            if self.to_host is None:
                self.to_host = DB.Transform.Identity
            try:
                self.to_link = self.to_host.Inverse
            except Exception:
                self.to_link = DB.Transform.Identity
        else:
            self.to_host = DB.Transform.Identity
            self.to_link = DB.Transform.Identity

    @property
    def is_linked(self):
        return self.link_instance is not None

    def _host_bounds_from_bbox(self, bbox):
        if bbox is None:
            return None
        if not self.is_linked:
            return (
                bbox.Min.X,
                bbox.Min.Y,
                bbox.Min.Z,
                bbox.Max.X,
                bbox.Max.Y,
                bbox.Max.Z,
            )

        xs = (bbox.Min.X, bbox.Max.X)
        ys = (bbox.Min.Y, bbox.Max.Y)
        zs = (bbox.Min.Z, bbox.Max.Z)
        points = []
        for x_val in xs:
            for y_val in ys:
                for z_val in zs:
                    points.append(self.to_host.OfPoint(DB.XYZ(x_val, y_val, z_val)))
        min_x = min(p.X for p in points)
        min_y = min(p.Y for p in points)
        min_z = min(p.Z for p in points)
        max_x = max(p.X for p in points)
        max_y = max(p.Y for p in points)
        max_z = max(p.Z for p in points)
        return (min_x, min_y, min_z, max_x, max_y, max_z)

    def get_host_bounds(self):
        bbox = self.spatial_elem.get_BoundingBox(None)
        return self._host_bounds_from_bbox(bbox)

    def is_inside_host_point(self, host_point):
        if self.is_linked:
            test_point = self.to_link.OfPoint(host_point)
        else:
            test_point = host_point
        try:
            if self.spatial_elem.IsPointInSpace(test_point):
                return True
        except Exception:
            pass
        try:
            if self.spatial_elem.IsPointInRoom(test_point):
                return True
        except Exception:
            pass
        return False


class FamilyGridLayoutForm(Form):
    def __init__(self, spaces, symbols):
        Form.__init__(self)
        self._spaces = spaces
        self._symbols = symbols
        self.result = None

        self.Text = "FamilyGridLayout"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(760, 280)

        self._build_ui()

    def _build_ui(self):
        margin_left = 14
        label_w = 300
        input_x = 320
        input_w = 420
        row_h = 34
        top = 16

        lbl_space = Label()
        lbl_space.Text = "Space/Room:"
        lbl_space.Location = Point(margin_left, top + 4)
        lbl_space.Size = Size(label_w, 22)
        self.Controls.Add(lbl_space)

        self.cmb_space = ComboBox()
        self.cmb_space.Location = Point(input_x, top)
        self.cmb_space.Size = Size(input_w, 22)
        self.cmb_space.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_space.DropDownWidth = 700
        for opt in self._spaces:
            self.cmb_space.Items.Add(opt.label)
        if self.cmb_space.Items.Count > 0:
            self.cmb_space.SelectedIndex = 0
        self.Controls.Add(self.cmb_space)

        top += row_h

        lbl_symbol = Label()
        lbl_symbol.Text = "Family : Type:"
        lbl_symbol.Location = Point(margin_left, top + 4)
        lbl_symbol.Size = Size(label_w, 22)
        self.Controls.Add(lbl_symbol)

        self.cmb_symbol = ComboBox()
        self.cmb_symbol.Location = Point(input_x, top)
        self.cmb_symbol.Size = Size(input_w, 22)
        self.cmb_symbol.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_symbol.DropDownWidth = 700
        default_idx = 0
        for idx, opt in enumerate(self._symbols):
            self.cmb_symbol.Items.Add(opt.label)
            if "sprinkler" in opt.label.lower() and default_idx == 0:
                default_idx = idx
        if self.cmb_symbol.Items.Count > 0:
            self.cmb_symbol.SelectedIndex = default_idx
        self.Controls.Add(self.cmb_symbol)

        top += row_h

        lbl_x = Label()
        lbl_x.Text = "X spacing (feet):"
        lbl_x.Location = Point(margin_left, top + 4)
        lbl_x.Size = Size(label_w, 22)
        self.Controls.Add(lbl_x)

        self.txt_x = TextBox()
        self.txt_x.Location = Point(input_x, top)
        self.txt_x.Size = Size(120, 22)
        self.txt_x.Text = "10"
        self.Controls.Add(self.txt_x)

        top += row_h

        lbl_y = Label()
        lbl_y.Text = "Y spacing (feet):"
        lbl_y.Location = Point(margin_left, top + 4)
        lbl_y.Size = Size(label_w, 22)
        self.Controls.Add(lbl_y)

        self.txt_y = TextBox()
        self.txt_y.Location = Point(input_x, top)
        self.txt_y.Size = Size(120, 22)
        self.txt_y.Text = "10"
        self.Controls.Add(self.txt_y)

        top += row_h

        lbl_off = Label()
        lbl_off.Text = "Distance below ceiling/roof (inches):"
        lbl_off.Location = Point(margin_left, top + 4)
        lbl_off.Size = Size(label_w, 22)
        self.Controls.Add(lbl_off)

        self.txt_offset = TextBox()
        self.txt_offset.Location = Point(input_x, top)
        self.txt_offset.Size = Size(120, 22)
        self.txt_offset.Text = "10"
        self.Controls.Add(self.txt_offset)

        top += row_h + 14

        btn_ok = Button()
        btn_ok.Text = "Place Grid"
        btn_ok.Location = Point(560, top)
        btn_ok.Size = Size(90, 28)
        btn_ok.Click += self._on_ok
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Point(660, top)
        btn_cancel.Size = Size(80, 28)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def _parse_inches(self, raw_value):
        if raw_value is None:
            raise ValueError("Value is empty.")
        text = str(raw_value).strip().lower()
        text = text.replace("inches", "").replace("inch", "").replace('"', "")
        text = re.sub(r"\s+", "", text)
        if text.endswith("in"):
            text = text[:-2]
        if not text:
            raise ValueError("Value is empty.")
        return float(text)

    def _parse_feet(self, raw_value):
        if raw_value is None:
            raise ValueError("Value is empty.")
        text = str(raw_value).strip().lower()
        text = text.replace("feet", "").replace("foot", "").replace("ft", "").replace("'", "")
        text = re.sub(r"\s+", "", text)
        if not text:
            raise ValueError("Value is empty.")
        return float(text)

    def _on_ok(self, sender, args):
        selected_space_idx = self.cmb_space.SelectedIndex
        selected_symbol_idx = self.cmb_symbol.SelectedIndex
        if selected_space_idx < 0:
            MessageBox.Show("Select a space or room.")
            return
        if selected_symbol_idx < 0:
            MessageBox.Show("Select a family type.")
            return
        try:
            x_ft = self._parse_feet(self.txt_x.Text)
            y_ft = self._parse_feet(self.txt_y.Text)
            offset_in = self._parse_inches(self.txt_offset.Text)
        except Exception:
            MessageBox.Show("X/Y spacing must be numeric feet values. Offset must be inches.")
            return
        if x_ft <= 0 or y_ft <= 0:
            MessageBox.Show("X and Y spacing must be greater than zero.")
            return
        if offset_in < 0:
            MessageBox.Show("Offset below ceiling/roof cannot be negative.")
            return

        self.result = {
            "space": self._spaces[selected_space_idx].element,
            "symbol": self._symbols[selected_symbol_idx].element,
            "x_spacing_ft": x_ft,
            "y_spacing_ft": y_ft,
            "offset_in": offset_in,
        }
        self.DialogResult = DialogResult.OK
        self.Close()


def _safe_text(value):
    if value is None:
        return ""
    if isinstance(value, basestring):
        return value.strip()
    return str(value).strip()


def _get_level_name(elem):
    if elem is None:
        return ""
    return _safe_text(getattr(elem, "Name", ""))


def _get_type_name(symbol):
    name = _safe_text(getattr(symbol, "Name", ""))
    if name:
        return name
    try:
        p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return _safe_text(p.AsString()) or _safe_text(p.AsValueString())
    except Exception:
        pass
    return "<Unnamed Type>"


def _spatial_display_base(kind, spatial_elem):
    number = _safe_text(getattr(spatial_elem, "Number", ""))
    name = _safe_text(getattr(spatial_elem, "Name", ""))
    if number and name:
        return "{} {} - {}".format(kind, number, name)
    if number:
        return "{} {}".format(kind, number)
    if name:
        return "{} {}".format(kind, name)
    return "{} <Unnamed>".format(kind)


def _collect_space_options():
    options = []
    for bic, kind in (
        (DB.BuiltInCategory.OST_MEPSpaces, "Space"),
        (DB.BuiltInCategory.OST_Rooms, "Room"),
    ):
        collector = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        for elem in collector:
            try:
                if hasattr(elem, "Area") and elem.Area <= 0:
                    continue
            except Exception:
                pass
            bbox = elem.get_BoundingBox(None)
            if bbox is None:
                continue
            level_name = ""
            try:
                lvl = doc.GetElement(elem.LevelId)
                level_name = _get_level_name(lvl)
            except Exception:
                level_name = ""
            base = _spatial_display_base(kind, elem)
            label = "{} [{}]".format(base, level_name) if level_name else base
            options.append(_Option(_SpatialTarget(elem, doc), label))

    link_instances = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.RevitLinkInstance)
        .WhereElementIsNotElementType()
    )
    for link_inst in link_instances:
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        link_name = _safe_text(getattr(link_inst, "Name", "")) or _safe_text(getattr(link_doc, "Title", ""))
        for bic, kind in (
            (DB.BuiltInCategory.OST_MEPSpaces, "Space"),
            (DB.BuiltInCategory.OST_Rooms, "Room"),
        ):
            linked_collector = (
                DB.FilteredElementCollector(link_doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            )
            for linked_elem in linked_collector:
                try:
                    if hasattr(linked_elem, "Area") and linked_elem.Area <= 0:
                        continue
                except Exception:
                    pass
                bbox = linked_elem.get_BoundingBox(None)
                if bbox is None:
                    continue
                level_name = ""
                try:
                    lvl = link_doc.GetElement(linked_elem.LevelId)
                    level_name = _get_level_name(lvl)
                except Exception:
                    level_name = ""
                base = _spatial_display_base(kind, linked_elem)
                level_label = " [{}]".format(level_name) if level_name else ""
                link_label = " (Linked: {})".format(link_name) if link_name else " (Linked)"
                label = "{}{}{}".format(base, level_label, link_label)
                options.append(_Option(_SpatialTarget(linked_elem, link_doc, link_inst), label))
    options.sort(key=lambda x: x.label.lower())
    return options


def _collect_symbol_options():
    options = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    for symbol in collector:
        category = getattr(symbol, "Category", None)
        if category is None:
            continue
        if category.CategoryType != DB.CategoryType.Model:
            continue
        category_name = _safe_text(category.Name) or "Model"
        family_name = _safe_text(getattr(symbol, "FamilyName", ""))
        if not family_name:
            try:
                family_name = _safe_text(symbol.Family.Name)
            except Exception:
                family_name = "<Unnamed Family>"
        type_name = _get_type_name(symbol)
        placement_type = "UnknownPlacement"
        try:
            placement_type = _safe_text(symbol.Family.FamilyPlacementType)
        except Exception:
            placement_type = "UnknownPlacement"
        label = "{} | {} : {} [{}]".format(
            category_name, family_name, type_name, placement_type
        )
        options.append(_Option(symbol, label))
    options.sort(key=lambda x: x.label.lower())
    return options


def _axis_values(min_value, max_value, spacing_ft):
    values = []
    half = spacing_ft / 2.0
    usable_min = min_value + half
    usable_max = max_value - half
    if usable_min <= usable_max:
        v = usable_min
        while v <= usable_max + 1e-9:
            values.append(v)
            v += spacing_ft
    if not values:
        values.append((min_value + max_value) / 2.0)
    return values


def _grid_points_in_space(spatial_target, x_spacing_ft, y_spacing_ft, offset_ft):
    bounds = spatial_target.get_host_bounds()
    if bounds is None:
        return None, "Selected space/room has no bounding box."

    min_x, min_y, min_z, max_x, max_y, max_z = bounds

    if max_x <= min_x or max_y <= min_y:
        return None, "Selected space/room has invalid bounds."

    place_z = max_z - offset_ft
    if place_z <= min_z + 1e-6:
        return None, "Offset is too large for selected space/room height."

    sample_z = min_z + ((max_z - min_z) * 0.5)
    x_values = _axis_values(min_x, max_x, x_spacing_ft)
    y_values = _axis_values(min_y, max_y, y_spacing_ft)

    points = []
    for x_val in x_values:
        for y_val in y_values:
            test_point = DB.XYZ(x_val, y_val, sample_z)
            if spatial_target.is_inside_host_point(test_point):
                points.append(DB.XYZ(x_val, y_val, place_z))

    if not points:
        center_test = DB.XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, sample_z)
        if spatial_target.is_inside_host_point(center_test):
            points = [DB.XYZ(center_test.X, center_test.Y, place_z)]
        else:
            return None, "No valid grid points were found inside the selected space/room."

    return points, None


def _level_for_spatial(spatial_target):
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    if not levels:
        return None
    target_z = None
    if not spatial_target.is_linked:
        try:
            lvl_id = spatial_target.spatial_elem.LevelId
            if lvl_id and lvl_id != DB.ElementId.InvalidElementId:
                lvl = doc.GetElement(lvl_id)
                if lvl is not None:
                    return lvl
        except Exception:
            pass
    else:
        try:
            src_lvl_id = spatial_target.spatial_elem.LevelId
            if src_lvl_id and src_lvl_id != DB.ElementId.InvalidElementId:
                src_lvl = spatial_target.source_doc.GetElement(src_lvl_id)
                if src_lvl is not None:
                    host_point = spatial_target.to_host.OfPoint(DB.XYZ(0, 0, src_lvl.Elevation))
                    target_z = host_point.Z
        except Exception:
            target_z = None

    if target_z is None:
        bounds = spatial_target.get_host_bounds()
        target_z = bounds[2] if bounds else 0.0
    levels.sort(key=lambda l: abs(l.Elevation - target_z))
    return levels[0]


def _set_instance_offset(instance, level, target_z):
    if instance is None or level is None:
        return False
    offset = target_z - level.Elevation
    for bip in (
        DB.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
        DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM,
    ):
        try:
            p = instance.get_Parameter(bip)
            if p and (not p.IsReadOnly) and p.StorageType == DB.StorageType.Double:
                p.Set(offset)
                return True
        except Exception:
            pass
    return False


def _try_place(symbol, point, level):
    try:
        return (
            doc.Create.NewFamilyInstance(
                point,
                symbol,
                DB.Structure.StructuralType.NonStructural,
            ),
            None,
        )
    except Exception as ex_point:
        point_msg = _safe_text(ex_point)

    if level is not None:
        try:
            inst = doc.Create.NewFamilyInstance(
                point,
                symbol,
                level,
                DB.Structure.StructuralType.NonStructural,
            )
            _set_instance_offset(inst, level, point.Z)
            return inst, None
        except Exception as ex_level:
            level_msg = _safe_text(ex_level)
    else:
        level_msg = "No valid level was found for the selected space."

    return None, "{} | {}".format(point_msg, level_msg)


def _run():
    spaces = _collect_space_options()
    if not spaces:
        forms.alert("No placed Spaces or Rooms were found in this model.", exitscript=True)

    symbols = _collect_symbol_options()
    if not symbols:
        forms.alert("No model Family Types were found in this model.", exitscript=True)

    dialog = FamilyGridLayoutForm(spaces, symbols)
    result = dialog.ShowDialog()
    if result != DialogResult.OK or not dialog.result:
        script.exit()

    selected_space = dialog.result["space"]
    selected_symbol = dialog.result["symbol"]
    x_spacing_ft = dialog.result["x_spacing_ft"]
    y_spacing_ft = dialog.result["y_spacing_ft"]
    offset_ft = dialog.result["offset_in"] / 12.0

    points, err = _grid_points_in_space(
        selected_space, x_spacing_ft, y_spacing_ft, offset_ft
    )
    if err:
        forms.alert(err, title="FamilyGridLayout", exitscript=True)

    if len(points) > 5000:
        should_continue = forms.alert(
            "This will place {} instances. Continue?".format(len(points)),
            yes=True,
            no=True,
        )
        if not should_continue:
            script.exit()

    level = _level_for_spatial(selected_space)
    placed_count = 0
    skipped_count = 0
    first_skip_error = None

    tx = DB.Transaction(doc, "Family Grid Layout")
    try:
        tx.Start()
        if not selected_symbol.IsActive:
            selected_symbol.Activate()
            doc.Regenerate()

        first_instance, first_error = _try_place(selected_symbol, points[0], level)
        if first_instance is None:
            raise Exception(
                "Selected family type could not be placed at the computed point.\n{}".format(
                    first_error
                )
            )
        placed_count += 1

        for point in points[1:]:
            inst, place_err = _try_place(selected_symbol, point, level)
            if inst is not None:
                placed_count += 1
            else:
                skipped_count += 1
                if first_skip_error is None:
                    first_skip_error = place_err

        tx.Commit()
    except Exception as place_ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        forms.alert(
            "FamilyGridLayout failed and no instances were placed.\n\n{}".format(place_ex),
            title="FamilyGridLayout",
            exitscript=True,
        )

    summary = "Placed {} instance(s).".format(placed_count)
    if skipped_count:
        summary += "\nSkipped {} point(s).".format(skipped_count)
        if first_skip_error:
            summary += "\nFirst skip reason: {}".format(first_skip_error)
    forms.alert(summary, title="FamilyGridLayout")


if __name__ == "__main__":
    _run()
