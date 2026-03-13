# -*- coding: utf-8 -*-
__title__ = "Place all Coils"
__doc__ = "Place coil families in spaces based on an Excel description list."

import re
from difflib import SequenceMatcher

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System
from System import Type, Activator
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from System.Windows.Forms import (
    Form,
    DataGridView,
    DataGridViewTextBoxColumn,
    DataGridViewComboBoxColumn,
    DockStyle,
    FormStartPosition,
    DataGridViewAutoSizeColumnsMode,
    DialogResult,
    Button,
)
from System.Drawing import Size, Point

from pyrevit import revit, DB, forms, script
from pyrevit.revit import query


logger = script.get_logger()
doc = revit.doc

VERTICAL_OFFSET_FT = 2.0
SPACE_MATCH_THRESHOLD = 0.6
MODEL_MATCH_THRESHOLD = 0.45
SHEET_NAME = "Circuit Schedule"
HEADER_SCAN_ROWS = 60
REQUIRED_MANUFACTURER = "KRACK"
SKIP_SPACE_LABEL = "<Do not place in any space>"

DESC_KEYS = ("description", "desc", "space", "spacename")
COUNT_KEYS = ("coilcount", "coils", "coil", "count")
MODEL_KEYS = ("model", "modelnumber", "modelno")
MFR_KEYS = ("manufacturer", "mfr", "mfg")


try:
    basestring
except NameError:
    basestring = str


def _args_array(*args):
    return System.Array[System.Object](list(args))


def _set(obj, prop, val):
    obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))


def _get(obj, prop):
    try:
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except Exception:
        return None


def _call(obj, name, *args):
    t = obj.GetType()
    try:
        return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
    except Exception:
        try:
            return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args) if args else None)
        except Exception:
            return None


def _cell(cells, r, c):
    it = _call(cells, "Item", r, c)
    v = _get(it, "Value2")
    return ("" if v is None else str(v)).strip()


def _norm_key(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"[^0-9a-zA-Z]+", "", text).lower()
    return text


def _norm_text(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"[^0-9a-zA-Z]+", " ", text).strip().lower()
    return re.sub(r"\s+", " ", text)


def _text_similarity(a, b):
    na = _norm_text(a)
    nb = _norm_text(b)
    if not na or not nb:
        return 0.0
    if na in nb or nb in na:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def _find_header_columns(cells, nrows, ncols):
    desc_col = coil_col = model_col = mfr_col = None
    desc_row = coil_row = model_row = mfr_row = 0
    desc_score = coil_score = model_score = mfr_score = 0
    max_rows = min(nrows, HEADER_SCAN_ROWS)

    for r in range(1, max_rows + 1):
        for c in range(1, ncols + 1):
            raw = _cell(cells, r, c)
            if not raw:
                continue
            raw_l = raw.strip().lower()
            compact = re.sub(r"\s+", "", raw_l)
            if "description" in raw_l:
                score = 2 if raw_l == "description" else 1
                if score > desc_score:
                    desc_col, desc_row, desc_score = c, r, score
            if "coil count" in raw_l:
                score = 2 if raw_l == "coil count" else 1
                if score > coil_score:
                    coil_col, coil_row, coil_score = c, r, score
            if "model#" in compact or compact == "model":
                score = 2 if "model#" in compact or compact == "model#" else 1
                if score > model_score:
                    model_col, model_row, model_score = c, r, score
            if "manufacturer" in raw_l:
                score = 2 if raw_l == "manufacturer" else 1
                if score > mfr_score:
                    mfr_col, mfr_row, mfr_score = c, r, score
        if desc_score == 2 and coil_score == 2 and model_score == 2 and mfr_score == 2:
            break

    return desc_col, desc_row, coil_col, coil_row, model_col, model_row, mfr_col, mfr_row


def _load_circuit_schedule_rows(path):
    xl = wb = ws = used = cells = rows_prop = cols_prop = None
    rows = []
    try:
        t = Type.GetTypeFromProgID("Excel.Application")
        if t is None:
            raise Exception("Excel is not registered on this machine.")
        xl = Activator.CreateInstance(t)
        _set(xl, "Visible", False)
        _set(xl, "DisplayAlerts", False)
        wb = _call(_get(xl, "Workbooks"), "Open", path)
        ws = _call(_get(wb, "Worksheets"), "Item", SHEET_NAME)
        if ws is None:
            raise Exception("Sheet not found: {}".format(SHEET_NAME))

        used = _get(ws, "UsedRange")
        cells = _get(used, "Cells")
        rows_prop = _get(used, "Rows")
        cols_prop = _get(used, "Columns")
        nrows = int(_get(rows_prop, "Count") or 0)
        ncols = int(_get(cols_prop, "Count") or 0)

        desc_col, desc_row, coil_col, coil_row, model_col, model_row, mfr_col, mfr_row = _find_header_columns(
            cells, nrows, ncols
        )
        if not (desc_col and coil_col and model_col and mfr_col):
            raise Exception(
                "Could not locate Description / Coil Count / Model # / Manufacturer columns on '{}'.".format(
                    SHEET_NAME
                )
            )

        start_row = max(desc_row, coil_row, model_row, mfr_row) + 1
        for r in range(start_row, nrows + 1):
            desc = _cell(cells, r, desc_col)
            coil = _cell(cells, r, coil_col)
            model = _cell(cells, r, model_col)
            mfr = _cell(cells, r, mfr_col)
            if not desc and not coil and not model:
                continue
            rows.append({
                "Description": desc,
                "Coil Count": coil,
                "Model #": model,
                "Manufacturer": mfr,
                "_row": r,
            })
    finally:
        try:
            if wb:
                _call(wb, "Close", False)
            if xl:
                _call(xl, "Quit")
        except Exception:
            pass
        try:
            if ws:
                Marshal.ReleaseComObject(ws)
            if wb:
                Marshal.ReleaseComObject(wb)
            if xl:
                Marshal.ReleaseComObject(xl)
        except Exception:
            pass
    return rows


def _get_location_point(elem):
    loc = getattr(elem, "Location", None)
    if loc and hasattr(loc, "Point"):
        return loc.Point
    return None


def _get_bbox_center(elem):
    bbox = None
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(revit.active_view)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _find_spatial_element(doc, point):
    if not point:
        return None
    getter = getattr(doc, "GetSpaceAtPoint", None)
    if callable(getter):
        try:
            space = getter(point)
            if space:
                return space
        except Exception:
            pass
    getter = getattr(doc, "GetRoomAtPoint", None)
    if callable(getter):
        try:
            room = getter(point)
            if room:
                return room
        except Exception:
            pass
    return None


def _find_linked_spatial_element(link_doc, point):
    if not point:
        return None
    getter = getattr(link_doc, "GetSpaceAtPoint", None)
    if callable(getter):
        try:
            space = getter(point)
            if space:
                return space
        except Exception:
            pass
    getter = getattr(link_doc, "GetRoomAtPoint", None)
    if callable(getter):
        try:
            room = getter(point)
            if room:
                return room
        except Exception:
            pass
    return None


def _get_param_string(elem, bip):
    try:
        param = elem.get_Parameter(bip)
    except Exception:
        param = None
    if not param:
        return ""
    try:
        val = param.AsString()
        if val:
            return val
    except Exception:
        pass
    try:
        val = param.AsValueString()
        if val:
            return val
    except Exception:
        pass
    return ""


def _space_name(space):
    name = getattr(space, "Name", None)
    if not name:
        name = _get_param_string(space, DB.BuiltInParameter.ROOM_NAME)
    if not name:
        name = _get_param_string(space, DB.BuiltInParameter.SPACE_NAME)
    return (name or "").strip()


def _space_number(space):
    number = getattr(space, "Number", None)
    if not number:
        number = _get_param_string(space, DB.BuiltInParameter.ROOM_NUMBER)
    if not number:
        number = _get_param_string(space, DB.BuiltInParameter.SPACE_NUMBER)
    return (number or "").strip()


def _collect_host_spaces():
    spaces = []
    seen = set()
    categories = [DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Rooms]
    for cat in categories:
        elements = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
        for space in elements:
            sid = space.Id.IntegerValue
            if sid in seen:
                continue
            seen.add(sid)
            name = _space_name(space)
            number = _space_number(space)
            if not name and not number:
                continue
            point = _get_bbox_center(space) or _get_location_point(space)
            spaces.append({
                "element": space,
                "name": name,
                "number": number,
                "point": point,
                "source": "Host",
                "link": None,
                "link_doc": None,
                "link_name": None,
                "key": "host:{}".format(sid),
                "level_name": None,
            })
    return spaces


def _collect_linked_spaces():
    spaces = []
    categories = [DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Rooms]
    link_instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
    for link in link_instances:
        link_doc = link.GetLinkDocument()
        if link_doc is None:
            continue
        transform = link.GetTransform()
        link_name = getattr(link, "Name", None) or "Link {}".format(link.Id.IntegerValue)
        for cat in categories:
            elements = DB.FilteredElementCollector(link_doc).OfCategory(cat).WhereElementIsNotElementType()
            for space in elements:
                name = _space_name(space)
                number = _space_number(space)
                if not name and not number:
                    continue
                point = _get_bbox_center(space) or _get_location_point(space)
                host_point = transform.OfPoint(point) if point else None
                level_name = None
                try:
                    level = getattr(space, "Level", None)
                    if level is not None:
                        level_name = level.Name
                    elif space.LevelId and space.LevelId != DB.ElementId.InvalidElementId:
                        level_elem = link_doc.GetElement(space.LevelId)
                        level_name = level_elem.Name if level_elem else None
                except Exception:
                    level_name = None
                space_key = "{}:{}".format(link.Id.IntegerValue, space.Id.IntegerValue)
                spaces.append({
                    "element": space,
                    "name": name,
                    "number": number,
                    "point": host_point,
                    "source": "Linked",
                    "link": link,
                    "link_doc": link_doc,
                    "link_name": link_name,
                    "key": space_key,
                    "level_name": level_name,
                })
    return spaces


def _collect_all_spaces():
    return _collect_linked_spaces() + _collect_host_spaces()


def _space_keys(space):
    keys = []
    if space.get("name"):
        keys.append(space["name"])
    if space.get("number"):
        keys.append(space["number"])
    if space.get("name") and space.get("number"):
        keys.append("{} {}".format(space["number"], space["name"]))
    return keys


def _space_display(space):
    name = space.get("name") or ""
    number = space.get("number") or ""
    label = name or number or "Unnamed Space"
    if name and number:
        label = "{} ({})".format(name, number)
    source = space.get("source") or "Host"
    if source == "Linked":
        link_name = space.get("link_name") or "Link"
        return "{} [Linked: {}]".format(label, link_name)
    return "{} [Host]".format(label)


def _prompt_space_mapping(descriptions, spaces):
    if not descriptions:
        return {}

    label_map = {}
    labels = []
    for space in spaces:
        label = _space_display(space)
        if label in label_map:
            label = "{} ({})".format(label, space.get("key"))
        label_map[label] = space
        labels.append(label)
    labels = [SKIP_SPACE_LABEL] + sorted(labels, key=lambda s: s.lower())

    form = Form()
    form.Text = "Map Descriptions to Spaces"
    form.Size = Size(900, 600)
    form.StartPosition = FormStartPosition.CenterScreen

    grid = DataGridView()
    grid.Dock = DockStyle.Top
    grid.Height = 520
    grid.AllowUserToAddRows = False
    grid.AllowUserToDeleteRows = False
    grid.ReadOnly = False
    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    col_desc = DataGridViewTextBoxColumn()
    col_desc.HeaderText = "Description"
    col_desc.ReadOnly = True
    grid.Columns.Add(col_desc)

    col_space = DataGridViewComboBoxColumn()
    col_space.HeaderText = "Space (Linked/Host)"
    col_space.DataSource = labels
    grid.Columns.Add(col_space)

    for desc in descriptions:
        idx = grid.Rows.Add()
        grid.Rows[idx].Cells[0].Value = desc
        grid.Rows[idx].Cells[1].Value = SKIP_SPACE_LABEL

    ok_btn = Button()
    ok_btn.Text = "OK"
    ok_btn.Size = Size(100, 30)
    ok_btn.Location = Point(680, 530)
    ok_btn.DialogResult = DialogResult.OK

    cancel_btn = Button()
    cancel_btn.Text = "Cancel"
    cancel_btn.Size = Size(100, 30)
    cancel_btn.Location = Point(790, 530)
    cancel_btn.DialogResult = DialogResult.Cancel

    form.Controls.Add(grid)
    form.Controls.Add(ok_btn)
    form.Controls.Add(cancel_btn)
    form.AcceptButton = ok_btn
    form.CancelButton = cancel_btn

    if form.ShowDialog() != DialogResult.OK:
        return None

    mapping = {}
    for row in grid.Rows:
        desc_val = row.Cells[0].Value
        space_val = row.Cells[1].Value
        if desc_val is None:
            continue
        desc_text = str(desc_val)
        if space_val is None:
            mapping[desc_text] = None
            continue
        space_label = str(space_val)
        if space_label == SKIP_SPACE_LABEL:
            mapping[desc_text] = None
            continue
        mapping[desc_text] = label_map.get(space_label)

    return mapping


def _best_space_match(description, spaces):
    best_space = None
    best_score = 0.0
    for space in spaces:
        for key in _space_keys(space):
            score = _text_similarity(description, key)
            if score > best_score:
                best_score = score
                best_space = space
    return best_space, best_score


def _collect_mech_symbols():
    symbols = []
    cat_id = int(DB.BuiltInCategory.OST_MechanicalEquipment)
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        try:
            if symbol.Category and symbol.Category.Id.IntegerValue == cat_id:
                symbols.append(symbol)
        except Exception:
            continue
    return symbols


def _collect_symbols_by_space(spaces):
    symbol_map = {}
    space_key_lookup = {}
    link_instances = {}
    for space in spaces:
        link = space.get("link")
        link_doc = space.get("link_doc")
        if not link or not link_doc:
            continue
        link_id = link.Id.IntegerValue
        link_instances[link_id] = link
        space_key_lookup[(link_id, space["element"].Id.IntegerValue)] = space.get("key")

    if not link_instances:
        return symbol_map

    cat = DB.BuiltInCategory.OST_MechanicalEquipment
    elems = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
    for elem in elems:
        point = _get_location_point(elem)
        if not point:
            continue
        symbol = getattr(elem, "Symbol", None)
        if not symbol:
            continue
        for link_id, link in link_instances.items():
            link_doc = link.GetLinkDocument()
            if link_doc is None:
                continue
            transform = link.GetTransform()
            inv = transform.Inverse
            link_point = inv.OfPoint(point)
            spatial = _find_linked_spatial_element(link_doc, link_point)
            if not spatial:
                continue
            key = space_key_lookup.get((link_id, spatial.Id.IntegerValue))
            if not key:
                continue
            symbol_map.setdefault(key, set()).add(symbol)
            break
    return symbol_map


def _best_symbol_match(model, symbols):
    if not model:
        return None, 0.0
    keys = _model_keys(model)
    best_symbol = None
    best_score = 0.0
    for symbol in symbols:
        try:
            fam_name = query.get_name(symbol.Family)
        except Exception:
            fam_name = ""
        try:
            type_name = query.get_name(symbol)
        except Exception:
            type_name = ""
        if REQUIRED_MANUFACTURER:
            mfr_u = REQUIRED_MANUFACTURER.upper()
            if mfr_u not in (fam_name or "").upper() and mfr_u not in (type_name or "").upper():
                continue
        label = "{} : {}".format(fam_name, type_name)
        label_key = _norm_model_key(label)
        score = 0.0
        if keys:
            if any(k and k in label_key for k in keys):
                score = 1.0
        else:
            score = max(
                _text_similarity(model, label),
                _text_similarity(model, type_name),
            )
        if score > best_score:
            best_score = score
            best_symbol = symbol
    return best_symbol, best_score


def _resolve_level(space):
    level_name = None
    if isinstance(space, dict):
        level_name = space.get("level_name")
        space_elem = space.get("element")
    else:
        space_elem = space

    if level_name:
        for lvl in DB.FilteredElementCollector(doc).OfClass(DB.Level):
            if lvl.Name == level_name:
                return lvl

    level = getattr(space_elem, "Level", None)
    if level:
        return level
    try:
        if space_elem.LevelId and space_elem.LevelId != DB.ElementId.InvalidElementId:
            return doc.GetElement(space_elem.LevelId)
    except Exception:
        pass
    view = revit.active_view
    if hasattr(view, "GenLevel") and view.GenLevel:
        return view.GenLevel
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    return levels[0] if levels else None


def _parse_models(raw):
    if raw is None:
        return []
    text = str(raw).strip()
    if not text:
        return []
    return [p.strip() for p in text.split(",") if p.strip()]


def _norm_model_key(value):
    if value is None:
        return ""
    return re.sub(r"[^0-9A-Za-z]+", "", str(value)).upper()


def _model_keys(model):
    base = _norm_model_key(model)
    if not base:
        return []
    keys = [base]
    trimmed = re.sub(r"[A-Z]+$", "", base)
    if trimmed and trimmed != base:
        keys.append(trimmed)
    return keys


def _expand_models(models, count):
    if count <= 0:
        return []
    if not models:
        return [None] * count
    if count <= len(models):
        return models[:count]
    expanded = []
    idx = 0
    while len(expanded) < count:
        expanded.append(models[idx % len(models)])
        idx += 1
    return expanded


def _coerce_int(value):
    if value is None:
        return 0
    if isinstance(value, basestring):
        text = value.strip()
        if not text:
            return 0
        try:
            return int(float(text))
        except Exception:
            return 0
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return 0


def _row_value(row, keyset):
    if not row:
        return None
    norm = {}
    for key, val in row.items():
        if key is None:
            continue
        norm[_norm_key(key)] = val
    for key in keyset:
        if key in norm:
            return norm[key]
    return None


def _load_excel_rows(path):
    return _load_circuit_schedule_rows(path)


def _build_placements(rows, spaces, symbol_by_space, all_symbols):
    placements = []
    warnings = []
    match_info = []
    stats = {
        "skipped_non_mfr": 0,
        "skipped_missing_mfr": 0,
        "skipped_no_model_match": 0,
    }
    desc_order = []
    desc_seen = set()
    for row in rows:
        desc = _row_value(row, DESC_KEYS)
        count_val = _row_value(row, COUNT_KEYS)
        count = _coerce_int(count_val)
        if not desc:
            continue
        if count <= 0:
            continue
        if desc not in desc_seen:
            desc_seen.add(desc)
            desc_order.append(desc)
    mapping = _prompt_space_mapping(desc_order, spaces)
    if mapping is None:
        script.exit()
    for idx, row in enumerate(rows, start=2):
        row_idx = row.get("_row") if isinstance(row, dict) else None
        idx_label = row_idx if row_idx is not None else idx
        desc = _row_value(row, DESC_KEYS)
        count = _coerce_int(_row_value(row, COUNT_KEYS))
        models_raw = _row_value(row, MODEL_KEYS)
        manufacturer = _row_value(row, MFR_KEYS)
        if manufacturer:
            if REQUIRED_MANUFACTURER.lower() not in _norm_text(manufacturer):
                stats["skipped_non_mfr"] += 1
                continue
        else:
            stats["skipped_missing_mfr"] += 1
            continue

        if count <= 0:
            continue
        if not desc:
            warnings.append("Row {}: missing description.".format(idx_label))
            continue

        space = mapping.get(desc)
        if space is None:
            warnings.append("Row {}: no space selected for '{}'.".format(idx_label, desc))
            continue
        score = 0.0
        for key in _space_keys(space):
            score = max(score, _text_similarity(desc, key))
        match_info.append({
            "row": idx_label,
            "desc": desc,
            "space": (space.get("name") or space.get("number")) if space else "",
            "score": score,
            "passed": True,
            "source": space.get("source") or "Host",
        })

        models = _expand_models(_parse_models(models_raw), count)
        base_point = space.get("point")
        if not base_point:
            warnings.append(
                "Row {}: no placement point for space '{}'".format(
                    idx_label, space.get("name") or space.get("number")
                )
            )
            continue

        space_key = space.get("key") or space["element"].Id.IntegerValue
        symbols = list(symbol_by_space.get(space_key, [])) or list(all_symbols)
        if not symbols:
            warnings.append(
                "Row {}: no mechanical equipment types available for '{}'".format(
                    idx_label, space.get("name") or space.get("number")
                )
            )
            continue

        level = _resolve_level(space)
        for offset_idx, model in enumerate(models):
            symbol, sym_score = _best_symbol_match(model, symbols)
            if not symbol or sym_score < MODEL_MATCH_THRESHOLD:
                stats["skipped_no_model_match"] += 1
                warnings.append(
                    "Row {}: no close family type match for model '{}' in '{}' (score {:.2f}).".format(
                        idx_label, model or "", space.get("name") or space.get("number"), sym_score
                    )
                )
                continue
            fam_name = ""
            typ_name = ""
            try:
                fam_name = query.get_name(symbol.Family)
            except Exception:
                fam_name = ""
            try:
                typ_name = query.get_name(symbol)
            except Exception:
                typ_name = ""
            placements.append({
                "symbol": symbol,
                "point": base_point,
                "level": level,
                "offset": offset_idx * VERTICAL_OFFSET_FT,
                "space": space.get("name") or space.get("number"),
                "model": model,
                "manufacturer": manufacturer,
                "family": fam_name,
                "type": typ_name,
                "score": sym_score,
                "desc": desc,
                "space_score": score,
            })
    return placements, warnings, stats, match_info


def _place_instances(placements):
    placed_ids = []
    failures = []
    with revit.Transaction("Place all Coils"):
        for item in placements:
            symbol = item["symbol"]
            if not symbol:
                continue
            try:
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()
            except Exception:
                pass

            inst = None
            try:
                if item["level"] is not None:
                    inst = doc.Create.NewFamilyInstance(
                        item["point"],
                        symbol,
                        item["level"],
                        DB.Structure.StructuralType.NonStructural,
                    )
                else:
                    inst = doc.Create.NewFamilyInstance(
                        item["point"],
                        symbol,
                        DB.Structure.StructuralType.NonStructural,
                    )
            except Exception as ex:
                failures.append("Failed to place {}: {}".format(item.get("model") or symbol.Name, ex))
                continue

            if inst is None:
                failures.append("Failed to place {} (unknown error).".format(item.get("model") or symbol.Name))
                continue

            try:
                if item["offset"]:
                    DB.ElementTransformUtils.MoveElement(
                        doc,
                        inst.Id,
                        DB.XYZ(0, float(item["offset"]), 0),
                    )
            except Exception as ex:
                failures.append("Failed to offset {}: {}".format(inst.Id.IntegerValue, ex))

            placed_ids.append(inst.Id)

    return placed_ids, failures


def main():
    path = forms.pick_file(file_ext="xlsx", title="Select Coil Placement Excel File")
    if not path:
        return

    try:
        rows = _load_excel_rows(path)
    except Exception as ex:
        forms.alert(
            "Failed to read '{}' sheet: {}".format(SHEET_NAME, ex),
            exitscript=True,
        )
    if not rows:
        forms.alert("No readable rows found on '{}'.".format(SHEET_NAME), exitscript=True)

    spaces = _collect_host_spaces()
    if not spaces:
        forms.alert("No Spaces or Rooms found in this model.", exitscript=True)

    symbol_by_space = _collect_symbols_by_space(spaces)
    all_symbols = _collect_mech_symbols()
    if not all_symbols:
        forms.alert("No Mechanical Equipment family types found in this model.", exitscript=True)

    placements, warnings, stats, match_info = _build_placements(rows, spaces, symbol_by_space, all_symbols)
    if not placements:
        forms.alert("No valid coil placements were generated.", exitscript=True)

    placed_ids, failures = _place_instances(placements)

    if placed_ids:
        revit.get_selection().set_to(placed_ids)

    output = script.get_output()
    output.close_others()
    output.print_md("### Place all Coils")
    output.print_md("Placed {} coil(s).".format(len(placed_ids)))

    if stats.get("skipped_non_mfr") or stats.get("skipped_missing_mfr"):
        output.print_md(
            "Skipped {} row(s) (manufacturer not {}), {} row(s) (missing manufacturer).".format(
                stats.get("skipped_non_mfr", 0),
                REQUIRED_MANUFACTURER,
                stats.get("skipped_missing_mfr", 0),
            )
        )

    space_map = {}
    for p in placements:
        space_name = p.get("space") or "Unknown"
        fam_label = "{} : {}".format(p.get("family") or "", p.get("type") or "").strip()
        if fam_label == ":" or fam_label == "":
            fam_label = "Unknown Family"
        space_map.setdefault(space_name, {})
        space_map[space_name][fam_label] = space_map[space_name].get(fam_label, 0) + 1

    if space_map:
        output.print_md("#### Spaces and Coils Placed")
        for space_name in sorted(space_map.keys()):
            fam_counts = space_map[space_name]
            fam_list = ["{} x{}".format(name, fam_counts[name]) for name in sorted(fam_counts.keys())]
            output.print_md("- {}: {}".format(space_name, ", ".join(fam_list)))

    if match_info:
        output.print_md("#### Description → Space Matches")
        for entry in sorted(match_info, key=lambda e: e.get("row", 0)):
            desc = entry.get("desc") or ""
            space_name = entry.get("space") or "(no match)"
            score = entry.get("score", 0.0)
            status = "OK" if entry.get("passed") else "LOW"
            source = entry.get("source") or "Host"
            output.print_md(
                "- Row {}: '{}' → '{}' [{}] (score {:.2f}, {})".format(
                    entry.get("row"),
                    desc,
                    space_name,
                    source,
                    score,
                    status,
                )
            )

    if warnings:
        output.print_md("#### Warnings")
        for message in warnings:
            output.print_md("- {}".format(message))

    if failures:
        output.print_md("#### Placement Failures")
        for message in failures:
            output.print_md("- {}".format(message))


if __name__ == "__main__":
    main()
