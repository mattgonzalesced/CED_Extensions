# -*- coding: utf-8 -*-
# helper_power_balance.py
# After your main script runs, open each selected .rfa, delete any "Power Balance"
# electrical connector(s), SAVE back to the same path, and close.

import clr, os
from System.Collections.Generic import List as DotNetList

# Revit API
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, ElementId, OpenOptions,
    ModelPathUtils, IFailuresPreprocessor, FailureSeverity, FailureProcessingResult
)

# Some Revit versions expose ConnectorElement here
try:
    from Autodesk.Revit.DB import ConnectorElement
except:
    ConnectorElement = None

# Revit UI hook
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialogResult

# ----------------- dialog silencer (avoid "Document Opened" popups) -----------------
def _attach_dialog_silencer():
    def _handler(sender, e):
        try:
            if e.GetType().Name == 'TaskDialogShowingEventArgs':
                msg = getattr(e, "Message", "") or ""
                did = getattr(e, "DialogId", "") or ""
                if ("document opened" in msg.lower()) or ("pyrevit" in msg.lower()) or ("pyrevit" in did.lower()):
                    try: e.OverrideResult(int(TaskDialogResult.Ok))
                    except: pass
        except: pass
    try:
        __revit__.DialogBoxShowing += _handler
    except: pass
    return _handler

def _detach_dialog_silencer(handler):
    try: __revit__.DialogBoxShowing -= handler
    except: pass

# ----------------- warnings preprocessor -----------------
class _SwallowWarnings(IFailuresPreprocessor):
    def PreprocessFailures(self, accessor):
        for msg in list(accessor.GetFailureMessages()):
            try:
                if msg.GetSeverity() == FailureSeverity.Warning:
                    accessor.DeleteWarning(msg)
            except: pass
        return FailureProcessingResult.Continue

def _tx(doc, name):
    t = Transaction(doc, name); t.Start()
    try:
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_SwallowWarnings())
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)
    except: pass
    return t

# ----------------- open/save helpers -----------------
def _silent_open_family(path):
    mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
    opts = OpenOptions(); opts.Audit = False
    return __revit__.Application.OpenDocumentFile(mp, opts)

def _save_and_close(doc):
    try:
        if doc.IsModified:
            doc.Save()
        doc.Close(False)
        return True, ""
    except Exception as e:
        try: doc.Close(False)
        except: pass
        return False, str(e)

# ----------------- find + delete "Power Balance" connectors -----------------
def _param_as_text(doc, p):
    try:
        s = p.AsString()
        if s: return s
    except: pass
    try:
        v = p.AsValueString()
        if v: return v
    except: pass
    try:
        eid = p.AsElementId()
        if eid and eid != ElementId.InvalidElementId:
            el = doc.GetElement(eid)
            if el and hasattr(el, "Name"):
                return el.Name
    except: pass
    return ""

def _is_power_balance_connector(doc, conn):
    # 1) Try a type/enum-ish property if present
    try:
        st = getattr(conn, "SystemType", None)
        if st is not None:
            if "power" in str(st).lower() and "balance" in str(st).lower():
                return True
    except: pass
    # 2) Fallback: check connector parameters for something like "System Type" => "Power Balance"
    try:
        for p in conn.Parameters:
            try:
                dn = p.Definition.Name.lower()
            except:
                continue
            if "system" in dn and "type" in dn:
                val = _param_as_text(doc, p).lower()
                if ("power" in val) and ("balance" in val):
                    return True
    except: pass
    return False

def _find_power_balance_connector_ids(doc):
    ids = []
    if ConnectorElement is None:
        return ids
    try:
        for c in FilteredElementCollector(doc).OfClass(ConnectorElement):
            try:
                if _is_power_balance_connector(doc, c):
                    ids.append(c.Id)
            except: pass
    except: pass
    return ids

# ----------------- public entrypoint -----------------
def run_helper_delete_power_balance_and_save(paths):
    """
    Open each .rfa in 'paths', delete Power Balance connector(s) if present,
    then SAVE back to same path and close.
    """
    if not paths:
        print("[HELPER] No family paths provided; skipping.")
        return

    sil = _attach_dialog_silencer()
    try:
        total_deleted = 0
        for path in paths:
            print("\n[HELPER] Opening:", path)
            doc = None
            try:
                doc = _silent_open_family(path)
                if not doc.IsFamilyDocument:
                    print("  [WARN] Not a family. Skipping.")
                    if doc: doc.Close(False)
                    continue

                ids = _find_power_balance_connector_ids(doc)
                if ids:
                    t = _tx(doc, "Delete Power Balance Connector(s)")
                    try:
                        lst = DotNetList[ElementId]()
                        for i in ids: lst.Add(i)
                        doc.Delete(lst)
                        t.Commit()
                        print("  - Deleted {} Power Balance connector(s)".format(len(ids)))
                        total_deleted += len(ids)
                    except Exception as e:
                        t.RollBack()
                        print("  [ERROR] Delete failed:", e)
                else:
                    print("  - None found; nothing to delete.")

                ok, err = _save_and_close(doc)
                if ok:
                    print("  - Saved to same path.")
                else:
                    print("  [ERROR] Save failed:", err)

            except Exception as e:
                try:
                    if doc: doc.Close(False)
                except: pass
                print("  [ERROR] {} => {}".format(os.path.basename(path), e))
        print("\n[HELPER] Total Power Balance connectors deleted:", total_deleted)
    finally:
        _detach_dialog_silencer(sil)