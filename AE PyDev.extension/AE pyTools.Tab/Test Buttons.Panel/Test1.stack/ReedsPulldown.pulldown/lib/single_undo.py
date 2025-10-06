# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import TransactionGroup

def run_as_single_undo(doc, title, work_fn, *args, **kwargs):
    """
    Runs work_fn(*args, **kwargs) inside a TransactionGroup so the user sees
    a single Undo/Redo step. If work_fn raises, the whole group is rolled back.
    Returns whatever work_fn returns.
    """
    tg = TransactionGroup(doc, title)
    tg.Start()
    try:
        result = work_fn(*args, **kwargs)
        tg.Assimilate()            # merge child transactions into ONE undo item
        return result
    except Exception as ex:
        try: tg.RollBack()
        except: pass
        raise