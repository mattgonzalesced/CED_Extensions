# -*- coding: utf-8 -*-
from __future__ import absolute_import
from Autodesk.Revit.DB import Transaction, SubTransaction

class TransactionGroup(object):
    def __init__(self, doc, name="MEPKit"):
        self.doc = doc; self.name = name; self.tx = None
    def __enter__(self):
        self.tx = Transaction(self.doc, self.name); self.tx.Start(); return self
    def __exit__(self, et, ev, tb):
        if not self.tx: return
        (self.tx.Commit if et is None else self.tx.RollBack)()

class SubTx(object):
    def __init__(self, doc): self.doc = doc; self.sub = None
    def __enter__(self): self.sub = SubTransaction(self.doc); self.sub.Start(); return self
    def __exit__(self, et, ev, tb):
        if not self.sub: return
        (self.sub.Commit if et is None else self.sub.RollBack)()

def run_as_single_transaction(doc, name, fn, *args, **kwargs):
    """Run fn inside a single transaction and return its result."""
    with TransactionGroup(doc, name): return fn(*args, **kwargs)

def RunInTransaction(name=None):
    """Decorator to run (doc, *args) functions inside a single transaction."""
    def deco(fn):
        def wrapper(doc, *a, **k):
            label = name or ("MEPKit::" + fn.__name__)
            with TransactionGroup(doc, label):
                return fn(doc, *a, **k)
        return wrapper
    return deco