# Calling pyRevit Buttons from MCP Without Transaction Errors

## Problem

When MCP runs `sessionmgr.execute_command("<button_id>")` from inside `execute_revit_code`, the button runs in the same context as the MCP executor. Revit then rejects **Starting a new transaction** because that context is not allowed to open transactions (e.g. already inside a transaction or read-only).

## Solution

**Defer the button call** so it runs when Revit is idle, in a context that can start transactions:

1. Use Revit’s **ExternalEvent** (here via **CustomizableEvent** from `CEDLib.lib`).
2. From the MCP-executed code, do **not** call `sessionmgr.execute_command(...)` directly.
3. Instead, create a `CustomizableEvent`, pass a callable that runs `sessionmgr.execute_command(...)`, and call **`raise_event(callable)`**. Return immediately.
4. Revit later runs the callable in the ExternalEvent handler, where no transaction is open, so the button’s transaction succeeds.

## Pattern (IronPython run via MCP)

```python
import sys
sys.path.insert(0, r'c:\CED_Extensions\CEDLib.lib')
from pyrevitmep.event import CustomizableEvent
from pyrevit.loader import sessionmgr

def run_button():
    sessionmgr.execute_command("aepytools-aepytools-orientation-rotate1-orientleft")

ce = CustomizableEvent()
ce.raise_event(run_button)
```

- **Button ID**: Get from Shift+Windows+Click on the pyRevit button (e.g. `aepytools-aepytools-orientation-rotate1-orientleft`).
- **Path**: Adjust `CEDLib.lib` path if the extension lives elsewhere; only needed if the MCP executor doesn’t already have it on `sys.path`.

## MCP server change

When the server wants to “run a pyRevit button”, it should execute code that uses this **raise_event( callable )** pattern instead of calling `sessionmgr.execute_command(...)` directly. The callable can be a thin wrapper that calls `sessionmgr.execute_command(button_id)` so the button runs in a transaction-safe context.
