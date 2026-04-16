# Circuit Browser (User Guide)

## What This Tool Does
Circuit Browser helps you find circuits fast, check the ones you want, run circuit actions, and recalculate without jumping around Revit dialogs.

## Quick Start
1. Open **Circuit Browser** from the Circuit Manager split button.
2. Use **Search** or **Filter** to narrow the list.
3. Check circuits you want to work on.
4. Run an **Action** or use **Calculate**.

## Top Bar Buttons
- **Search**: filters circuits by panel, circuit number, load name, and related text.
- **Filter**: open filter options (types + special filters).
- **Actions**: opens action tools (Neutral/IG, Auto Size, Mark New/Existing, etc.).
<<<<<<< HEAD
=======
- **BatchSwap**: opens the batch swap workflow for panel/circuit remap operations.
- **Alerts Browser**: opens alert-focused review for issues found in the current circuit set.
>>>>>>> main
- **Refresh**: reload current circuit state from model.
- **List/Card toggle**: switch display style.
- **Options (...)**: theme, accent, and display preferences.

<<<<<<< HEAD
=======
## Circuit Browser - New Features
- **BatchSwap launch** directly from Circuit Browser.
- **Alerts Browser launch** directly from Circuit Browser.
- **Alert badge integration** on rows so issues are visible before opening detail views.
- **Faster multi-circuit workflows** with checkbox + selection behaviors for review, calculate, and actions.

>>>>>>> main
## Circuit Row Badges and Icons
- **Type badge**: shows circuit type (Branch, Feeder, Space, Spare, etc.).
- **N**: neutral included.
- **IG**: isolated ground included.
- **Override icon**: user override is enabled.
- **Sync lock icon**: operation is blocked by ownership/worksharing.
- **Alert (!)**: this circuit has alerts; click to view details.

## Select In Model
Use these to highlight elements in Revit:
- **Panel**: upstream panel
- **Circuit**: circuit element
- **Device**: downstream connected devices
- **X**: clear Revit selection

<<<<<<< HEAD
=======
## Interacting With The Circuit List
### Checked vs Selected
- **Checked** circuits are your action scope for most batch tools (actions, calculate selected, special filters like checked).
- **Selected** circuits are your current UI highlight/focus and row context.
- A circuit can be selected but not checked, checked but not currently selected, or both.

### Ctrl/Shift Click Behavior
- **Click** selects one row (single selection focus).
- **Ctrl+Click** adds/removes individual rows from the current selection set.
- **Shift+Click** selects a contiguous range between the last focused row and the clicked row.
- Selection is for navigation/context; checking controls what gets processed.

### Right-Click Options
- Right-click a row to open context actions for that circuit or the current selection.
- Typical options include opening related model selection targets, quick circuit actions, and alert/detail access.
- Use right-click when you need row-specific operations without changing your current top-bar workflow.

>>>>>>> main
## Calculate
- **All**: calculate all circuits shown by current context.
- **Selected**: calculate only checked/selected circuits.
- **Gear**: opens calculate settings.

## Filters
- Circuit type filters let you show/hide types.
- Special filters (warnings, overrides, blocked, checked) help isolate problem sets quickly.
- **Reset Filters** restores the normal full view.

## Actions (Common Uses)
- **Add/Remove Neutral**
- **Add/Remove IG**
- **Auto Size Breaker/Frame**
- **Mark as New/Existing**

Most actions open a review window before apply so you can verify changes first.

## If Something Is Disabled
Usually this means one of these:
- no circuits are checked/selected
- the circuit or related elements are owned by another user
- current filter hides what you expect to see

Use **Refresh** after model edits or ownership changes.
<<<<<<< HEAD
=======

# BatchSwap (User Guide)

## What This Tool Does
BatchSwap lets you stage and apply many slot-level panel schedule changes in one workflow, instead of editing each circuit one-by-one.

## Typical Workflow
1. Open **BatchSwap** from Circuit Manager or from Circuit Browser.
2. Pick source/target panel schedule context and build your staged operations.
3. Review sequence/order and target conditions.
4. Apply the staged set and validate in Circuit Browser.

## What BatchSwap Can Do
- Move circuits **within the same panel** to different slots.
- Move circuits **to different panels** (cross-panel transfer).
- Move circuits to **specific target slots**.
- Add **SPARE** entries.
- Add **SPACE** entries.
- Remove **SPARE** entries.
- Remove **SPACE** entries.
- Apply mixed plans in one pass (circuit moves + SPARE/SPACE add/remove together).

## Good Practices
- Run on a filtered subset first, then expand scope.
- Resolve ownership/worksharing locks before apply.
- Recalculate after major reassignment changes.

## Built-In Protections
- BatchSwap filters to **compatible targets** and blocks invalid cross-panel moves when no compatible target panel exists.
- SPARE/SPACE/circuit placement is constrained by panel schedule and slot validity rules (including equipment capacity/available slot checks).
- Operations that are not compatible with the target panel distribution setup are prevented before apply.

## Panel Schedule Creation In BatchSwap
- If a panel does not yet have a panel schedule view, BatchSwap can create it directly from the interface.
- Use **Create Schedule** after selecting a compatible template, then continue staging in the same workflow.

# Alerts Browser (User Guide)

## What This Tool Does
Alerts Browser centralizes saved circuit alerts in a modeless window so you can triage issues, interact with the model, and clear alerts quickly.

## Typical Workflow
1. Open **Alerts Browser** from Circuit Manager or Circuit Browser.
2. Select a circuit and review **Active** vs **Hidden** alert types.
3. Use model selection actions (panel/circuit/devices) to locate related elements.
4. Make corrections in Revit while Alerts Browser stays open.
5. Click **Refresh** to recalculate the selected circuit and rebuild alert state.
6. Confirm the alert clears (or continue triage if still present).

## Hiding Alert Types
- Alerts can be hidden per circuit by selecting an alert type and using **Hide Type**.
- Hidden types are shown on the **Hidden** tab and can be restored with **Unhide Type**.
- Only mapped/approved alert definitions are hideable.

## Hideable Alert Types
- `Design.NonStandardOCPRating`
- `Design.BreakerLugSizeLimitOverride`
- `Design.BreakerLugQuantityLimitOverride`
- `Design.CircuitLoadsNull`
- `Calculations.BreakerLugSizeLimit`
- `Calculations.BreakerLugQuantityLimit`

## Working While Window Is Open
- Alerts Browser is modeless, so you can keep it open while you work in the model.
- Use the selection tools to highlight related elements:
- **Panel** (upstream equipment)
- **Circuit** (electrical system)
- **Device** (downstream connected elements)
- **Clear** (clear Revit selection)
- After edits, click **Refresh** to recalculate and update active/hidden alert results for the selected circuit.

## Working Efficiently
- Start with high-impact alerts first.
- Use Circuit Browser filters together with Alerts Browser to isolate related problem sets.
- Re-run calculation after corrective actions when alert types are load/sizing related.
>>>>>>> main
