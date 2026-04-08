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
- **Refresh**: reload current circuit state from model.
- **List/Card toggle**: switch display style.
- **Options (...)**: theme, accent, and display preferences.

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
