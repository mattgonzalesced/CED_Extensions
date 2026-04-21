# Add / Remove Spares and Spaces Tool Tutorial (Draft)

## 1. What the Tool Does
Add / Remove Spares and Spaces is a panel schedule utility for bulk managing default spares and spaces.
It supports two workflows:
- Quick UI for fast context-based actions.
- Full staged planner window for multi-panel review and controlled apply.

## 2. Quick UI vs Normal Window
The tool decides which UI to open at launch.

### Quick UI opens when:
1. The active view is a `PanelScheduleView` that resolves to a panel.
2. Or, if not in a panel schedule view, the current selection contains one or more `PanelScheduleSheetInstance` elements that resolve to panels.

Quick-mode precedence:
- If an active panel schedule view is found, Quick UI targets that active panel and opens immediately.
- Sheet-instance selection is only used when the active view is not a panel schedule view.

### Normal window opens when:
- No quick-mode panel context is found from active view or selected sheet schedule instances.

## 3. Quick UI Features
- Displays panel context:
  - Single panel: panel/type/distribution/open slot info.
  - Multiple panels from sheet selection: count summary.
- Action type toggle:
  - Fill mode: add defaults.
  - Remove mode: remove defaults.
- Mode buttons:
  - Spare
  - Space
  - Both
- Executes immediately after button click (no staging grid).
- Shows completion summary with counts:
  - Added/removed spare count
  - Added/removed space count
  - Panel count touched
- Shows warnings when applicable:
  - Unlock warnings for newly added rows
  - Switchboard pole-setting warnings

## 4. Normal Window Features (Staged Planner)
- Loads all panel schedules (non-template schedule-backed options).
- Shows each panel row with:
  - Panel identity
  - Distribution system
  - Open slots
  - Staged action text
- Selection controls:
  - Row checkboxes
  - Multi-row grid selection
  - Check All / Uncheck All
- Mode controls:
  - Spare
  - Space
  - Both
- Staging controls:
  - Stage Add
  - Stage Remove
  - Reset Selected (clears staged actions for selected rows)
- Apply runs all staged actions in one transactional pass.
- If apply fails, the transaction group is rolled back.
- After apply, open-slot counts are refreshed and staged actions are cleared.

## 5. Action Rules and Protections
- Add actions only use available open slots.
- In Add + Both mode, slots are split as:
  - First half spare
  - Remaining half space
- Remove actions are filtered to removable defaults only:
  - Spares must match removable-default criteria.
  - Spaces must match removable-space criteria.
- Newly added defaults are finalized after add:
  - Attempts to unlock added cells/slots.
  - Attempts to set switchboard added circuits to 3-pole.
  - Reports any failures as warnings.

## 6. Suggested Training Walkthrough
1. Start in a panel schedule view and demonstrate Quick UI.
2. Run Quick Fill for Spare, Space, and Both.
3. Run Quick Remove and explain default-only removals.
4. Exit panel schedule context and open full staged planner.
5. Demonstrate check/select behavior and Stage Add/Remove.
6. Demonstrate Reset Selected.
7. Apply staged actions and review status/warnings.
