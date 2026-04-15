# Circuit Browser

## Intro
`Circuit Browser` is a dockable pane for managing electrical circuits at scale.  
It combines filtering, batch actions, model selection tools, and calculation triggers in one UI.

## Main Components

### Top Toolbar
- Search box: text filter for panel, circuit number, load name, and related circuit row content.
- Filter button (`CED.Icon.Filter`): opens filter menu for circuit type toggles and special filters.
- Actions button (`CED.Icon.ChevronDown`): opens action menu for circuit operations.
- Refresh button (`CED.Icon.Refresh`): rebuilds the browser list from current model state.
- Display mode toggle (`CED.Icon.List` / `CED.Icon.Card`): switches compact list and card templates.
- Options button (`...`): opens theme, accent, and display-related options.

### Circuit List
- Supports checkboxes, row selection, virtualization, and context menu behavior.
- Works in compact and card templates with shared data bindings.
- Status and warning badges are computed from branch data and alert payload.

### Action Group
- Select in Model:
  - `Panel`
  - `Circuit`
  - `Device`
  - Clear selection (`CED.Icon.Close`)
- Calculate:
  - `All`
  - `Selected`
  - Settings (`CED.Icon.Cog`)

## Circuit Row Markers

### Type Badge
Category badge representing circuit type (`BRANCH`, `FEEDER`, `SPACE`, `SPARE`, `XFMR PRI`, `XFMR SEC`, `CONDUIT ONLY`, `N/A`).

### Neutral / IG Badges
- `N` and `IG` badges indicate neutral and isolated-ground inclusion state.

### User Override Badge
- Override badge icon indicates a user override condition is set.

### Sync Lock Badge
- Sync lock icon (`CED.Icon.SyncAlert`) indicates calculation/writeback is blocked by ownership constraints.
- Tooltip explains why and includes owner details when available.

### Alert Badge
- Alert button (`!`) indicates alert records exist for the circuit.
- Hover shows summary text.
- Click opens circuit alert details.

## Filter Menu

### Circuit Type Toggles
- Enable or disable each circuit type category.

### Exclusive Filters
- Warning-only
- User override-only
- Failed/blocked-only
- Checked-only
- These modes are mutually exclusive and automatically show all types while active.

### Reset Filters
- Restores full type visibility and clears special filter flags.

## Options Menu
- Theme mode: `Light`, `Dark`, `Dark Alt`
- Accent mode: `Blue`, `Red`, `Green`, `Neutral`
- Display mode options used by card/list presentation

## Actions Menu
- Add/Remove Neutral
- Add/Remove IG
- Auto Size Breaker / Frame
- Mark as New/Existing
- Alerts views and related workflows

Each action opens a dedicated window that uses shared UI styles and ownership-aware blocking.

## Ownership and Blocking Behavior
- Circuit and downstream element ownership is evaluated before writeback operations.
- If blocked, the UI marks affected rows and prevents unsafe operations.
- Tooltip/status messaging describes the exact ownership reason.

## Typical Usage
1. Search and filter to isolate circuits.
2. Check target rows.
3. Use `Select in Model` to verify affected elements.
4. Run an action or calculate selected circuits.
5. Review alert/sync markers and rerun as needed.

## Notes
- The browser is designed for large models and uses virtualization.
- Theme and accent preferences are shared via `AE-pyTools-Theme` config.
