# Alerts Browser

## Intro
`Alerts Browser` is a modeless circuit-alert triage tool.  
It lets users inspect alert records by circuit, select related model elements, and recalculate one circuit at a time.

## Main Components

### Header
- Title and active document label
- Count of circuits currently carrying alert payload

### Left Pane: Circuit List
- Single-selection list of circuits with alerts
- Each item shows:
  - Panel/circuit text
  - Load name
  - Alert counts (total, active, hidden)

### Right Pane: Alert Tabs
- `Active` tab: currently active alerts
- `Hidden` tab: hidden/suppressed alert definitions
- Grid columns:
  - Severity
  - Group
  - Alert ID
  - Message

## Bottom Command Area

### Select in Model
- `Panel`: selects upstream panel element
- `Circuit`: selects the electrical system element
- `Device`: selects downstream connected elements
- Clear selection (`CED.Icon.Close`)

### Refresh
- Recalculates the currently selected circuit and refreshes snapshot data.
- Disabled if:
  - no circuit is selected
  - operation is already running
  - selected circuit is blocked by ownership constraints
- Disabled tooltip explains the blocking reason.

## Modeless Behavior
- Window opens modeless (`Show`) and can remain open while editing model.
- Re-running the command focuses existing window instance.
- Operations are routed through `ExternalEvent` for safe API context.

## Ownership and Blocking
- Worksharing ownership is evaluated per circuit.
- Recalc button state reflects block status.
- Tooltip shows owner information when available.

## Typical Usage
1. Open Alerts Browser and pick a circuit from the left list.
2. Inspect active alert details.
3. Use `Panel`, `Circuit`, or `Device` to locate source elements.
4. Edit model as needed.
5. Click `Refresh` to recalculate selected circuit and verify alerts clear.

## Notes
- Alert payload source parameter: `Circuit Data_CED`
- Theme and accent are loaded from shared config (`AE-pyTools-Theme`).
