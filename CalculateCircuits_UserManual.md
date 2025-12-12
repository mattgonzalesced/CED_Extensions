# Calculate Circuits – User Manual

## Purpose and Scope
Calculate Circuits automates sizing and documentation for Revit electrical circuits. It supports automatic sizing, supervised manual overrides, voltage-drop checks, conduit fill, downstream write-back to equipment/fixtures, and structured alerting. This guide explains how to configure settings, run the tool, and interpret results across all supported circuit types.

## Modes of Operation
- **Automatic mode (default)**
  - Tool sizes conductors, neutrals, grounds, and conduit based on connected/demand loads, code tables, and project defaults.
  - Missing or invalid override fields silently fall back to defaults, with alerts when settings are exceeded.
  - Voltage drop and conduit fill are enforced; sizing escalates until targets are met.
- **Manual override mode**
  - User-entered sizes drive the calculation. The tool validates entries, issues alerts for risky inputs, and preserves overrides except when they are invalid.
  - Clearing tokens control whether wire or conduit should be ignored (see **Clearing tokens** below).
  - Neutral sizing follows **Neutral Behavior**; grounds respect override values but are validated against NEC tables.

## Clearing Tokens
- Enter a single hyphen (`-`) to intentionally clear a size:
  - **Hot size = `-`**: Treat as “conduit only” (no voltage-drop calc). Circuit type set to `CONDUIT ONLY`; wire size string shows `-`.
  - **Hot size = `-` and Conduit size = `-`**: Treat as empty; branch data for conduit/cable is cleared and circuit type is `N/A`.
  - **Conduit size = `-`**: Conduit is cleared; conduit type is cleared for manual mode. Wire calculations continue if hots are provided.
- Clearing tokens are preserved in parameters so subsequent runs respect the user intent.

## Circuit Types and Special Cases
- **BRANCH**: Standard branch circuits sized to branch voltage-drop target.
- **FEEDER**: Feeders to panels, switchboards, and transformers sized to feeder voltage-drop target and feeder VD method.
- **XFMR PRI / XFMR SEC**: Transformer primaries/secondaries. Secondary circuits size service grounds from the service ground table using hot size; primaries follow equipment-ground rules.
- **CONDUIT ONLY / N/A**: Generated via clearing tokens as described above.

## Settings (Project-Level)
Access via **Calculate Circuits Settings**. Defaults are italic/gray; user selections are normal weight. The inline help panel shows the following guidance:

- **Minimum Conduit Size**
  - Smallest conduit size proposed during automatic calculations (has no effect on manual user overrides).
  - Options: `Selected: 3/4"`, `Selected: 1/2"`.
- **Max Conduit Fill**
  - Maximum allowable conduit fill as a percentage. In automatic mode, the conduit will be upsized until this fill is not exceeded. In manual override mode, the tool will alert the user if this value is exceeded.
- **Neutral Behavior**
  - Determines how neutrals are sized when in manual override mode (in automatic mode, neutral size always matches the hot size).
  - Options: `[Match hot conductors]` neutral matches hots; `[Manual Neutral]` user specifies neutral independently in manual mode.
- **Max Branch Voltage Drop**
  - Target maximum voltage drop for branch circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.
- **Max Feeder Voltage Drop**
  - Target maximum voltage drop for feeder circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.
- **Feeder VD Method**
  - Which feeder load basis to use for voltage drop calculations and automatic sizing (only applies to feeder circuits that supply panels, switchboards, and transformers. Branch circuits are always based on connected load).
  - Options: `[80% of Breaker]`, `[100% of Breaker]`, `[Demand Load]`, `[Connected Load]`.
- **Write Results (Equipment / Fixtures & Devices)**
  - When enabled, calculated values write back to downstream elements. Disabling either option prompts to clear stored circuit data from that category.

## Voltage Drop Behavior
- Branch circuits always use connected load for VD calculations.
- Feeders use the selected **Feeder VD Method**:
  - 80% or 100% of breaker (unless demand is higher), Demand Load, or Connected Load.
- Conductor upsizing considers cmil growth; equipment grounds upsize proportionally to hot conductor cmil increases.
- Warning thresholds: Branch/feeder VD targets generate alerts when exceeded in manual mode.

## Grounding Rules
- **Equipment grounds**: Sized from the EGC table by breaker rating and material; overrides must meet or exceed minimums.
- **Service grounds (transformer secondary)**: Sized from the service ground table by final hot conductor size; overrides must meet or exceed minimums.
- **Isolated grounds**: When enabled, IG size always matches equipment ground in manual and automatic modes.

## Materials, Insulation, and Conduit Types
- Material and insulation inputs are case-insensitive; tool normalizes to uppercase.
- Invalid properties fall back to defaults with alerts; sizing still uses the validated properties (no silent CU fallback for AL entries).
- Conduit entries accept normalized forms (e.g., `4"`, `4"C`, `4C`, `4`).

## Alerts and Notices
- Alerts are grouped by category (Override, Calculation, Design, Calculation Failure) with severity coloring:
  - **None/Info**: General notes.
  - **Medium (orange)**: Lug limits, non-standard breakers, excessive fill/VD.
  - **High (red)**: Undersized wires/OCP, ground deficiencies.
  - **Critical (red)**: Calculation failures (e.g., cannot size conduit/wire).
- Output per circuit:
  - **Circuit Name**
    - (Category) Alert message
- Developer logging can be toggled separately from user-facing alerts.

## Manual Mode Tips
- Keep materials/insulation valid so ampacity/VD use the intended tables.
- Use clearing tokens intentionally; they stay stored for future runs.
- Lug limits and parallel-set limits are soft: the tool preserves overrides but issues design warnings.
- Neutral Behavior controls whether neutral follows hot size or a user-entered size in manual mode.

## Automatic Mode Tips
- Leave fields blank to leverage project defaults.
- Min conduit size and max fill govern conduit upsizing; voltage-drop targets govern conductor upsizing.
- Lug and feeder guidance warn but do not override unless inputs are invalid.

## Data Write-Back
- Settings let you enable/disable writing to electrical equipment and to fixtures/devices. Disabling triggers a confirmation and clears stored values for the selected categories using filtered collectors (main model only).

## Transformer Secondary Handling
- Circuit type labeled `XFMR SEC` and uses service-ground sizing from the service ground table based on final hot size (post-VD).
- Voltage drop uses feeder logic; EGC rules are replaced by service-ground rules.

## Troubleshooting
- Review grouped alerts at the end of a run for any circuit. Critical alerts indicate sizing could not complete.
- If values look wrong, confirm materials/insulation are valid and that clearing tokens were not left unintentionally.
- Ensure downstream write-back is enabled if you expect parameters on equipment/fixtures to update.

