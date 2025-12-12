# Calculate Circuits – User Manual

## Purpose and Scope
Calculate Circuits automates sizing and documentation for Revit electrical circuits. It supports automatic sizing, supervised manual overrides, voltage-drop checks, conduit fill, downstream write-back to equipment/fixtures, and structured alerting. This guide explains how to configure settings, run the tool, and interpret results across all supported circuit types.

## Modes of Operation
- **Automatic mode (default)**
  - Active when **CKT_User Override_CED = No** (or left blank). The tool sizes conductors, neutrals, grounds, and conduit based on connected/demand loads, code tables, and project defaults.
  - Missing or invalid property fields (material, insulation, temperature, conduit type) silently fall back to defaults with alerts. Numeric sizes left blank are auto-sized.
  - Voltage drop and conduit fill are enforced; sizing escalates until targets are met.
- **Manual override mode**
  - Active when **CKT_User Override_CED = Yes**. User-entered sizes drive the calculation; the tool validates entries, issues alerts for risky inputs, and preserves overrides except when they are invalid.
  - Clearing tokens control whether wire or conduit should be ignored (see **Clearing tokens** below) and are honored only in manual mode.
  - Neutral sizing follows **Neutral Behavior**; grounds respect override values but are validated against NEC tables.
  - **Length Makeup_CED** is always honored as the circuit length input for VD sizing in both modes.

## Clearing Tokens (Manual Mode Only)
- Enter a single hyphen (`-`) to intentionally clear a size. Tokens are preserved in parameters so future runs keep the intent.
  - **Hot size = `-`**: Treat as “conduit only” (no voltage-drop calc). Circuit type set to `CONDUIT ONLY`; wire size string shows `-`; conduit sizes still calculate unless also cleared.
  - **Hot size = `-` and Conduit size = `-`**: Treat as empty; branch data for conduit/cable is cleared and circuit type is `N/A`. Wire size string shows `-`; conduit size string shows `-`.
  - **Conduit size = `-`**: Conduit is cleared; conduit type is cleared for manual mode. Wire calculations continue if hots are provided; conduit size string shows `-`.
- In **automatic mode**, cleared size tokens are ignored and replaced by calculated values; property overrides (material, insulation, temperature, conduit type) still validate.

## Circuit Types and Special Cases
- **BRANCH**: Standard branch circuits sized to branch voltage-drop target.
- **FEEDER**: Feeders to panels, switchboards, and transformers sized to feeder voltage-drop target and feeder VD method.
- **XFMR PRI / XFMR SEC**: Transformer primaries/secondaries. Secondary circuits size service grounds from the service ground table using hot size; primaries follow equipment-ground rules.
- **CONDUIT ONLY / N/A**: Generated via clearing tokens as described above.
- **Circuit type inference** follows load classification and clearing tokens; strings update automatically based on results.

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
  - Options: `[80% of Breaker]`, `[100% of Breaker]`, `[Demand Load]`, `[Connected Load]`. If demand exceeds the breaker percentage options, the higher demand governs.
- **Write Results (Equipment / Fixtures & Devices)**
  - When enabled, calculated values write back to downstream elements. Disabling either option prompts to clear stored circuit data from that category.

## Voltage Drop Behavior
- Branch circuits always use connected load for VD calculations.
- Feeders use the selected **Feeder VD Method**:
  - `[80% of Breaker]` or `[100% of Breaker]`: uses that fraction of breaker rating unless the demand load is higher, in which case demand governs.
  - `[Demand Load]`: uses the estimated demand load.
  - `[Connected Load]`: uses connected VA with no demand factors.
- Conductor upsizing considers cmil growth; equipment grounds upsize proportionally to hot conductor cmil increases.
- Warning thresholds: Branch/feeder VD targets generate alerts when exceeded in manual mode.

## Grounding Rules
- **Equipment grounds**: Sized from the EGC table by breaker rating and material; overrides must meet or exceed minimums.
- **Service grounds (transformer secondary)**: Sized from the service ground table by final hot conductor size; overrides must meet or exceed minimums.
- **Isolated grounds**: When enabled, IG size always matches equipment ground in manual and automatic modes and is written to the isolated ground size parameter.
- **Neutrals**: In automatic mode, neutrals match hots. In manual mode, behavior follows the **Neutral Behavior** setting; no warnings appear when neutrals are intentionally omitted per system type.

## Wire and Conduit Properties
- **Editable in both automatic and manual modes.** Defaults apply when left blank in auto mode; manual entries are validated.
- **Wire material, insulation, temperature**: Case-insensitive; normalized to uppercase (temperature must match table entries). Invalid values fall back to defaults with override alerts.
- **Conduit type**: Case-insensitive; normalized before lookup. Invalid values fall back to defaults with override alerts.
- **Conduit size**: Accepts normalized forms (e.g., `4"`, `4"C`, `4C`, `4`).

## Alerts and Notices
- Alerts are grouped by category (Override, Calculation, Design, Calculation Failure) with severity coloring:
  - **None/Info**: General notes.
  - **Medium (orange)**: Lug limits, non-standard breakers, excessive fill/VD.
  - **High (red)**: Undersized wires/OCP, ground deficiencies.
  - **Critical (red)**: Calculation failures (e.g., cannot size conduit/wire).
- Output per circuit renders as:
  - **Circuit Name**
    - (Override) message
    - (Calculation) message
    - (Design) message
- Developer logging can be toggled separately from user-facing alerts.

## Manual Mode Tips
- Keep materials/insulation/temperature valid so ampacity/VD use the intended tables.
- Use clearing tokens intentionally; they stay stored for future runs and drive circuit type/conduit-only behavior.
- Lug limits and parallel-set limits are soft: the tool preserves overrides but issues design warnings.
- Neutral Behavior controls whether neutral follows hot size or a user-entered size in manual mode; neutral/IG parameters still write even when hots are cleared (where applicable).
- Length Makeup_CED drives voltage-drop distance; confirm it matches field conditions.

## Automatic Mode Tips
- Leave size fields blank to leverage project defaults; property fields (material, insulation, temperature, conduit type) still validate.
- Min conduit size and max fill govern conduit upsizing; voltage-drop targets govern conductor upsizing.
- Lug and feeder guidance warn but do not override unless inputs are invalid.
- Clearing tokens are ignored in automatic mode—calculated sizes overwrite them.

## Wire Size Strings
- Wire size strings are read-only outputs summarizing the final configuration (sets × size × material × insulation × temperature plus neutrals/grounds/IGs).
- When neutrals or IGs differ from hots, the strings reflect the unique sizing. Clearing tokens replace the associated segment with `-`.


## Data Write-Back
- Settings let you enable/disable writing to electrical equipment and to fixtures/devices. Disabling triggers a confirmation and clears stored values for the selected categories using filtered collectors (main model only).

## Transformer Secondary Handling
- Circuit type labeled `XFMR SEC` and uses service-ground sizing from the service ground table based on final hot size (post-VD).
- Voltage drop uses feeder logic; EGC rules are replaced by service-ground rules.

## Troubleshooting
- Review grouped alerts at the end of a run for any circuit. Critical alerts indicate sizing could not complete.
- If values look wrong, confirm materials/insulation are valid and that clearing tokens were not left unintentionally.
- Ensure downstream write-back is enabled if you expect parameters on equipment/fixtures to update.

