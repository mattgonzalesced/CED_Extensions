# Calculate Circuits – User Manual

## Purpose and Scope
Calculate Circuits automates sizing and documentation for Revit electrical circuits. It supports automatic sizing, supervised manual overrides, voltage-drop checks, conduit fill, downstream write-back to equipment/fixtures, and structured alerting. This guide explains how to configure settings, run the tool, and interpret results across all supported circuit types.

## Modes of Operation
- Toggle mode via **CKT_User Override_CED** on the circuit: set to **No/blank** for automatic sizing or **Yes** for manual overrides.
- **Automatic mode (default)**
  - Active when **CKT_User Override_CED = No** (or left blank). The tool sizes conductors, neutrals, grounds, and conduit based on connected/demand loads, code tables, and project defaults.
  - Missing or invalid property fields (material, insulation, temperature, conduit type) silently fall back to defaults with alerts. Numeric sizes left blank are auto-sized.
  - Voltage drop and conduit fill are enforced; sizing escalates until targets are met.
- **Manual override mode**
  - Active when **CKT_User Override_CED = Yes**. User-entered sizes drive the calculation; the tool validates entries, issues alerts for risky inputs, and preserves overrides except when they are invalid.
  - Clearing tokens control whether wire or conduit should be ignored (see **Clearing tokens** below) and are honored only in manual mode.
  - Neutral sizing follows **Neutral Behavior**; grounds respect override values but are validated against NEC tables.
  - **CKT_Length Makeup_CED** is always honored as the circuit length input for VD sizing in both modes.

## Clearing Tokens (Manual Mode Only)
- Enter a single hyphen (`-`) to intentionally clear a size. Tokens are preserved in parameters so future runs keep the intent.
  - **Hot size = `-`**: Treat as “conduit only” (no voltage-drop calc). Circuit type set to `CONDUIT ONLY`; wire size string shows `-`; conduit sizes still calculate unless also cleared.
  - **Hot size = `-` and Conduit size = `-`**: Treat as empty; branch data for conduit/cable is cleared and circuit type is `N/A`. Wire size string shows `-`; conduit size string shows `-`.
  - **Conduit size = `-`**: Conduit is cleared; conduit type is cleared for manual mode. Wire calculations continue if hots are provided; conduit size string shows `-`.
- In **automatic mode**, cleared size tokens are ignored and replaced by calculated values; property overrides (material, insulation, temperature, conduit type) still validate and cleared size tokens are overwritten.

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
- **Isolated Ground Behavior**
  - Controls how isolated grounds size when in manual override mode (automatic mode always matches the equipment ground).
  - Options: `[Match ground conductors]` isolated ground size mirrors equipment ground; `[Manual Isolated Ground]` user specifies isolated ground size independently in manual mode.
- **Max Branch Voltage Drop**
  - Target maximum voltage drop for branch circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.
- **Max Feeder Voltage Drop**
  - Target maximum voltage drop for feeder circuits. In automatic mode, calculated sizes will grow until this threshold is met. In manual override mode, the tool will alert the user if this threshold is exceeded.
- **Feeder VD Method**
  - Which feeder load basis to use for voltage drop calculations and automatic sizing (only applies to feeder circuits that supply panels, switchboards, and transformers. Branch circuits are always based on connected load).
  - Options: `[80% of Breaker]`, `[100% of Breaker]`, `[Demand Load]`, `[Connected Load]`. If demand exceeds the breaker percentage options, the higher demand governs.
- **Write Results (Equipment / Fixtures & Devices)**
  - When enabled, calculated values write back to downstream elements. Disabling either option prompts to clear stored circuit data from that category.
- **Wire Material Display**
  - Controls when the wire material suffix is shown in wire size strings.
  - Options: `[Show material for Aluminum only]` and `[Show material for Copper and Aluminum]`.
- **Wire String Separator**
  - Controls the separator used in wire size strings.
  - Options: `[Use "+" separators]` or `[Use "," separators]`.

## Voltage Drop Behavior
- Branch circuits always use connected load for VD calculations.
- Feeders use the selected **Feeder VD Method**:
  - `[80% of Breaker]` or `[100% of Breaker]`: uses that fraction of breaker rating unless the demand load is higher, in which case demand governs.
  - `[Demand Load]`: uses the estimated demand load.
  - `[Connected Load]`: uses connected VA with no demand factors.
- Conductor upsizing considers cmil growth; equipment grounds upsize proportionally to hot conductor cmil increases.
- Warning thresholds: Branch/feeder VD targets generate alerts when exceeded in manual mode.

## Length Makeup
- **CKT_Length Makeup_CED** applies in both automatic and manual modes to adjust the effective run length for voltage drop.
- If *(Revit circuit length – Length Makeup)* is below zero, the tool reverts to the native Revit circuit length.

## Neutral Rules
- **Activation**: **CKT_Include Neutral_CED** adds or removes neutrals in both modes.
- **Quantity logic**:
  - 1P circuits default to one neutral when the distribution system includes a line-to-neutral voltage.
  - Branch circuits omit neutrals automatically if the base equipment distribution system lacks an LN voltage.
  - Feeder circuits include neutrals automatically when the downstream equipment distribution system has an LN voltage; otherwise they are omitted.
- **Sizing logic**:
  - Automatic mode: neutrals match hot size by default.
  - Manual mode: follows **Neutral Behavior**—either match hot or use a user-entered neutral size. When neutrals are intentionally omitted (include neutral unchecked or no LN system), no override warnings are produced and the neutral size writes as `-`.

## Isolated Grounds
- **CKT_Include Isolated Ground_CED** enables or removes isolated grounds in both modes.
- IG size always mirrors the equipment ground size; outputs write to the isolated ground size parameter and wire size strings.

## Grounding Rules
- **Equipment grounds**: Sized from the EGC table by breaker rating and material; overrides must meet or exceed minimums.
- **Service grounds (transformer secondary)**: Sized from the service ground table by final hot conductor size; overrides must meet or exceed minimums.
- **Isolated grounds**: When enabled, IG size matches equipment ground in manual and automatic modes and is written to the isolated ground size parameter.

## Wire and Conduit Properties
- **Editable in both automatic and manual modes** for material, insulation, temperature, and conduit type. Defaults apply when left blank in auto mode; manual entries are validated.
- **Wire material, insulation, temperature**: Case-insensitive; normalized to uppercase (temperature must match table entries). Invalid values fall back to defaults with override alerts.
- **Conduit type**: Case-insensitive; normalized before lookup. Invalid values fall back to defaults with override alerts.

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

### Standard Alerts Reference
The following alerts are available during Calculate Circuits runs:

| Alert ID | Severity | Meaning | Tool action |
| --- | --- | --- | --- |
| Overrides.InvalidCircuitProperty | None | A user-specified property (wire material, temperature, insulation, conduit type) was invalid. | Resets the property to the configured default and continues. |
| Overrides.InvalidEquipmentGround | None | A user-specified equipment ground size was invalid. | Replaces the override with NEC 250.122 sizing. |
| Overrides.InvalidServiceGround | None | A user-specified service ground size was invalid. | Replaces the override with NEC 250.102(c) sizing. |
| Overrides.InvalidHotWire | None | A user-specified hot conductor size was invalid. | Reverts to the calculated hot size. |
| Overrides.InvalidConduit | None | A user-specified conduit size was invalid. | Reverts to the calculated conduit size. |
| Overrides.InvalidIsolatedGround | None | A user-specified isolated ground size was invalid. | Uses the equipment ground size instead. |
| Design.NonStandardOCPRating | Medium | The breaker rating is non-standard. | Uses the next standard breaker size for calculations. |
| Design.BreakerLugSizeLimitOverride | Medium | User override exceeds recommended lug size for the breaker. | Keeps the override but flags a design warning. |
| Design.BreakerLugQuantityLimitOverride | Medium | User override exceeds recommended parallel set limit for the breaker. | Keeps the override but flags a design warning. |
| Calculations.BreakerLugSizeLimit | Medium | Calculated hot size exceeds recommended lug size for the breaker. | Keeps the calculated value and flags a design warning. |
| Calculations.BreakerLugQuantityLimit | Medium | Calculated parallel set count exceeds recommended lug limit for the breaker. | Limits set count to the recommended maximum. |
| Design.ExcessiveConduitFill | Medium | User-specified conduit size exceeds the max fill target. | Keeps the override but flags a design warning. |
| Design.UndersizedWireEGC | High | User-specified equipment ground size is undersized per NEC 250.122. | Keeps the override but flags a high-severity warning. |
| Design.UndersizedWireServiceGround | High | User-specified service ground size is undersized per NEC 250.102. | Keeps the override but flags a high-severity warning. |
| Design.ExcessiveVoltDrop | Medium | User-specified wire fails the voltage drop check. | Keeps the override but flags a design warning. |
| Design.InsufficientAmpacity | High | User-specified wire fails ampacity check versus circuit load. | Keeps the override but flags a high-severity warning. |
| Design.InsufficientAmpacityBreaker | High | User-specified wire fails ampacity check versus breaker rating. | Keeps the override but flags a high-severity warning. |
| Design.UndersizedOCP | High | User-specified breaker rating is undersized relative to circuit load. | Keeps the override but flags a high-severity warning. |
| Calculations.WireSizingFailed | Critical | Automatic wire sizing failed. | Marks calculation failed and outputs a critical alert. |
| Calculations.ConduitSizingFailed | Critical | Automatic conduit sizing failed. | Marks calculation failed and outputs a critical alert. |

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

## Valid Inputs and Parameter Names
- **Manual override toggle**: `CKT_User Override_CED` (Yes/No).
- **Length adjustment**: `CKT_Length Makeup_CED` (decimal feet; negative values fall back to Revit length).
- **Neutral inclusion**: `CKT_Include Neutral_CED` (Yes/No); **Isolated ground**: `CKT_Include Isolated Ground_CED` (Yes/No).
- **Wire properties** (editable in auto/manual):
  - Material: `CU`, `AL` (case-insensitive).
  - Insulation: `THHN`, `THWN`, `XHHW`, `XHHW-2`, etc., as supported by the project tables.
  - Temperature rating: `60 C`, `75 C`, `90 C` per tables.
- **Conduit type** (editable in auto/manual): EMT, RMC, IMC, FMC, ENT, and other supported types from conduit tables.
- **Size overrides (manual mode)**: Hot, neutral (when manual), ground, and conduit size accept standard trade sizes (e.g., `1/2"`, `3/4"`, `1"`, `4"C`, `4`). Clearing token `-` applies only in manual mode.
- **Settings options** (project-wide):
  - Minimum conduit size: `Selected: 3/4"`, `Selected: 1/2"`.
  - Neutral behavior: `[Match hot conductors]`, `[Manual Neutral]`.
  - Feeder VD method: `[80% of Breaker]`, `[100% of Breaker]`, `[Demand Load]`, `[Connected Load]`.
  - Max conduit fill: `0.1–1.0` (percent inputs accepted, normalized to decimal in storage).
  - Max voltage drop targets: `0.001–1.0` (percent inputs accepted, normalized to decimal in storage).


## Data Write-Back
- Settings let you enable/disable writing to electrical equipment and to fixtures/devices. Disabling triggers a confirmation and clears stored values for the selected categories using filtered collectors (main model only).

## Transformer Secondary Handling
- Circuit type labeled `XFMR SEC` and uses service-ground sizing from the service ground table based on final hot size (post-VD).
- Voltage drop uses feeder logic; EGC rules are replaced by service-ground rules.

## Troubleshooting
- Review grouped alerts at the end of a run for any circuit. Critical alerts indicate sizing could not complete.
- If values look wrong, confirm materials/insulation are valid and that clearing tokens were not left unintentionally.
- Ensure downstream write-back is enabled if you expect parameters on equipment/fixtures to update.
