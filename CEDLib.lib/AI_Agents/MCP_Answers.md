## 1. Mechanical Checks (Section 4.2)

### 1.1 Duct clearances vs structure, other duct, pipe, cable tray

**M1.1.1 – Target categories and system types**

- **Duct**:
  - `OST_DuctCurves`
  - `OST_FlexDuctCurves`
  - Optionally fabrication ductwork via a config flag when present.
- **Structure**:
  - `OST_StructuralFraming`
  - `OST_StructuralColumns`
  - `OST_StructuralFoundation`
- **Pipe**:
  - `OST_PipeCurves`
  - `OST_FlexPipeCurves`
  - Optional fabrication pipework via a config flag.
- **Cable tray**:
  - `OST_CableTray`
  - `OST_CableTrayFitting`

**M1.1.2 – Clearance thresholds (defaults, inches)**

- Duct vs structure: **2"**.
- Duct vs duct: **1.5"**.
- Duct vs pipe: **1.5"**.
- Duct vs cable tray: **1.5"**.

These are global defaults; allow per-project overrides via configuration/ExtensibleStorage.

**M1.1.3 – Direction and dimensionality**

- Use **3D bounding-box minimum distance** as the primitive.
- For duct vs structure, treat any 3D distance below the threshold as a hit.

### 1.2 Equipment placement: access, no clash

**M1.2.1 – Equipment families**

- **Mechanical equipment requiring access**:
  - Category: `OST_MechanicalEquipment`.
  - Initial family/type name filters (case-insensitive contains): `AHU`, `RTU`, `VAV`, `FAN`, `PUMP`, `FCU`, `ERV`.
- **Electrical equipment requiring access**:
  - Category: `OST_ElectricalEquipment`.
  - Treat all panels/switchboards/transformers as in-scope; refine by name later if needed.

**M1.2.2 – Required access space**

- Mechanical front access: **36"**.
- Electrical front access: **42"**.
- First implementation may approximate access as a symmetric bubble (radius 36"/42") around equipment, with a directional box in front added later if facing can be reliably derived.

**M1.2.3 – Obstruction categories**

Count as obstructions:

- `OST_Walls`
- `OST_StructuralFraming`
- `OST_PipeCurves`, `OST_FlexPipeCurves`
- `OST_DuctCurves`, `OST_FlexDuctCurves`
- `OST_CableTray`, `OST_CableTrayFitting`
- `OST_MechanicalEquipment`
- `OST_ElectricalEquipment`
- `OST_SpecialityEquipment`

Doors and openings are allowed in the access zone and are not treated as obstructions initially.

### 1.3 Insulation on cold duct / chilled water / refrigeration

**M1.3.1 – Identification of “cold” systems**

- Prefer system type name/parameter:
  - Duct systems with names containing: `CHW`, `CHILLED WATER`, `COLD AIR`, `REFRIG`.
  - Pipe systems with names containing: `CHW`, `CHILLED WATER`, `REFRIG`, `SUCTION`, `LIQUID LINE`.
- Allow a shared parameter override (e.g. `CED_IsColdSystem = 1`) per project.

**M1.3.2 – Insulation parameters and minimums**

- For duct: use duct insulation thickness parameter or a shared length parameter such as `CED_InsulationThickness`.
- For pipe: use pipe insulation thickness or the same shared parameter.
- Minimum thickness defaults:
  - Chilled water duct/pipe: **1.0"**.
  - Refrigeration suction/liquid: **1.5"**.

Flag missing insulation or thickness below these values.

### 1.4 Drainage: traps, slope, backflow

**M1.4.1 – Identifying drain lines**

- System type name contains: `SANITARY`, `WASTE`, `CONDENSATE`, `DRAIN`.
- Optionally a shared parameter `CED_IsDrain = 1`.

**M1.4.2 – Traps**

- Traps are dedicated families (e.g. `OST_PlumbingFixtures`, `OST_PipeAccessory`) whose names contain `TRAP`, `P-TRAP`, `CONDENSATE TRAP`, or fixtures with integral traps (e.g. sinks, floor drains).
- A drain segment is served by a trap if a trap family is found within **5 ft** along the same system downstream, or the connected fixture family/type implies an integral trap.

**M1.4.3 – Slope**

- Use `RBS_PIPE_SLOPE_PARAM` where available.
- Minimum slopes:
  - Sanitary/waste/condensate: **1/4" per foot** (~0.0208 ft/ft).
- Only enforce minimum (no maximum) in the first implementation.

## 2. Electrical Checks (Section 4.3)

### 2.1 Panel/board not in wet areas; access clearance

**E2.1.1 – Wet area definition**

- Wet rooms/spaces are identified by:
  - Room/space name patterns containing (case-insensitive): `BATH`, `SHOWER`, `RESTROOM`, `TOILET`, `LAV`, `LOCKER`, `DISH`, `KITCHEN`, `WASH`, `JANITOR`, `MOP`.
  - If available, a shared parameter `CED_WetArea = 1` takes precedence over name pattern.

**E2.1.2 – Target panels/boards**

- Category: `OST_ElectricalEquipment`.
- Names containing: `PANEL`, `SWBD`, `SWITCHBOARD`, `MCC`, `TRANSFORMER`, `SWITCHGEAR`.

**E2.1.3 – Access clearance**

- Front access: **42"**.
- Obstructions:
  - `OST_Walls`, `OST_Doors`
  - `OST_DuctCurves`, `OST_FlexDuctCurves`
  - `OST_PipeCurves`, `OST_FlexPipeCurves`
  - `OST_CableTray`, `OST_CableTrayFitting`
  - `OST_MechanicalEquipment`, `OST_ElectricalEquipment`, `OST_SpecialityEquipment`, `OST_Casework`

### 2.2 Conduit vs hot/cold pipe clearance

**E2.2.1 – Identification**

- Conduit: `OST_Conduit`, `OST_ConduitFitting`.
- Hot pipes: pipe systems with names containing `HW`, `HOT WATER`, `STEAM`, `HEATING`, `GAS`.
- Cold pipes: pipe systems with names containing `CHW`, `CHILLED`, `REFRIG`, `COLD WATER`.

**E2.2.2 – Thresholds**

- Conduit vs hot systems: **6"**.
- Conduit vs cold/chilled/refrigeration: **3"**.

Use 3D bounding distance; no vertical/horizontal split is required initially.

### 2.3 Circuiting/load checks

**E2.3.1 – Data source**

- Use built-in electrical circuit and panel parameters for:
  - Circuit load (VA/A).
  - Panel rating (A) and number of poles/spaces.
- If present, shared parameters like `CED_PanelRating`, `CED_PanelSpaces` may be used as overrides.

**E2.3.2 – Rule definition**

- Panel is overloaded if:
  - Total connected load > **80%** of panel rating for continuous load, or
  - Total connected load > **100%** of panel rating.
- Report:
  - Per-panel summary (OK/Overloaded, % of rating).
  - Per-circuit hits where a circuit exceeds its breaker rating or contributes to overload beyond the 80% rule.

### 2.4 Egress/exit lighting presence and obstruction

**E2.4.1 – Egress path and fixtures**

- Egress spaces:
  - Rooms/spaces named: `CORRIDOR`, `EXIT`, `EGRESS`, `LOBBY`, `HALL` (case-insensitive), or
  - Shared parameter `CED_IsEgress = 1`.
- Egress/exit fixtures:
  - Category: `OST_LightingFixtures`.
  - Names containing: `EXIT`, `EGRESS`, `EMERGENCY`, `EM`, `ECU`.

**E2.4.2 – Obstruction rule**

- Obstructions: `OST_DuctCurves`, `OST_StructuralFraming`, `OST_CableTray`, `OST_GenericModel`, `OST_SpecialityEquipment`, `OST_Casework`.
- Rule: treat an obstruction as blocking if its bounding box intrudes into a 3D region approximating the light beam (simplified as distance < **12"** from the fixture center along its facing direction for initial implementation).

## 3. Plumbing Checks (Section 4.4)

### 3.1 Pipe vs structure/electrical/duct clearance

**P3.1.1 – Targets**

- Pipes: `OST_PipeCurves`, `OST_FlexPipeCurves` (excluding explicit fire protection systems).
- Clearance vs:
  - Structure: `OST_StructuralFraming`, `OST_StructuralColumns`.
  - Electrical: `OST_ElectricalEquipment`, `OST_Conduit`, `OST_CableTray`, `OST_CableTrayFitting`.
  - Duct: `OST_DuctCurves`, `OST_FlexDuctCurves`.

**P3.1.2 – Thresholds**

- Pipe vs structure: **2"**.
- Pipe vs duct: **1.5"**.
- Pipe vs electrical: **3"**.

### 3.2 Traps and slopes on drain lines

**P3.2.1 – Drain vs vent vs domestic**

- Drains: system names containing `SANITARY`, `WASTE`, `CONDENSATE`, `FLOOR DRAIN`.
- Vents: system names containing `VENT`.
- Domestic water: system names containing `DOMESTIC`, `CWS`, `CWH`, `HWS`, `HWR`.

**P3.2.2 – Trap presence and distance**

- Same trap logic as mechanical drains (M1.4.2).
- Require a trap within **5 ft** of the fixture or start of run; otherwise flag.

**P3.2.3 – Slope rules**

- Sanitary/waste: **1/4" per foot** minimum.
- Vents: **1/8" per foot** minimum.
- Storm (if used): **1/8" per foot** minimum.

### 3.3 Separation of domestic vs waste/vent vs electrical

**P3.3.1 – Identification**

- Domestic: system names containing `DOMESTIC`, `CWS`, `CWH`, `HWS`, `HWR`.
- Waste/vent: system names containing `SANITARY`, `WASTE`, `VENT`, `DRAIN`.
- Electrical: `OST_Conduit`, `OST_CableTray`, `OST_CableTrayFitting`, `OST_ElectricalEquipment`.

**P3.3.2 – Separation distances**

- Domestic vs waste/vent: **6"**.
- Domestic vs electrical: **12"**.
- Waste/vent vs electrical: **6"**.

### 3.4 Cold-water in unconditioned spaces

**P3.4.1 – Unconditioned areas**

- A room/space is unconditioned if:
  - Shared parameter `CED_Conditioned = 0`, or
  - Name contains `MEZZANINE`, `STORAGE`, `VESTIBULE`, `EXTERIOR`, `PARKING`, `DOCK`, `FREEZER`, `COOLER`.

**P3.4.2 – Mitigation**

- For domestic cold water in unconditioned spaces:
  - Require insulation ≥ **1.0"**, or
  - A shared parameter `CED_HeatTrace = 1`.

## 4. Refrigeration Checks (Section 4.5)

### 4.1 Coils vs heat sources & coils vs sprinklers

**R4.1.1 – Refrigeration coil identification**

- Refrigeration coils/condensers/evaporators:
  - Category: `OST_MechanicalEquipment` or `OST_SpecialityEquipment`.
  - Names containing: `CED-R-KRACK`, `KRACK`, `COIL`, `EVAPORATOR`, `CONDENSER`, `REFRIG`, `RACK`.

**R4.1.2 – Heat source identification**

- Heat sources include:
  - Unit heaters/hot coils: names containing `UNIT HEATER`, `UH`, `HW COIL`, `HEATING COIL`.
  - Hot water/steam pipes: systems with names containing `HW`, `HEATING`, `STEAM`.

**R4.1.3 – Thresholds**

- Coils vs heat sources: **24"**.
- Coils vs sprinklers: **18"**.

### 4.2 Insulation on refrigeration lines/vessels

**R4.2.1 – Targets and parameters**

- Refrigeration lines: pipe systems with names containing `REFRIG`, `SUCTION`, `LIQUID LINE`, `HOT GAS`.
- Vessels: mechanical/specialty equipment with names containing `RECEIVER`, `SEPARATOR`, `ACCUMULATOR`.
- Use same insulation parameters as in 1.3 with a minimum of **1.5"**.

### 4.3 Service clearances around racks/condensers/evaporators

**R4.3.1 – Equipment families**

- Same as R4.1.1.

**R4.3.2 – Clearance distances**

- Front: **36"**.
- Sides: **24"**.
- Rear: **24"**.

### 4.4 Condensate drain slope and termination

**R4.4.1 – Identification**

- Condensate drains:
  - System names containing `COND`, `CONDENSATE`, `CD`.
  - Pipes connected to cooling/refrigeration coils.

**R4.4.2 – Slope and termination rules**

- Slope: **1/8" per foot** minimum.
- Acceptable terminations:
  - Floor/hub drains and indirect connections with host/family names containing `FLOOR DRAIN`, `FD`, `HUB DRAIN`, `DRAIN`.

## 5. Fire Protection Checks (Section 4.6)

### 5.1 Sprinkler vs obstructions (confirmation/expansion)

**F5.1.1 – Obstruction set**

- Obstructions:
  - `OST_StructuralFraming`, `OST_StructuralColumns`
  - `OST_DuctCurves`, `OST_FlexDuctCurves`
  - `OST_CableTray`, `OST_CableTrayFitting`
  - `OST_LightingFixtures`
  - `OST_GenericModel`, `OST_SpecialityEquipment`

**F5.1.2 – Thresholds**

- Single default: **18"** from sprinkler deflector to obstruction within the protection zone.

### 5.2 Coverage/spacing vs hazard

**F5.2.1 – Hazard classification**

- Preferred: shared parameter `CED_HazardClass` on rooms/spaces.
- Fallback: mapping from room name patterns (e.g. `STORAGE`, `WAREHOUSE`, `OFFICE`) to hazard classes.

**F5.2.2 – Spacing rules**

- Internal table of hazard → max spacing/coverage, for example:
  - Light Hazard (LH): max **130 ft²/head**, **15 ft** max spacing.
  - Ordinary Hazard (OH1/OH2): max **100 ft²/head**, **15 ft** max spacing.

### 5.3 Clearance to ceiling/soffit

**F5.3.1 – Distances**

- Vertical distance from deflector to ceiling/soffit:
  - Minimum: **1"** below ceiling.
  - Maximum: **12"** below ceiling.

### 5.4 Standpipes/hose cabinets: access, no blocking

**F5.4.1 – Identification**

- Families in `OST_SpecialityEquipment` or `OST_PlumbingFixtures` with names containing: `STANDPIPE`, `HOSE CABINET`, `FHC`, `HOSE VALVE`.

**F5.4.2 – Access bubble**

- Front: **36"** clear.
- Sides: **18"** radius.
- Obstructions: walls, mechanical/electrical equipment, casework, duct, pipe, cable tray, structural framing.

### 5.5 Penetrations through fire-rated assemblies

**F5.5.1 – Rated assemblies and penetrations**

- Rated hosts: walls/floors/ceilings with `Fire Rating` or `CED_FireRating` set.
- Penetrations modeled as:
  - Dedicated penetration/seal families (`SLEEVE`, `SEAL`, `FIRESTOP` in name), or
  - Pipes/ducts/conduits crossing a rated host without a penetration family.

**F5.5.2 – Rule intent**

- Flag when:
  - A pipe/duct/conduit crosses a rated host with no associated penetration/seal family, or
  - Penetration lacks a required rating parameter.

## 6. General Configuration and Thresholds (Section 5)

**G6.1 – Per-check configuration**

- Enable per-document overrides (ExtensibleStorage) for:
  - All proximity/clearance thresholds.
  - Insulation thickness and slope requirements.
  - Panel/circuit loading thresholds.
- Global defaults are defined in code; overrides are optional.

**G6.2 – Run-on-sync vs manual-only**

- Recommended run-on-sync:
  - Sprinkler vs obstructions.
  - Lights vs coils and extended refrigeration proximity.
  - Duct/pipe/tray clearance checks.
  - Panel in wet area check.
- Manual-only (initially):
  - Full egress lighting coverage/obstruction.
  - Detailed panel/circuit loading analysis.
  - Fire-rated penetration audits.
  - Large-scope refrigeration service clearance and condensate termination checks.

