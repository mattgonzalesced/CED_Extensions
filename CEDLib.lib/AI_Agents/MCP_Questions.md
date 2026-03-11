# MCP Questions for Quality Check Implementation

This document collects questions for the Revit MCP reviewer so that we can safely implement all checks described in §§4–4.6 of `QualityCheck_Coding_Plan.md`. Please answer or adjust as needed based on actual project standards and content.

---

## 1. Mechanical Checks (Section 4.2)

### 1.1 Duct clearances vs structure, other duct, pipe, cable tray

- **M1.1.1 – Target categories and system types**
  - Which Revit categories and/or system types should be considered for:
    - *Duct*: (e.g. `OST_DuctCurves`, `OST_FlexDuctCurves`, fabrication ductwork?)
    - *Structure*: (e.g. `OST_StructuralFraming`, `OST_StructuralColumns`, `OST_StructuralFoundation`, others?)
    - *Pipe*: (e.g. `OST_PipeCurves`, `OST_FlexPipeCurves`, fabrication pipework?)
    - *Cable tray*: (e.g. `OST_CableTray`, `OST_CableTrayFitting`?)

- **M1.1.2 – Clearance thresholds**
  - What default minimum clearance (inches) should we use for:
    - Duct vs structure:
    - Duct vs duct:
    - Duct vs pipe:
    - Duct vs cable tray:
  - Are there any per-discipline or per-system overrides (e.g. different values for supply vs return, or for different building types)?

- **M1.1.3 – Direction and dimensionality**
  - Should these clearances be pure 3D bounding-box distances (like current proximity checks), or only in plan (XY), or only vertical?
  - For duct vs structure, do we care about:
    - Any minimum 3D distance, or
    - Only if they overlap in plan and are within a vertical tolerance (e.g. 6 in, 12 in)?

### 1.2 Equipment placement: access, no clash

- **M1.2.1 – Equipment families**
  - Which categories/families/types should be considered “equipment requiring access”?
    - Mechanical: (e.g. AHUs, VAVs, fans, pumps – what families or naming patterns?)
    - Electrical: (e.g. panels, switchboards, transformers?)

- **M1.2.2 – Required access space**
  - What default access distance (inches) should be used in front of:
    - Mechanical equipment:
    - Electrical equipment:
  - Do we need separate side/back clearance rules, or is “minimum bubble around equipment” sufficient for now?

- **M1.2.3 – Obstruction categories**
  - Which categories count as obstructions blocking access (e.g. walls, duct, pipe, structure, cable tray, other equipment)?
  - Are doors and openings allowed to be within the access zone, or should they be treated as obstructions?

### 1.3 Insulation on cold duct / chilled water / refrigeration

- **M1.3.1 – Identification of “cold” systems**
  - How do we identify cold duct and chilled/refrigeration piping:
    - By system type (e.g. chilled water supply/return, refrigerant lines)?
    - By family/type names or parameters?
    - By workset or some other convention?

- **M1.3.2 – Insulation parameters**
  - Which parameters should we inspect for insulation presence/thickness:
    - For duct:
    - For pipe:
  - Is there a minimum required thickness per system type (e.g. 1 in for chilled water, 2 in for certain refrigeration lines)?

### 1.4 Drainage: traps, slope, backflow

- **M1.4.1 – Identifying drain lines**
  - How do we identify drain lines that need slope and traps:
    - By piping system type (e.g. sanitary, waste, vent, condensate)?
    - By family/type name patterns?

- **M1.4.2 – Traps**
  - How are traps modeled in your content:
    - Dedicated trap family?
    - Integrated in certain fixture types?
  - What rule should determine that a given drain segment is “served by” a trap (e.g. trap within X feet downstream, in same system)?

- **M1.4.3 – Slope**
  - Which parameter(s) hold the design slope for drain lines (e.g. `RBS_PIPE_SLOPE_PARAM` or others)?
  - What minimum and maximum slopes should we enforce per system type (e.g. 1/4" per foot for certain drains)?

---

## 2. Electrical Checks (Section 4.3)

### 2.1 Panel/board not in wet areas; access clearance

- **E2.1.1 – Wet area definition**
  - How do we detect “wet areas”:
    - By room/space parameters (e.g. room name contains BATH, SHOWER, etc.)?
    - By a room/space classification parameter?
  - Do you have a list of room name patterns or a shared parameter to drive this?

- **E2.1.2 – Target panels/boards**
  - Which categories/families count as “panels/boards” for this rule (e.g. Electrical Equipment with certain type/family names)?

- **E2.1.3 – Access clearance**
  - What minimum access clearance in front of panels/boards should be checked (default 36 inches? 42 inches? Other)?
  - Which categories count as obstructions for panel access (walls, duct, pipe, casework, other equipment, etc.)?

### 2.2 Conduit vs hot/cold pipe clearance

- **E2.2.1 – Identification**
  - How do we distinguish:
    - Conduit (category `OST_Conduit` + others?) vs
    - Hot and cold pipes (which system types or parameters mark “hot” vs “cold”?)

- **E2.2.2 – Thresholds**
  - What minimum clearance (inches) is required between:
    - Conduit vs hot water/steam lines
    - Conduit vs cold water/chilled water/refrigeration lines
  - Are thresholds different for vertical vs horizontal proximity?

### 2.3 Circuiting/load checks

- **E2.3.1 – Data source**
  - Which Revit parameters hold:
    - Circuit load in VA/A for each circuit
    - Panel capacity (total load, number of poles, permitted circuits)
  - Are you using built-in parameters only, or are there custom/shared parameters we must read?

- **E2.3.2 – Rule definition**
  - How should we define “overloaded”:
    - Load > panel rating by any amount?
    - Load > rating * certain factor (e.g. 80%)?
  - Should we report:
    - Per-panel summaries (ok/overloaded), or
    - Per-circuit issues (e.g. circuit exceeds breaker size)?

### 2.4 Egress/exit lighting presence and obstruction

- **E2.4.1 – Egress path and fixtures**
  - How do we identify:
    - Exit/egress pathways (rooms, corridors, or model lines?)
    - Exit/egress lighting fixtures (family/type name patterns, parameters, or a specific category subset)?

- **E2.4.2 – Obstruction rule**
  - What constitutes an obstruction for egress lighting:
    - Duct, structure, signage, casework, other elements?
  - Are we checking:
    - Simple “lights not blocked by objects within X inches” or
    - More detailed line-of-sight / coverage that requires geometry?

---

## 3. Plumbing Checks (Section 4.4)

### 3.1 Pipe vs structure/electrical/duct clearance

- **P3.1.1 – Targets**
  - Which pipe systems and categories are in scope for clearance vs:
    - Structure
    - Electrical (equipment, conduit)
    - Duct

- **P3.1.2 – Thresholds**
  - What default clearance (inches) is required for each pair (pipe vs structure, pipe vs duct, pipe vs electrical)?
  - Any exceptions (e.g. certain systems may be allowed closer)?

### 3.2 Traps and slopes on drain lines

- **P3.2.1 – Drain vs vent vs domestic**
  - Which system types or parameters should we treat as:
    - Drains requiring traps and slope
    - Vents
    - Domestic water (no trap requirement, but maybe slope rules?)

- **P3.2.2 – Trap presence and distance**
  - Same as M1.4.2, but specific to plumbing:
    - Which families are traps?
    - How far from the fixture or start of run can a trap be before we flag it?

- **P3.2.3 – Slope rules**
  - Same as M1.4.3, but per plumbing system category (sanitary, storm, vent, etc.):
    - Min/max slopes per system?

### 3.3 Separation of domestic vs waste/vent vs electrical

- **P3.3.1 – Identification**
  - How do we identify:
    - Domestic (hot/cold) water lines
    - Waste/vent lines
    - Electrical elements to separate from (conduit, cable tray, panels, etc.)

- **P3.3.2 – Separation distances**
  - What minimum separation (inches) is required between:
    - Domestic vs waste/vent
    - Domestic vs electrical
    - Waste/vent vs electrical

### 3.4 Cold-water in unconditioned spaces

- **P3.4.1 – Unconditioned areas**
  - How do we determine that a space/zone is “unconditioned”:
    - By room/space parameters (e.g. occupancy, HVAC zone, a “conditioned” flag)?
    - By level or area classification?

- **P3.4.2 – Mitigation**
  - Which mitigation do we check:
    - Insulation present (which parameter)?
    - Heat trace present (which parameter)?
  - Do we require a specific minimum insulation thickness or just presence?

---

## 4. Refrigeration Checks (Section 4.5)

### 4.1 Coils vs heat sources & coils vs sprinklers

- **R4.1.1 – Refrigeration coil identification**
  - Beyond the existing `CED-R-KRACK` family prefix, what other families/types should be treated as refrigeration coils/condensers/evaporators?

- **R4.1.2 – Heat source identification**
  - How do we identify “heat sources” for clearance checks:
    - Unit heaters, hot-water coils, steam pipes, ovens, etc. – which families, categories, system types, or parameters?

- **R4.1.3 – Thresholds**
  - What minimum clearances (inches) are required for:
    - Coils vs heat sources
    - Coils vs sprinklers

### 4.2 Insulation on refrigeration lines/vessels

- **R4.2.1 – Targets and parameters**
  - Which pipe/equipment categories and system types are “refrigeration lines/vessels”?
  - Which parameters store required insulation values, and what are the minimum thicknesses?

### 4.3 Service clearances around racks/condensers/evaporators

- **R4.3.1 – Equipment families**
  - Which families should be treated as refrigeration racks/condensers/evaporators?

- **R4.3.2 – Clearance distances**
  - What service clearances (inches) do you want enforced:
    - Front
    - Sides
    - Rear
  - Are these symmetric (bubble around the equipment) or directional (front only is different)?

### 4.4 Condensate drain slope and termination

- **R4.4.1 – Identification**
  - How do we identify condensate drains vs other pipe systems?

- **R4.4.2 – Slope and termination rules**
  - What slope should we enforce on condensate drains?
  - What termination conditions are acceptable (e.g. floor drain, hub drain, indirect connection, specific family types)?

---

## 5. Fire Protection Checks (Section 4.6)

> Note: “Sprinkler vs obstructions (beams, duct, lights, cable tray)” already has an initial implementation (`clearance_sprinkler_obstructions`). The questions below focus on confirming/expanding that and implementing the remaining bullets.

### 5.1 Sprinkler vs obstructions (confirmation/expansion)

- **F5.1.1 – Obstruction set**
  - Are there any additional categories we should treat as obstructions beyond beams, duct, lights, and cable tray (e.g. signage, structure, ceilings)?

- **F5.1.2 – Thresholds**
  - Are different thresholds required for different obstruction types or hazard categories, or is a single default distance acceptable?

### 5.2 Coverage/spacing vs hazard

- **F5.2.1 – Hazard classification**
  - How do we determine the hazard category for each area:
    - Room name patterns?
    - A room/space classification parameter?
    - A project-wide mapping file?

- **F5.2.2 – Spacing rules**
  - What spacing rules should we enforce per hazard category (max spacing between sprinklers, max coverage area per head)?
  - Are any of these already encoded in parameters we can read, or should we use fixed values?

### 5.3 Clearance to ceiling/soffit

- **F5.3.1 – Distances**
  - What vertical distances should we enforce between sprinkler deflector and:
    - Ceiling
    - Soffit or dropped elements above?

### 5.4 Standpipes/hose cabinets: access, no blocking

- **F5.4.1 – Identification**
  - Which families/categories represent standpipes and hose cabinets?

- **F5.4.2 – Access bubble**
  - What access distances (inches) should be checked in front of and around these elements?
  - Which categories count as access obstructions?

### 5.5 Penetrations through fire-rated assemblies

- **F5.5.1 – Rated assemblies and penetrations**
  - How do we identify fire-rated walls/floors/ceilings in your models (parameters, tags, family types)?
  - How are penetrations modeled:
    - Dedicated penetration families?
    - Pipes/ducts simply passing through, with ratings on the host?

- **F5.5.2 – Rule intent**
  - What specific condition should cause a hit:
    - Any penetration through a rated assembly that is not associated with an approved penetration family?
    - Certain systems only (e.g. large pipes/ducts, not small conduits)?

---

## 6. General Configuration and Thresholds (Section 5)

- **G6.1 – Per-check configuration**
  - For which of the above checks do you want:
    - Per-document overrides for thresholds (via ExtensibleStorage)?
    - Global defaults only?

- **G6.2 – Run-on-sync vs manual-only**
  - Which checks should be eligible for “run on sync” behavior (like the current Lights–Coils check)?

