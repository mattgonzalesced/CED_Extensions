# Quality Check & Analysis — Coding Plan

**Purpose:** This document plans and documents how quality checking and analysis for Revit should be coded. It does not contain implementation code; it specifies architecture, contracts, data flow, and implementation guidance for developers.

**Scope:** Mechanical, Electrical, Plumbing, Refrigeration, and Fire Protection design criteria in Revit.

**Convention:** New folders and documents for this planning/agent scope live under `CEDLib.lib/AI_Agents`.

---

## 1. Goals

- Define a **consistent way** to implement individual quality checks (e.g. proximity, clearance, parameter rules).
- Allow checks to be **run manually** (pyRevit buttons), **on sync** (optional), or **in batch** without duplicating logic.
- Support **configurable thresholds** (e.g. distance in inches) and **filtering** (design options, phases, categories) per check.
- Produce **repeatable, reportable results** (pass/fail, list of violating elements, distances/values) for auditing.

---

## 2. Existing Reference Implementation

The existing **Proximity Check (Lights–Coils)** under `AE pyTools.extension` → QualityChecks.panel is the reference pattern:

- **Location:** `QualityChecks.panel/Proximity Check.pushbutton/proximity_lights_coils.py`
- **Behavior:** Collects lighting fixtures and CED-R-KRACK coil family instances; computes bounding-box minimum distance; reports pairs within a threshold (e.g. 18 inches).
- **Integration:** Uses `CEDLib.lib/ExtensibleStorage` for a user setting (enable/disable run-on-sync); uses pyRevit `script.get_output()` for reporting and `forms.alert()` for notifications.
- **Library path:** Scripts resolve `CEDLib.lib` and append to `sys.path` to import shared modules.

**Coding plan should align with:**

- Same pattern for **collect → evaluate → report**.
- Same use of **ExtensibleStorage** (or a single designated mechanism) for per-document or per-check settings.
- Same **output contract**: results as a list of structured hits (element IDs, labels, numeric values like distance) so the same results can be driven to UI, sync hook, or batch export.

---

## 3. Recommended Architecture (How It Should Be Coded)

### 3.1 Layering

- **Check definitions (rules):** One module or one class per logical check (e.g. “lights vs refrigerated coils,” “sprinkler vs obstructions”). Each check defines:
  - **Input:** Revit document, optional filters (design option, phase, category/family filters).
  - **Parameters:** Thresholds (e.g. minimum clearance in inches), family name prefixes, category lists—preferably from a small set of named constants or a single config structure, not magic numbers in the middle of logic.
  - **Output:** A list of “hit” objects (e.g. element A id, element B id, distance or other measured value, human-readable labels). No UI or Revit API types in this output; keep it data-only so it can be reused for UI, sync, or batch.

- **Runners/orchestration:** A thin layer that:
  - Takes a document and a list of check definitions (or check identifiers).
  - Runs each check in sequence (or in parallel later if needed).
  - Aggregates results (e.g. by check name, pass/fail, list of hits).
  - Does not contain rule logic; it only invokes checks and collects results.

- **Reporting/presentation:** A separate layer that:
  - Takes the aggregated results and presents them (e.g. pyRevit output panel, alert, export to file).
  - Handles “no issues” vs “N issues” messaging and formatting (tables, links to elements).
  - Does not compute hits; it only formats and displays.

- **Persistence/settings:** Use the existing ExtensibleStorage pattern (or one designated alternative) for:
  - Per-document “run this check on sync” toggles.
  - Optional: stored thresholds or family filters per check, if desired, so that the same code can run with different configs per project.

### 3.2 Contracts (Interfaces to Implement)

- **Check function contract:**  
  - Input: `doc`, optional `options` (e.g. design option id, phase, set of category ids).  
  - Output: list of hit dictionaries (or a small dataclass) with at least: check name, element A id, element B id (if applicable), measured value (e.g. distance in feet or inches), and optional labels for UI.  
  - No side effects (no transaction, no UI). This allows the same function to be called from a button, from a sync event, or from a batch script.

- **Runner contract:**  
  - Input: `doc`, list of check functions or check ids.  
  - Output: one structure per check (check id/name, pass/fail, list of hits).  
  - Fail = at least one hit; pass = zero hits (or define explicitly for checks that have no “pair” concept).

- **Reporting contract:**  
  - Input: aggregated results (list of check results).  
  - Output: formatted output (e.g. markdown table, CSV, or in-memory structure for pyRevit).  
  - Reporting should be pluggable (e.g. “report to pyRevit” vs “report to file”) without changing check or runner code.

### 3.3 File and Module Organization

- **Under `CEDLib.lib/AI_Agents`:**  
  - Keep planning and spec documents (like this file) here.  
  - Optionally, a subfolder for “quality check rule definitions” (e.g. `QualityCheck_Rules`) could live here if rules are shared across multiple entry points (e.g. different pyRevit buttons or sync hooks).  
  - If rule definitions stay next to the UI (e.g. under `AE pyTools.extension`), then `AI_Agents` holds only the **planning docs** and any **shared contracts** (e.g. a small module that defines the hit schema and runner interface in one place).

- **Under the extension (e.g. `AE pyTools.extension`):**  
  - Each pyRevit button or sync hook is a thin script that:  
    - Resolves `CEDLib.lib` and imports the check (and optionally runner/reporting) from the library or from a sibling module.  
    - Calls the check with `doc` and any options (e.g. from UI or from stored settings).  
    - Passes results to the reporting layer and shows UI (output panel, alert).

- **Naming:**  
  - Check modules: one name per check, e.g. `proximity_lights_coils`, `clearance_sprinkler_obstruction`, `parameter_rule_xyz`.  
  - Runner/reporting: names like `quality_check_runner`, `quality_check_reporting` (or equivalent in the chosen language/structure).

### 3.4 Data Flow

1. **Trigger:** User runs a button, or sync event fires (if setting enabled).
2. **Resolve doc:** Current document (or passed doc for batch); skip if family document.
3. **Resolve checks:** Either a fixed list for that button or a list from config (e.g. “run these check ids when sync runs”).
4. **Run checks:** For each check, call the check function with `doc` and options; collect hits.
5. **Aggregate:** Build result set: check name, pass/fail, hits.
6. **Report:** Send result set to the reporting layer; reporting layer writes to output and/or alert.
7. **Optional:** Persist run summary (e.g. last run time, counts) via ExtensibleStorage if needed for “last results” or dashboards.

---

## 4. Check Categories and Rule Types (What to Code)

These are the **categories of rules** that the codebase should support. Each concrete check should be implemented according to the contracts above.

### 4.1 Proximity / Clearance (Distance)

- **Lights vs refrigerated coils:** Minimum clearance in inches (existing pattern).  
  - Config: family prefix for coils (e.g. CED-R-KRACK), category for lights, threshold in inches.  
  - Output: light id, coil id, distance.

- **Lights vs sprinklers:** Minimum clearance so coverage isn’t blocked; configurable threshold.  
  - Output: light id, sprinkler id, distance.

- **Electrical vs water/refrigeration:** Panels, conduits, boxes vs piping/equipment; minimum clearance.  
  - Output: electrical element id, MEP element id, distance.

- **Refrigeration vs heat sources:** Coils/condensers vs unit heaters, duct, high-temp piping.  
  - Output: refrig element id, heat source id, distance.

- **Equipment access:** Required clear space in front (and optionally sides) of equipment vs walls/duct/pipe.  
  - Output: equipment id, obstructing element id, distance (and which “side”).

- **Fire protection vs obstructions:** Sprinklers vs beams, duct, lights—minimum clearance per standard.  
  - Output: sprinkler id, obstruction id, distance.

**Implementation guidance:** Reuse the same distance primitive (e.g. bounding-box minimum distance, or point-to-point when a single placement point is defined). Parameterize element collection by category and/or family name/prefix so one “proximity check” helper can be reused across multiple rule definitions.

### 4.2 Mechanical

- Duct clearances vs structure, other duct, pipe, cable tray.  
- Equipment (AHU, VAV, fan) placement: access, no clash.  
- Insulation on cold duct / chilled water / refrigeration (presence and thickness if parameters exist).  
- Drainage: traps, slope, no backflow (logic: slope direction, trap presence by family or type).

### 4.3 Electrical

- Panel/board not in wet areas; access clearance.  
- Conduit vs hot/cold pipe clearance.  
- Circuiting/load checks (if Revit parameters support it): load per panel, circuit count vs capacity.  
- Egress/exit lighting presence and obstruction.

### 4.4 Plumbing

- Pipe vs structure/electrical/duct clearance.  
- Traps and slopes on drain lines.  
- Separation of domestic vs waste/vent; separation from electrical.  
- Cold-water in unconditioned spaces (insulation or heat trace by parameter or zone).

### 4.5 Refrigeration

- Coils vs heat sources (see proximity).  
- Coils vs sprinklers.  
- Insulation on refrig lines/vessels.  
- Service clearances around racks/condensers/evaporators.  
- Condensate drain slope and termination.

### 4.6 Fire Protection

- Sprinkler vs obstructions (beams, duct, lights, cable tray).  
- Coverage/spacing vs hazard (if data available).  
- Clearance to ceiling/soffit.  
- Standpipes/hose cabinets: access, no blocking.  
- Penetrations through fire-rated assemblies (if modeled and tagged).

**Implementation guidance:** For each of the above, define one “check” that follows the contract: input doc + options, output list of hits. Prefer shared helpers (element collectors, distance, parameter readers) in a common module so that new checks are mostly “which elements A, which elements B, what threshold, what label.”

---

## 5. Configuration and Thresholds

- **Where to store:**  
  - Defaults in the check module (constants or a small config object).  
  - Overrides: optional per-document or per-project storage via ExtensibleStorage (or a single config file under the project/model), keyed by check name.

- **What to make configurable:**  
  - Thresholds (e.g. minimum clearance in inches).  
  - Family name prefixes or category sets (e.g. “refrigerated coil” families).  
  - Design option and phase filters.  
  - Enable/disable “run on sync” per check.

- **Coding rule:** No magic numbers in the middle of logic; use named constants or a single config structure read at the start of the check.

---

## 6. Reporting and Output

- **In-Revit:**  
  - Use pyRevit output panel: markdown header, table of hits with columns (e.g. Element A, Element B, Distance).  
  - Use `output.linkify(element_id)` so elements are clickable.  
  - Alert when there are hits (and optionally when there are none, if “show empty” is on).

- **Structured export (future):**  
  - Same hit list can be serialized to JSON/CSV (check name, element ids, distance, labels) for batch runs or external dashboards.  
  - Reporting layer should have one “format results” function that returns a structure (list of rows or dicts), and separate “print to pyRevit” / “write to file” implementations.

---

## 7. Testing and Maintainability

- **Unit-testable parts:**  
  - Distance and geometry helpers (given two bounding boxes or two points, return distance).  
  - Hit structure construction (given two elements and a distance, return the standard hit dict).  
  - Filtering logic (category, family prefix) in isolation where possible.  
  - Reporting: given a list of hits, “format results” returns the expected structure (no Revit dependency).

- **Integration:**  
  - Run a check on a test document with known elements and assert hit count and approximate distances.  
  - Keep Revit-dependent collection (FilteredElementCollector, get_BoundingBox) in a thin layer so that, if needed, collectors can be mocked or replaced for tests.

---

## 8. Summary for Implementers

- **Do:**  
  - One module/function per check; same input/output contract.  
  - Separate collection → evaluation → reporting; reuse shared helpers.  
  - Store settings (and optional thresholds) via ExtensibleStorage or a single config pattern.  
  - File new planning docs and agent-related specs under `CEDLib.lib/AI_Agents`.  
  - Use the existing Proximity Check (lights–coils) as the reference implementation pattern.

- **Do not:**  
  - Put UI or Revit API types inside the core “hit” result; keep results data-only.  
  - Scatter magic numbers for thresholds; use named config.  
  - Duplicate collection/distance logic across checks; centralize in a small shared module.

This document should be updated when new check categories or architectural decisions are added, so that it remains the single source of truth for how quality checking and analysis should be coded.

---

## 9. MCP code reviewer feedback and plan updates

The **Quality Check Planner** receives feedback from the **MCP code reviewer** and updates this coding plan (and related docs under `CEDLib.lib/AI_Agents`) to reflect that feedback.

**Process:**

1. **Receive feedback:** The MCP code reviewer provides review comments (e.g. on architecture, contracts, naming, testability, consistency with existing code, or missing cases).
2. **Interpret:** Determine which parts of the coding plan are affected (e.g. §3 Contracts, §4 Check categories, §5 Configuration).
3. **Adjust the plan:** Edit this document (and any other planning docs) to incorporate the reviewer’s recommendations. Preserve the “plan only, no code” rule; changes are to structure, contracts, guidance, and scope—not to implementation code.
4. **Log (optional):** Add an entry to `AI_Agents/Review_Log.md` with date, summary of reviewer feedback, and brief description of plan changes made.

**What to incorporate:**

- **Contracts and interfaces:** If the reviewer suggests different function signatures, output shapes, or layering, update §3 (Architecture and contracts) and any check descriptions that depend on them.
- **Rule coverage:** If the reviewer identifies missing checks, edge cases, or incorrect categorization, update §4 (Check categories and rule types).
- **Config and reporting:** If the reviewer suggests different configuration or reporting behavior, update §5 and §6.
- **Testing/maintainability:** If the reviewer recommends different test boundaries or mock strategies, update §7.
- **Reference implementation:** If the reviewer points out that existing code (e.g. Proximity Check) should be treated differently or has changed, update §2 and any references to it.

**What not to do:**

- Do not add implementation code to the plan; keep this document as specification and guidance only.
- Do not change implementation files in the repo as part of “plan updates”; the planner adjusts plans; implementers (or the reviewer workflow) change code.
