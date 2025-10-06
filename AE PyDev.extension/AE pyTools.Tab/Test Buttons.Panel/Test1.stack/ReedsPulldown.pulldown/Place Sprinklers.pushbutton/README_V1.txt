script.py: orchestrates one run (Extract → Propose → Validate → Apply → Report). No discipline logic here.

lib/: reusable plumbing. Think “framework” (I/O, geometry, common validators, reporting).

lib/tools/: adapters that call your existing calculators. They return proposals as data, not Revit writes.

lib/rules/: where constraints live. Each discipline has its own file so constraints are explicit, testable, and versionable.

data/: editable standards (Excel/JSON). Non-programmers can tweak targets/limits here.

ui/: optional forms to pick scope & options.

tests/: smoke/regression tests for rules (fast, model-agnostic where possible).


Example folder structure.
RPTab.tab/
  RPPanel.panel/
    AI Agent.pushbutton/
      bundle.yaml                 # pyRevit metadata (tooltip, author, etc.)
      script.py                   # entrypoint; orchestration only
      icon.png

      lib/                        # stable, reusable building blocks
        __init__.py
        extractors.py             # read BIM: rooms, ceilings, panels, circuits, etc.
        geometry.py               # vectors, grids, clearance buffers, bbox ops
        selectors.py              # scope helpers (by level, selection, filters)
        apply.py                  # write ops (place instances, set params, tags)
        validators.py             # shared validation hooks (no discipline logic)
        report.py                 # JSON log + human-readable run notes
        io_config.py              # load Excel/JSON/YAML standards
        optimize.py               # generic OR-Tools helpers (if you use a solver)

        rules/                    # **discipline-specific rule packs**
          __init__.py
          core.py                 # base Rule classes, check plumbing, error model
          lighting.py             # lighting-only constraints & scoring
          electrical.py           # power-only constraints & scoring
          mechanical.py           # (optional) HVAC constraints
          plumbing.py             # (optional) plumbing constraints
          versions/               # frozen, versioned bundles for compliance
            lighting_v1_2.py
            electrical_v1_0.py

        tools/                    # **thin adapters around your existing scripts**
          __init__.py
          lighting.py             # wraps your spacing/illuminance script; pure I/O
          panels.py               # wraps panel placement heuristics
          ducts.py                # wraps sizing calcs
          circuits.py             # wraps branch-circuit grouping/balancing

      data/                       # project/company standards (editable)
        standards.json            # entry point (points to per-discipline files)
        lighting_standards.xlsx   # IES targets, spacing, clearances
        electrical_standards.xlsx # NEC clearances, mounting heights, naming
        symbols_map.csv           # friendly name -> FamilySymbol

      ui/
        form.py                   # small pyRevit UI (scope, task, dry-run)

      tests/                      # quick regression tests (optional but helpful)
        test_lighting_rules.py
        test_electrical_rules.py