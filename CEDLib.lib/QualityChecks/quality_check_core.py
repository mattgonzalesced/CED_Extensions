import math
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from pyrevit import script


Hit = Dict[str, Any]
CheckFunc = Callable[[Any, Optional[Dict[str, Any]]], Sequence[Hit]]


def _coerce_hits(value: Any) -> List[Hit]:
    if value is None:
        return []
    if isinstance(value, dict):
        return [value]
    try:
        return list(value)
    except TypeError:
        return []


def run_checks(
    doc: Any,
    checks: Iterable[Tuple[str, CheckFunc, Optional[Dict[str, Any]]]],
) -> List[Dict[str, Any]]:
    """Run multiple quality checks and aggregate results.

    Each item in ``checks`` is a tuple of:
        (check_name, check_function, options_dict_or_None)

    The check function must:
        - accept (doc, options) and
        - return an iterable of hit dictionaries.
    """
    results: List[Dict[str, Any]] = []
    for name, func, options in checks:
        if func is None:
            continue
        try:
            hits = _coerce_hits(func(doc, options))
        except TypeError:
            # Fallback for check functions that do not accept options yet.
            hits = _coerce_hits(func(doc))  # type: ignore[arg-type]
        result = {
            "check_name": name,
            "hits": hits,
            "pass": not bool(hits),
        }
        results.append(result)
    return results


def summarize_results(results: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    total_checks = len(results)
    failing = [r for r in results if not r.get("pass", False)]
    total_hits = sum(len(r.get("hits") or []) for r in results)
    return {
        "total_checks": total_checks,
        "failing_checks": len(failing),
        "total_hits": total_hits,
    }


def _format_distance_inches(value_ft: Optional[float]) -> str:
    if value_ft is None:
        return ""
    return "{:.2f}".format((value_ft or 0.0) * 12.0)


def report_proximity_hits(
    title: str,
    subtitle: str,
    hits: Sequence[Hit],
    columns: Optional[Sequence[str]] = None,
    show_empty: bool = False,
) -> None:
    """Standard pyRevit reporting for proximity-style checks.

    Expects each hit to contain:
        - "a_id", "a_label"
        - "b_id", "b_label"
        - "distance_ft" (float, feet)
    """
    output = script.get_output()
    output.set_width(1000)

    if not hits:
        if show_empty:
            output.print_md("# {}".format(title))
            if subtitle:
                output.print_md("### {}".format(subtitle))
            output.print_md("No issues found.")
        return

    output.print_md("# {}".format(title))
    if subtitle:
        output.print_md("### {}".format(subtitle))

    rows = []
    for hit in hits:
        rows.append(
            [
                output.linkify(hit.get("a_id")),
                hit.get("a_label") or "<unknown>",
                output.linkify(hit.get("b_id")),
                hit.get("b_label") or "<unknown>",
                _format_distance_inches(hit.get("distance_ft")),
            ]
        )

    if not columns:
        columns = ["Element A ID", "Element A", "Element B ID", "Element B", "Distance (in)"]

    output.print_table(rows, columns=list(columns))

