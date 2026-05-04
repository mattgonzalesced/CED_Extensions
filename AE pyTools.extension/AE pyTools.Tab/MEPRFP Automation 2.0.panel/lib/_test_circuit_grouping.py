# -*- coding: utf-8 -*-
"""Tests for circuit_grouping + circuit_phasing."""

from __future__ import print_function

from circuit_grouping import (
    CircuitItem,
    assemble_groups,
    BUCKET_NORMAL,
    BUCKET_DEDICATED,
    BUCKET_BYPARENT,
    BUCKET_SECONDBYPARENT,
    NEEDS_REVIEW_PANEL,
    NEEDS_REVIEW_CIRCUIT,
)
from circuit_phasing import (
    PanelPhaseTracker,
    select_distribution_system_id,
    PHASE_A, PHASE_B, PHASE_C, PHASE_AB, PHASE_BC, PHASE_CA, PHASE_ABC,
)


def _check(label, cond):
    print("  {}  {}".format("PASS" if cond else "FAIL", label))
    return bool(cond)


def _make(elem_id, panel="L1", ckt="12", load="LOAD",
          bucket=BUCKET_NORMAL, poles=1, parent=None, world_pt=None):
    return CircuitItem(
        element_id=elem_id,
        panel_name=panel,
        circuit_token=ckt,
        load_name=load,
        bucket=bucket,
        poles=poles,
        parent_element_id=parent,
        world_pt=world_pt,
    )


def test_normal_grouping():
    print("\n[grouping] normal: same panel + circuit number bucket together")
    items = [
        _make(1, "L1", "12"),
        _make(2, "L1", "12"),
        _make(3, "L1", "14"),
        _make(4, "L2", "12"),
    ]
    groups = assemble_groups(items)
    fails = []
    if not _check("3 groups for normal items", len(groups) == 3):
        fails.append("count")
    g_l1_12 = next(g for g in groups
                   if g.panel_name == "L1" and g.circuit_token == "12")
    if not _check("L1/12 has both 1 and 2", g_l1_12.member_count == 2):
        fails.append("merge")
    return fails


def test_dedicated_one_per_item():
    print("\n[grouping] dedicated: one circuit per item even with same panel")
    items = [
        _make(11, "L1", "DEDICATED", bucket=BUCKET_DEDICATED),
        _make(12, "L1", "DEDICATED", bucket=BUCKET_DEDICATED),
        _make(13, "L1", "DEDICATED", bucket=BUCKET_DEDICATED),
    ]
    groups = assemble_groups(items)
    fails = []
    if not _check("3 dedicated groups", len(groups) == 3):
        fails.append("count")
    if not _check("each holds 1 member",
                  all(g.member_count == 1 for g in groups)):
        fails.append("isolation")
    return fails


def test_byparent_groups_by_parent():
    print("\n[grouping] by-parent: items sharing parent + panel land on one circuit")
    items = [
        _make(21, "L1", "BYPARENT", bucket=BUCKET_BYPARENT, parent=999),
        _make(22, "L1", "BYPARENT", bucket=BUCKET_BYPARENT, parent=999),
        _make(23, "L1", "BYPARENT", bucket=BUCKET_BYPARENT, parent=888),
        _make(24, "L2", "BYPARENT", bucket=BUCKET_BYPARENT, parent=999),
    ]
    groups = assemble_groups(items)
    fails = []
    # Expected: L1+999 (2), L1+888 (1), L2+999 (1)
    if not _check("3 byparent groups (panel+parent)", len(groups) == 3):
        fails.append("count")
    pair = next((g for g in groups if g.panel_name == "L1"
                 and "999" in str(g.key)), None)
    if not _check("L1+999 has 2 members",
                  pair is not None and pair.member_count == 2):
        fails.append("merge")
    return fails


def test_secondbyparent_isolated_from_byparent():
    print("\n[grouping] secondbyparent: separate lane from byparent even with same panel+parent")
    items = [
        _make(101, "L1", "BYPARENT", bucket=BUCKET_BYPARENT, parent=999),
        _make(102, "L1", "BYPARENT", bucket=BUCKET_BYPARENT, parent=999),
        _make(103, "L1", "SECONDBYPARENT", bucket=BUCKET_SECONDBYPARENT, parent=999),
        _make(104, "L1", "SECONDBYPARENT", bucket=BUCKET_SECONDBYPARENT, parent=999),
    ]
    groups = assemble_groups(items)
    fails = []
    if not _check("2 groups (one byp, one byp2)", len(groups) == 2):
        fails.append("count")
    byp = next((g for g in groups if g.bucket == BUCKET_BYPARENT), None)
    sec = next((g for g in groups if g.bucket == BUCKET_SECONDBYPARENT), None)
    if not _check("byparent group has 2 members",
                  byp is not None and byp.member_count == 2):
        fails.append("byp")
    if not _check("secondbyparent group has 2 members",
                  sec is not None and sec.member_count == 2):
        fails.append("byp2")
    return fails


def test_combined_circuit_split_no_overrides():
    print("\n[clients] combined-circuit `&` splitter (no client overrides anymore)")
    from circuit_clients.base import CircuitClient
    from circuit_clients.heb import HebClient
    from circuit_clients.pf import PfClient
    fails = []
    base = CircuitClient()
    if not _check("base splits 6&9 -> ['6','9']",
                  base.split_combined_circuit_token("6&9") == ["6", "9"]):
        fails.append("split")
    if not _check("base splits '6 & 9' (whitespace tolerant)",
                  base.split_combined_circuit_token("6 & 9") == ["6", "9"]):
        fails.append("split ws")
    if not _check("base returns [] for plain '12'",
                  base.split_combined_circuit_token("12") == []):
        fails.append("plain")
    if not _check("base returns [] for empty",
                  base.split_combined_circuit_token("") == []):
        fails.append("empty")
    if not _check("base has no override map",
                  base.combined_circuit_load_overrides(["6", "9"]) == {}):
        fails.append("base override")
    if not _check("heb has no override map",
                  HebClient().combined_circuit_load_overrides(["6", "9"]) == {}):
        fails.append("heb override")
    if not _check("pf has no override map",
                  PfClient().combined_circuit_load_overrides(["6", "9"]) == {}):
        fails.append("pf override")
    return fails


def test_position_rules_match():
    print("\n[clients] match_position_rule: HEB only; PF has none")
    from circuit_clients.heb import HebClient
    from circuit_clients.pf import PfClient
    heb = HebClient()
    pf = PfClient()
    fails = []
    cases_heb = [
        ("CHECKSTAND RECEPT", "CHECKSTAND RECEPT"),
        ("checkstand jbox", "CHECKSTAND JBOX"),
        ("Receptacle for SELF CHECKOUT 4", "SELF CHECKOUT"),
        ("ELECTRIC CARTS bay 3", "ELECTRIC CARTS"),
        ("DESK QUAD - section A", "DESK QUAD"),
        ("DESK DUPLEX something", "DESK DUPLEX"),
        ("DESK something", None),
    ]
    for raw, expected in cases_heb:
        rule = heb.match_position_rule(raw)
        kw = (rule or {}).get("keyword")
        if not _check("heb {!r} -> {!r}".format(raw, expected), kw == expected):
            fails.append(("heb", raw))
    if not _check("pf has no position rules at all",
                  pf.match_position_rule("TVTRUSS") is None):
        fails.append(("pf", "tvtruss"))
    return fails


def test_load_priority_lookup():
    print("\n[clients] get_load_priority — PF only carries fitness equipment priorities")
    from circuit_clients.heb import HebClient
    from circuit_clients.pf import PfClient
    heb = HebClient()
    pf = PfClient()
    fails = []
    if not _check("pf TREADMILL -> 0", pf.get_load_priority("treadmill") == 0):
        fails.append("pf treadmill")
    if not _check("pf STAIRMASTER -> 2", pf.get_load_priority("STAIRMASTER") == 2):
        fails.append("pf stair")
    if not _check("pf SINK 1 -> 3", pf.get_load_priority("Sink 1") == 3):
        fails.append("pf sink")
    if not _check("heb does not carry fitness priorities (default 99)",
                  heb.get_load_priority("treadmill") == 99):
        fails.append("heb no fit")
    if not _check("heb unknown -> default 99",
                  heb.get_load_priority("Mixer") == 99):
        fails.append("default")
    return fails


def test_run_keyword_options():
    print("\n[clients] run_keyword_options: only DEDICATED/BYPARENT/SECONDBYPARENT + extras")
    from circuit_clients.heb import HebClient
    from circuit_clients.pf import PfClient
    heb = HebClient()
    pf = PfClient()
    fails = []
    heb_opts = heb.run_keyword_options(present_in_groups=["NIGHTLIGHT", "12"])
    standard_set = {"DEDICATED", "BYPARENT", "SECONDBYPARENT"}
    if not _check("heb has the 3 standard tokens",
                  standard_set.issubset(set(heb_opts))):
        fails.append("heb standard")
    if not _check("heb has CASECONTROLLER", "CASECONTROLLER" in heb_opts):
        fails.append("heb cc")
    if not _check("heb does NOT contain TVTRUSS",
                  "TVTRUSS" not in heb_opts):
        fails.append("heb tvtruss")
    if not _check("heb does NOT contain EMERGENCY (removed)",
                  "EMERGENCY" not in heb_opts):
        fails.append("heb emergency")
    if not _check("heb surfaces NIGHTLIGHT (present token)",
                  "NIGHTLIGHT" in heb_opts):
        fails.append("heb present")
    pf_opts = pf.run_keyword_options()
    if not _check("pf has 3 standard tokens",
                  standard_set.issubset(set(pf_opts))):
        fails.append("pf standard")
    if not _check("pf does NOT contain STANDARD (removed)",
                  "STANDARD" not in pf_opts):
        fails.append("pf standard")
    if not _check("pf does NOT contain TVTRUSS (removed)",
                  "TVTRUSS" not in pf_opts):
        fails.append("pf tvtruss")
    return fails


def test_classify_recognizes_second_tokens():
    print("\n[clients] classify_circuit_token recognises SECONDBYPARENT family")
    from circuit_clients.base import CircuitClient
    c = CircuitClient()
    cases = [
        ("SECONDBYPARENT", BUCKET_SECONDBYPARENT),
        ("secondbyparent", BUCKET_SECONDBYPARENT),
        ("SECONDCIRCUITBYPARENT", BUCKET_SECONDBYPARENT),
        ("secondcircuitbyparent", BUCKET_SECONDBYPARENT),
    ]
    fails = []
    for raw, expected in cases:
        bucket, _ = c.classify_circuit_token(raw)
        if not _check("{} -> {}".format(raw, expected), bucket == expected):
            fails.append(raw)
    return fails


def test_normal_with_blank_panel_marks_review():
    print("\n[grouping] normal with blank panel/circuit -> needs-review group")
    items = [
        _make(41, "", "12"),
        _make(42, "L1", ""),
    ]
    groups = assemble_groups(items)
    fails = []
    if not _check("2 needs-review groups", len(groups) == 2):
        fails.append("count")
    if not _check("both flagged needs_review",
                  all(g.needs_review for g in groups)):
        fails.append("flag")
    panels = {g.panel_name for g in groups}
    if not _check("blank-panel group uses sentinel",
                  NEEDS_REVIEW_PANEL in panels):
        fails.append("panel sentinel")
    return fails


def test_user_overrides_take_precedence():
    print("\n[grouping] user_override fields drive grouping when set")
    a = _make(51, "L1", "12")
    a.user_panel = "L9"
    a.user_circuit_token = "20"
    b = _make(52, "L9", "20")
    groups = assemble_groups([a, b])
    fails = []
    if not _check("user override merges with b on L9/20",
                  len(groups) == 1 and groups[0].member_count == 2):
        fails.append("merge")
    return fails


def test_phase_tracker_round_robin():
    print("\n[phasing] PanelPhaseTracker: 1-pole rotates A/B/C, 2-pole rotates AB/BC/CA, 3-pole always ABC")
    t = PanelPhaseTracker()
    fails = []
    if not _check("1-pole #1 -> A", t.next_phase_for_panel("L1", 1) == PHASE_A):
        fails.append("p1 a")
    if not _check("1-pole #2 -> B", t.next_phase_for_panel("L1", 1) == PHASE_B):
        fails.append("p1 b")
    if not _check("1-pole #3 -> C", t.next_phase_for_panel("L1", 1) == PHASE_C):
        fails.append("p1 c")
    if not _check("1-pole #4 -> A (wrap)",
                  t.next_phase_for_panel("L1", 1) == PHASE_A):
        fails.append("wrap")
    # Different panel keeps its own counter
    t2 = PanelPhaseTracker()
    if not _check("3-pole always ABC", t2.next_phase_for_panel("L2", 3) == PHASE_ABC):
        fails.append("3p")
    if not _check("3-pole stays ABC even after 2-pole",
                  t2.next_phase_for_panel("L2", 3) == PHASE_ABC):
        fails.append("3p again")
    if not _check("2-pole #1 -> AB", t2.next_phase_for_panel("L2", 2) == PHASE_AB):
        fails.append("2p ab")
    return fails


def test_dist_system_selection():
    print("\n[phasing] select_distribution_system_id")
    fails = []
    if not _check("None for empty list",
                  select_distribution_system_id([], 1) is None):
        fails.append("empty")
    if not _check("first for single-pole",
                  select_distribution_system_id([10, 20], 1) == 10):
        fails.append("1p")
    if not _check("last for 3-pole when multiple",
                  select_distribution_system_id([10, 20, 30], 3) == 30):
        fails.append("3p")
    if not _check("first when only one option even at 3-pole",
                  select_distribution_system_id([42], 3) == 42):
        fails.append("single")
    return fails


def run():
    fails = []
    for fn in (
        test_normal_grouping,
        test_dedicated_one_per_item,
        test_byparent_groups_by_parent,
        test_secondbyparent_isolated_from_byparent,
        test_classify_recognizes_second_tokens,
        test_combined_circuit_split_no_overrides,
        test_position_rules_match,
        test_load_priority_lookup,
        test_run_keyword_options,
        test_normal_with_blank_panel_marks_review,
        test_user_overrides_take_precedence,
        test_phase_tracker_round_robin,
        test_dist_system_selection,
    ):
        f = fn()
        if f:
            fails.append((fn.__name__, f))
    return fails


if __name__ == "__main__":
    fails = run()
    if fails:
        print("\nFAILURES:")
        for name, f in fails:
            print("  ", name, f)
    else:
        print("\nALL TESTS PASSED")
