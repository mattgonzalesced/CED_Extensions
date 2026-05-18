# -*- coding: utf-8 -*-
"""One-shot: derive HEB_profiles_V5.4.yaml from V5.3.

Rule (per user): a profile is DEPENDENT (allow_parentless: false) when
its profile name matches ANY of:
  * starts with 3+ digits
  * starts with "HEB" (case-insensitive)
  * contains "VendorProvided" (case-insensitive, no space)
  * contains a ":" anywhere
Every other profile is INDEPENDENT (allow_parentless: true).

Surgical block rewrite: only the profile-level ``allow_parentless:``
line in each ``equipment_definitions`` entry is touched. All other
lines (formatting, ordering, quoting, comments) pass through verbatim.
"""

import io
import re

SRC = r"c:\CED_Extensions\HEB_profiles_V5.3.yaml"
DST = r"c:\CED_Extensions\HEB_profiles_V5.4.yaml"

PROFILE_START_RE = re.compile(r"^- prompt_on_parent_mismatch:")
TOPLEVEL_RE = re.compile(r"^[A-Za-z_]")          # column-0 key (not "- ")
NAME_RE = re.compile(r"^  name:[ \t]?(.*)$")
ALLOW_RE = re.compile(r"^(  allow_parentless:)[ \t]*(.*?)[ \t]*$")
THREE_DIGIT_RE = re.compile(r"^\d{3}")


def decode_scalar(raw):
    s = raw.strip()
    if len(s) >= 2 and s[0] == "'" and s[-1] == "'":
        return s[1:-1].replace("''", "'")
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return s[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    return s


def is_dependent(name):
    if not name:
        return False
    if THREE_DIGIT_RE.match(name):
        return True
    if name.upper().startswith("HEB"):
        return True
    if "vendorprovided" in name.lower():
        return True
    if ":" in name:
        return True
    return False


# Read with newline="" so Python does NOT translate CRLF -> LF; the
# source uses CRLF and we must round-trip it exactly.
with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    text = fh.read()
nl = "\r\n" if "\r\n" in text else "\n"
lines = text.split(nl)

# 1. Find profile block boundaries.
starts = [i for i, ln in enumerate(lines) if PROFILE_START_RE.match(ln)]
bounds = []
for k, s in enumerate(starts):
    e = len(lines)
    for j in range(s + 1, len(lines)):
        if PROFILE_START_RE.match(lines[j]) or TOPLEVEL_RE.match(lines[j]):
            e = j
            break
    bounds.append((s, e))

n_dep = n_indep = n_no_name = n_no_flag = 0
samples_dep, samples_indep = [], []

for (s, e) in bounds:
    name_idx = allow_idx = None
    for i in range(s, e):
        if name_idx is None:
            m = NAME_RE.match(lines[i])
            if m:
                name_idx = i
        if allow_idx is None and ALLOW_RE.match(lines[i]):
            allow_idx = i
    if name_idx is None:
        n_no_name += 1
        continue
    name = decode_scalar(NAME_RE.match(lines[name_idx]).group(1))
    dep = is_dependent(name)
    new_val = "false" if dep else "true"
    if allow_idx is None:
        n_no_flag += 1
        continue
    am = ALLOW_RE.match(lines[allow_idx])
    lines[allow_idx] = "{} {}".format(am.group(1), new_val)
    if dep:
        n_dep += 1
        if len(samples_dep) < 8:
            samples_dep.append(name)
    else:
        n_indep += 1
        if len(samples_indep) < 8:
            samples_indep.append(name)

with io.open(DST, "w", encoding="utf-8", newline="") as fh:
    fh.write(nl.join(lines))

print("profiles            :", len(bounds))
print("dependent  (false)  :", n_dep)
print("independent (true)  :", n_indep)
print("blocks w/o name     :", n_no_name)
print("blocks w/o allow    :", n_no_flag)
print("--- sample DEPENDENT ---")
for s in samples_dep:
    print("  ", repr(s))
print("--- sample INDEPENDENT ---")
for s in samples_indep:
    print("  ", repr(s))
