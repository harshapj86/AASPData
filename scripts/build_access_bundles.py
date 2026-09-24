#!/usr/bin/env python3
"""
build_access_bundles.py
========================
Splits the full data.json (produced by build_data.py) into one scoped copy
per access tier, each paired with its own copy of index.html, ready to sit
behind its own Cloudflare Access policy.

Run this AFTER build_data.py, from the repo root:
    python scripts/build_access_bundles.py

Reads:
    data.json        (the full, unscoped dataset)
    index.html        (the dashboard itself — copied as-is into each bundle)
    access/roles.csv  (who maps to which scope — see that file's header)

Writes, under access/site/:
    admin/index.html + admin/data.json                    (everyone with an
                                                             Admin row, full data)
    arm/<arm-slug>/index.html + data.json                  (one per ARM)
    centre/<centre-slug>/index.html + data.json            (one per centre
                                                             with a CentreManager
                                                             row in roles.csv)

Also writes access/cloudflare-policies.md — for each bundle, the exact list
of emails that need a matching Cloudflare Access policy on that path. This
script only produces the files; it does not (and cannot) configure
Cloudflare itself — see that file's instructions for the one-time manual
setup in the Cloudflare dashboard.

A centre with NO CentreManager row in roles.csv simply gets no /centre/
bundle at all — nothing to configure for it, nothing extra deployed.
"""

import csv
import json
import re
import sys
from pathlib import Path
from collections import defaultdict

ROOT = Path(__file__).resolve().parent.parent
DATA_PATH = ROOT / "data.json"
INDEX_PATH = ROOT / "index.html"
ROLES_PATH = ROOT / "access" / "roles.csv"
SITE_OUT = ROOT / "access" / "site"
POLICY_DOC = ROOT / "access" / "cloudflare-policies.md"

# Fields in data.json keyed by [MonthKey, Centre, ...] or [DateStr, Centre, ...]
# or [Centre, ...] — the position of "Centre" varies by table, given here so
# filtering is table-driven rather than repeating logic per table.
CENTRE_INDEXED_TABLES = {
    "fact_month": 1, "fact_day": 1, "fact_bu": 1, "fact_cat": 1,
    "fact_pf": 1, "fact_item": 1, "fact_txntype": 1, "fact_exec": 0,
    "csat": 1, "target_revenue": 1, "target_gp": 1,
    "target_csat": 1, "target_licenses": 1, "gp_incentives": 1,
}


def slugify(name):
    s = re.sub(r"[^a-zA-Z0-9]+", "-", name.strip().lower()).strip("-")
    return s


def load_roles():
    if not ROLES_PATH.exists():
        print(f"ERROR: {ROLES_PATH} not found. Create it first — see the "
              f"template comments in that file for the expected columns.", file=sys.stderr)
        sys.exit(1)
    # '#'-prefixed lines are comments/inactive rows (including the doc header
    # at the top of the template) — stripped before CSV parsing so the real
    # header row ("email,role,scope,name") is always what DictReader sees
    # first, regardless of how many comment lines precede it in the file.
    with open(ROLES_PATH, encoding="utf-8") as f:
        content_lines = [ln for ln in f if not ln.lstrip().startswith("#") and ln.strip()]
    reader = csv.DictReader(content_lines)
    roles = []
    for row in reader:
        email = (row.get("email") or "").strip()
        role = (row.get("role") or "").strip()
        scope = (row.get("scope") or "").strip()
        if not email:
            continue
        if role not in ("Admin", "ARM", "CentreManager"):
            print(f"WARNING: roles.csv row for {email} has unrecognised "
                  f"role '{role}' (expected Admin/ARM/CentreManager) — skipped.", file=sys.stderr)
            continue
        if role != "Admin" and not scope:
            print(f"WARNING: roles.csv row for {email} (role={role}) has "
                  f"no scope value — skipped.", file=sys.stderr)
            continue
        if "YOUR-EMAIL-HERE" in email.upper():
            print(f"WARNING: roles.csv has an unfilled placeholder row "
                  f"(role={role}, scope={scope}) — skipped. Fill in a real "
                  f"email to activate it.", file=sys.stderr)
            continue
        roles.append({"email": email, "role": role, "scope": scope})
    return roles


def filter_data(full_data, allowed_centres):
    """allowed_centres=None means no filtering (Admin, full data)."""
    if allowed_centres is None:
        return full_data
    allowed = set(allowed_centres)
    out = dict(full_data)
    out["centres"] = [c for c in full_data["centres"] if c["Centre"] in allowed]
    for table, centre_idx in CENTRE_INDEXED_TABLES.items():
        rows = full_data.get(table, [])
        out[table] = [r for r in rows if r[centre_idx] in allowed]
    return out


def write_bundle(folder, data):
    folder.mkdir(parents=True, exist_ok=True)
    with open(folder / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    # index.html is identical everywhere — same file, same fetch('./data.json')
    # relative path, so each bundle folder is a fully self-contained site.
    (folder / "index.html").write_text(INDEX_PATH.read_text(encoding="utf-8"), encoding="utf-8")


def main():
    if not DATA_PATH.exists():
        print(f"ERROR: {DATA_PATH} not found — run build_data.py first.", file=sys.stderr)
        sys.exit(1)
    if not INDEX_PATH.exists():
        print(f"ERROR: {INDEX_PATH} not found.", file=sys.stderr)
        sys.exit(1)

    with open(DATA_PATH, encoding="utf-8") as f:
        full_data = json.load(f)

    centre_arm = {c["Centre"]: c["ARM"] for c in full_data["centres"]}
    all_centres = set(centre_arm)

    roles = load_roles()

    # Group emails by (role, scope) so N people sharing a scope produce ONE
    # bundle, not N duplicates — and so cloudflare-policies.md can list every
    # email that needs to go on that path's Access policy in one place.
    groups = defaultdict(list)
    for r in roles:
        groups[(r["role"], r["scope"])].append(r["email"])

    bundles = []  # (path, label, emails, data)

    for (role, scope), emails in groups.items():
        if role == "Admin":
            path = "admin"
            label = "Admin (full company)"
            data = filter_data(full_data, None)
        elif role == "ARM":
            arm_centres = [c for c, a in centre_arm.items() if a == scope]
            if not arm_centres:
                print(f"WARNING: no centres found with ARM == '{scope}' "
                      f"(check spelling matches Location Master exactly) — "
                      f"skipping bundle for {emails}.", file=sys.stderr)
                continue
            path = f"arm/{slugify(scope)}"
            label = f"ARM — {scope} ({len(arm_centres)} centres)"
            data = filter_data(full_data, arm_centres)
        else:  # CentreManager
            if scope not in all_centres:
                print(f"WARNING: centre '{scope}' not found in Location Master "
                      f"(check spelling) — skipping bundle for {emails}.", file=sys.stderr)
                continue
            path = f"centre/{slugify(scope)}"
            label = f"Centre Manager — {scope}"
            data = filter_data(full_data, [scope])

        write_bundle(SITE_OUT / path, data)
        bundles.append((path, label, sorted(set(emails)), data))
        print(f"  {path}: {label} — {len(data['centres'])} centre(s), "
              f"{len(emails)} email(s)")

    # Cloudflare Access policy reference doc — this script can't configure
    # Cloudflare itself (no API credentials, and this is a one-time manual
    # setup better done deliberately in their dashboard than scripted blind),
    # so it writes exactly what each policy needs instead.
    lines = [
        "# Cloudflare Access policies needed",
        "",
        "Auto-generated by build_access_bundles.py — do not hand-edit; it will",
        "be overwritten on the next build. Edit access/roles.csv instead.",
        "",
        "For each row below, create one Cloudflare Access Application "
        "(Zero Trust -> Access -> Applications -> Add an application -> Self-hosted):",
        "",
    ]
    for path, label, emails, data in bundles:
        lines.append(f"## {label}")
        lines.append(f"- **Path:** `yourdomain.pages.dev/{path}/`")
        lines.append(f"- **Policy emails:** {', '.join(emails)}")
        lines.append("")
    POLICY_DOC.parent.mkdir(parents=True, exist_ok=True)
    POLICY_DOC.write_text("\n".join(lines), encoding="utf-8")

    print(f"\nWrote {len(bundles)} bundle(s) under {SITE_OUT}")
    print(f"Wrote policy reference: {POLICY_DOC}")


if __name__ == "__main__":
    main()
