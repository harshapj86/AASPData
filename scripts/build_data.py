#!/usr/bin/env python3
"""
build_data.py — converts the Aptronix master Excel workbook into data.json,
the file the dashboard (index.html) fetches at load time.

USAGE
    python scripts/build_data.py path/to/workbook.xlsx [--out data.json]

WHAT IT EXPECTS IN THE WORKBOOK
    - One or more "raw data" sheets: any sheet whose name starts with "Raw"
      (case-insensitive), e.g. "Raw Data 24-25", "Raw Data 2025-26",
      "Rawdata 26-27". New raw-data sheets (next fiscal year, etc.) are
      picked up automatically — no code change needed, just add the sheet
      and re-run. Required columns: Branch ID, TXNDate, Transaction Type,
      Item_Name, Business Unit, Category, Executive, GP AT Amount, Total.
      A "TXN Month" column is NOT required — the month is derived directly
      from TXNDate. If a "TXN Month" column happens to be present, it's
      used only as a fallback for any row whose TXNDate fails to parse.
    - "Location Master": ARM | State | Centre Name | Location Type | Ship to ID
      (Ship to ID is optional — only needed for centres that receive Spare
      Parts Incentives. A free-text note in the column immediately after
      Ship to ID, e.g. "Until 2025", marks a reassigned ID as only valid
      through that year — see parse_spare_parts_incentives().)
    - "Spare Parts Incentives" (optional): a raw, transaction-level export
      pasted fresh each period — no manual editing. Required columns:
      Ship-To, Invoice Date, Incentives Amount. Adds Apple's spare-part
      incentive on top of transaction GP, for the centres and months it
      covers. If this sheet is absent, GP is transaction-only everywhere
      (same as if no incentive data existed at all).
    - "CSAT": Ship-To | Metrics | <date columns...>   (values as fractions, e.g. 0.95)
    - "Revenue Target", "GP Target", "CSAT Target", "License Target":
      Branch ID | <month columns...>
      (header row is row 4 — rows 1-3 are a title + instructions + a blank row)

WHAT IT PRODUCES (data.json)
    {
      months: [{MonthKey, MonthLabel, FY, Quarter, FYQ}, ...],
      gpCoveredMonths: [MonthKey, ...],
      centres: [{Centre, State, ARM, LocationType, City}, ...],
      fact_month:   [[MonthKey, Centre, Revenue, GP, Txns], ...],
      fact_bu:      [[MonthKey, Centre, BusinessUnit, Revenue, GP, Txns], ...],
      fact_cat:     [[MonthKey, Centre, Category, Revenue, GP, Txns], ...],
      fact_pf:      [[MonthKey, Centre, ProductFamily, Revenue, GP, Txns], ...],
      fact_item:    [[MonthKey, Centre, ItemName, Revenue, GP, Txns], ...],
      fact_txntype: [[MonthKey, Centre, TxnType, Revenue, GP, Txns], ...],
      fact_exec:    [[Centre, Executive, Revenue, GP, Txns], ...],
      fact_day:     [[DateStr(YYYY-MM-DD), Centre, Revenue, GP, Txns], ...],
      gp_incentives:   [[MonthKey, Centre, IncentiveAmount], ...],
      csat:            [[MonthKey, Centre, CSATPercent], ...],
      target_revenue:  [[MonthKey, Centre, TargetRevenue], ...],
      target_gp:       [[MonthKey, Centre, TargetGP], ...],
      target_csat:     [[MonthKey, Centre, TargetCSATPercent], ...],
      target_licenses: [[MonthKey, Centre, TargetLicenseCount], ...],
    }

KNOWN DATA-QUALITY CAVEAT (see README)
    The raw data's "Product Family" column is inconsistently filled — some
    rows have real product families, some have an item code pasted in by
    mistake, many are blank. Blank/NA values are bucketed into "Other".
    This mirrors the original dashboard's behaviour but inherits the
    underlying sheet's messiness; clean up "Product Family" at the source
    if you want cleaner fact_pf breakdowns.
"""
import sys
import re
import json
import argparse
from collections import defaultdict

import pandas as pd

MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}
MONTH_NUM = {v.lower(): k for k, v in MONTH_ABBR.items()}
# also accept full month names, just in case a sheet uses them
FULL_MONTH_NUM = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}


def month_key_from_any(value):
    """Parse a month cell into a YYYYMM int. Handles:
       - pandas/py datetime (Excel sometimes silently converts 'Jun 24' to a date)
       - strings like 'April 24', 'Apr 24', 'Apr-24', 'Jun 2024'
    Returns None if it can't be parsed (caller should skip such columns/rows).
    """
    if value is None:
        return None
    if hasattr(value, "year") and hasattr(value, "month"):
        return value.year * 100 + value.month
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("-", " ").replace("_", " ")
    parts = s.split()
    if len(parts) != 2:
        return None
    mon_raw, yr_raw = parts[0].lower(), parts[1]
    mon = MONTH_NUM.get(mon_raw) or FULL_MONTH_NUM.get(mon_raw)
    if mon is None:
        return None
    try:
        yr = int(yr_raw)
    except ValueError:
        return None
    if yr < 100:
        yr += 2000
    return yr * 100 + mon


def month_label(month_key):
    y, m = divmod(month_key, 100)
    return f"{MONTH_ABBR[m]} {y % 100:02d}"


def fy_quarter(month_key):
    """Indian fiscal year: Apr-Mar. FY label uses the two years it spans."""
    y, m = divmod(month_key, 100)
    if m >= 4:
        fy_start = y
    else:
        fy_start = y - 1
    fy = f"FY{fy_start % 100:02d}-{(fy_start + 1) % 100:02d}"
    if m in (4, 5, 6):
        q = "Q1"
    elif m in (7, 8, 9):
        q = "Q2"
    elif m in (10, 11, 12):
        q = "Q3"
    else:
        q = "Q4"
    return fy, q


def load_raw_data(xls):
    """Concatenate every sheet whose name starts with 'raw' (case-insensitive)."""
    raw_sheet_names = [s for s in xls.sheet_names if s.strip().lower().startswith("raw")]
    if not raw_sheet_names:
        raise SystemExit("No sheet found whose name starts with 'Raw' — nothing to build from.")
    print(f"Raw-data sheets found: {raw_sheet_names}")
    # Column names have drifted slightly between years (e.g. "TXN Month" vs
    # "TXNMonth"). Normalise per-sheet BEFORE concatenating — normalising
    # after concat can produce duplicate-named columns instead of merging
    # them, since sheets with the old/new spelling didn't share a column.
    def normalise_columns(df):
        rename_map = {}
        for col in df.columns:
            key = str(col).strip().lower().replace(" ", "")
            if key == "txnmonth":
                rename_map[col] = "TXN Month"
            elif key == "branchid":
                rename_map[col] = "Branch ID"
            elif key == "branchcity":
                rename_map[col] = "Branch City"
            elif key == "transactiontype":
                rename_map[col] = "Transaction Type"
            elif key in ("item_name", "itemname"):
                rename_map[col] = "Item_Name"
            elif key == "businessunit":
                rename_map[col] = "Business Unit"
            elif key == "category":
                rename_map[col] = "Category"
            elif key == "productfamily":
                rename_map[col] = "Product Family"
            elif key == "executive":
                rename_map[col] = "Executive"
            elif key == "gpatamount":
                rename_map[col] = "GP AT Amount"
            elif key == "total":
                rename_map[col] = "Total"
            elif key == "txndate":
                rename_map[col] = "TXNDate"
        return df.rename(columns=rename_map)

    frames = []
    for name in raw_sheet_names:
        df = normalise_columns(xls.parse(name))
        df["__source_sheet"] = name
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True, sort=False)

    required = ["Branch ID", "TXNDate", "Transaction Type",
                "Item_Name", "Business Unit", "Category", "Executive",
                "GP AT Amount", "Total"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise SystemExit(f"Raw data is missing expected column(s): {missing}")

    def to_date_str(v):
        """TXNDate is a real datetime in some sheets, a plain DD-MM-YYYY
        (or D-M-YYYY) text string in others (seen in 'Raw Data 24-25'), and
        — if a cell's date number-format ever gets lost/reset (seen in a
        later export of 'Raw Data 2025-26', where the column reads back as
        plain 'General'-formatted integers) — a bare Excel serial number.
        Handle all three, or fact_day silently loses whole sheets' worth of
        days, which corrupts MIN_DATE and breaks date-range filtering for
        any year affected."""
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        if isinstance(v, (int, float)):
            # Excel serial date (epoch 1899-12-30, matching Excel's own
            # leap-year quirk). Guarded to a plausible date range so a
            # stray non-date number in this column isn't misread as one.
            if 36526 <= v <= 73050:  # ~2000-01-01 .. ~2100-01-01
                try:
                    from datetime import datetime as _dt, timedelta as _td
                    return (_dt(1899, 12, 30) + _td(days=float(v))).strftime("%Y-%m-%d")
                except (OverflowError, ValueError):
                    return None
            return None
        s = str(v).strip()
        if not s:
            return None
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                from datetime import datetime as _dt
                return _dt.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None
    raw["DateStr"] = raw["TXNDate"].apply(to_date_str)
    bad_dates = raw["DateStr"].isna().sum()
    if bad_dates:
        print(f"WARNING: {bad_dates} row(s) had an unparseable TXNDate — "
              f"they're dropped entirely (no month, no day — see 'TXN Month "
              f"fallback' note below if a TXN Month column exists and could "
              f"cover these instead).")

    # MonthKey is derived straight from TXNDate — a "TXN Month" column is
    # NOT required. (It used to be the only source of MonthKey; checked
    # against a real workbook where both existed, TXNDate-derived and
    # TXN-Month-derived months agreed on all 95k+ rows, so this is a safe
    # simplification, not a behaviour change.) If a "TXN Month" column
    # exists, it's used only as a fallback for rows whose TXNDate didn't
    # parse — the row still needs a valid month from *some* source.
    raw["MonthKey"] = raw["DateStr"].apply(lambda d: int(d[:4])*100+int(d[5:7]) if d else None)
    if "TXN Month" in raw.columns:
        needs_fallback = raw["MonthKey"].isna()
        if needs_fallback.any():
            fallback_mk = raw.loc[needs_fallback, "TXN Month"].apply(month_key_from_any)
            raw.loc[needs_fallback, "MonthKey"] = fallback_mk
            n_recovered = fallback_mk.notna().sum()
            if n_recovered:
                print(f"Recovered {n_recovered} row(s) using the 'TXN Month' column "
                      f"where TXNDate didn't parse.")
    bad = raw["MonthKey"].isna().sum()
    if bad:
        print(f"WARNING: {bad} row(s) had no usable date and were dropped.")
    raw = raw.dropna(subset=["MonthKey"])
    raw["MonthKey"] = raw["MonthKey"].astype(int)

    if "Product Family" not in raw.columns:
        raw["Product Family"] = None
    raw["Product Family"] = raw["Product Family"].fillna("Other")
    raw.loc[raw["Product Family"].astype(str).str.strip().isin(["", "NA", "None"]), "Product Family"] = "Other"

    # Every column we ever group by must never be NaN — a NaN grouping key
    # produces a literal `NaN` token in the JSON output, which is valid
    # Python but INVALID per the JSON spec, so browsers refuse to parse the
    # file entirely (this broke the dashboard until caught by testing).
    fallback_label = {
        "Business Unit": "Unspecified",
        "Category": "Unspecified",
        "Item_Name": "Unspecified item",
        "Executive": "Unassigned",
        "Transaction Type": "Unspecified",
    }
    for col, label in fallback_label.items():
        if col not in raw.columns:
            raw[col] = label
            continue
        raw[col] = raw[col].fillna(label)
        blank_mask = raw[col].astype(str).str.strip().isin(["", "NA", "None", "nan"])
        raw.loc[blank_mask, col] = label

    raw = normalize_casing(raw, ["Business Unit", "Category", "Product Family",
                                  "Item_Name", "Executive", "Transaction Type"])

    raw["Total"] = pd.to_numeric(raw["Total"], errors="coerce").fillna(0.0)
    raw["GP AT Amount"] = pd.to_numeric(raw["GP AT Amount"], errors="coerce").fillna(0.0)

    return raw


def parse_spare_parts_incentives(xls):
    """The 'Spare Parts Incentives' sheet is a raw, transaction-level export
    (pasted fresh each week, no manual editing) — one row per incentive
    credit, identified by Ship-To ID and Invoice Date. This is the source
    for Apple's spare-part incentive on top of transaction GP (the raw
    transaction rows have no such column). It replaces the old hand-
    maintained 'GP' report sheet as that source; the downstream effect is
    identical (added on top of transaction GP for the months it covers),
    but the figure is now traced to a raw source instead of a manually
    reconciled one.

    Ship-To ID -> Centre comes from Location Master's 'Ship to ID' column.
    A Ship-To ID can be REASSIGNED when a centre closes and a newer centre
    takes it over (seen in practice: a closed centre's ID handed to a new
    one). Location Master marks the old row with a free-text note in the
    column immediately after 'Ship to ID' (e.g. "Until 2025"); an
    incentive is attributed to whichever centre's validity window actually
    covers its invoice date — the bounded (noted) rule wins if the date
    falls within it, otherwise the unbounded (currently active) rule does.
    This is what stops a reassigned ID from being double-counted or
    credited to the wrong centre.

    Incentives currently apply only to full-fledged Service Centres, not
    Repair Drop Off locations — this falls out naturally, since drop-off
    locations have no Ship-To ID in Location Master to match against.

    Returns {(MonthKey, Centre): incentive_amount} — ADDITIVE (an amount to
    add on top of transaction GP), not a full replacement figure like the
    old parse_gp_report returned.
    """
    if "Spare Parts Incentives" not in xls.sheet_names:
        print("WARNING: no 'Spare Parts Incentives' sheet found — GP will NOT include Apple's spare-part incentive.")
        return {}
    if "Location Master" not in xls.sheet_names:
        print("WARNING: 'Spare Parts Incentives' sheet exists but 'Location Master' doesn't — "
              "can't map Ship-To IDs to centres, so incentives will be skipped entirely.")
        return {}

    lm = xls.parse("Location Master", header=0)
    lm.columns = [str(c).strip() for c in lm.columns]
    ship_col = next((c for c in lm.columns if c.replace(" ", "").lower() == "shiptoid"), None)
    if ship_col is None:
        print("WARNING: Location Master has no 'Ship to ID' column — incentives will be skipped entirely.")
        return {}
    ship_col_pos = list(lm.columns).index(ship_col)
    note_col = lm.columns[ship_col_pos + 1] if ship_col_pos + 1 < len(lm.columns) else None

    # ship_id -> [(centre, valid_until_date_str_or_None)]
    rules_by_ship_id = defaultdict(list)
    for _, r in lm.iterrows():
        centre = str(r.get("Centre Name", "")).strip()
        if not centre or centre.lower() == "grand total":
            continue
        raw_ship_id = r.get(ship_col)
        if pd.isna(raw_ship_id):
            continue
        try:
            ship_id = int(raw_ship_id)
        except (TypeError, ValueError):
            continue
        valid_until = None
        if note_col is not None:
            note = r.get(note_col)
            if isinstance(note, str) and note.strip():
                m = re.search(r"(\d{4})", note)
                if m:
                    valid_until = f"{m.group(1)}-12-31"
        rules_by_ship_id[ship_id].append((centre, valid_until))

    def resolve_centre(ship_id, invoice_date_str):
        rules = rules_by_ship_id.get(ship_id)
        if not rules:
            return None
        bounded = sorted([r for r in rules if r[1] is not None], key=lambda r: r[1])
        for centre, valid_until in bounded:
            if invoice_date_str <= valid_until:
                return centre
        unbounded = [r for r in rules if r[1] is None]
        if unbounded:
            return unbounded[0][0]
        # every rule for this ID is bounded and expired before this date —
        # fall back to the latest-expiring one rather than dropping the row
        return bounded[-1][0] if bounded else None

    sp = xls.parse("Spare Parts Incentives")
    sp.columns = [str(c).strip() for c in sp.columns]
    required = ["Ship-To", "Invoice Date", "Incentives Amount"]
    missing = [c for c in required if c not in sp.columns]
    if missing:
        print(f"WARNING: 'Spare Parts Incentives' sheet is missing column(s) {missing} — "
              f"skipping incentives entirely.")
        return {}

    def to_date_str(v):
        if pd.isna(v):
            return None
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip().replace("Z", "")
        try:
            return pd.Timestamp(s).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return None

    totals = defaultdict(float)
    unmatched = 0
    for _, r in sp.iterrows():
        raw_ship_id = r["Ship-To"]
        if pd.isna(raw_ship_id):
            continue
        try:
            ship_id = int(raw_ship_id)
        except (TypeError, ValueError):
            continue
        date_str = to_date_str(r["Invoice Date"])
        if not date_str:
            continue
        amount = r["Incentives Amount"]
        if pd.isna(amount):
            continue
        centre = resolve_centre(ship_id, date_str)
        if centre is None:
            unmatched += 1
            continue
        mk = int(date_str[:4]) * 100 + int(date_str[5:7])
        totals[(mk, centre)] += float(amount)

    if unmatched:
        print(f"WARNING: {unmatched} Spare Parts Incentives row(s) had a Ship-To ID "
              f"not found in Location Master — skipped.")
    result = {k: round(v, 2) for k, v in totals.items()}
    print(f"Spare Parts Incentives: {len(sp)} row(s) -> {len(result)} centre/month figure(s), "
          f"total Rs.{sum(result.values()):,.2f}")
    return result


def normalize_casing(df, columns):
    """Collapse values that differ only by case or leading/trailing
    whitespace into one canonical spelling — the most frequent variant
    already in the data — so the same real-world category doesn't get
    split into separate rows across every breakdown and trend chart (e.g.
    'Incident fees' vs 'Incident Fees' showing as two different business
    units, each with fake zero-gaps wherever the other casing was used
    that month). Applied to every grouping column, not only the ones with
    known drift today, so a new inconsistency introduced by a future
    sheet gets caught automatically instead of silently fragmenting a
    chart again.
    """
    for col in columns:
        if col not in df.columns:
            continue
        stripped = df[col].astype(str).str.strip()
        norm_key = stripped.str.lower()
        canonical = stripped.groupby(norm_key).transform(lambda s: s.value_counts().idxmax())
        distinct_before = stripped.nunique()
        distinct_after = canonical.nunique()
        if distinct_before != distinct_after:
            print(f"Normalized casing/whitespace for '{col}': "
                  f"{distinct_before} -> {distinct_after} distinct values.")
        df[col] = canonical
    return df


def group_sum(df, group_cols):
    g = df.groupby(group_cols, dropna=False).agg(
        Revenue=("Total", "sum"),
        GP=("GP AT Amount", "sum"),
        Txns=("Total", "size"),
    ).reset_index()
    return g


def to_records(df, cols):
    """Build the [[...], [...]] row lists the dashboard expects, in `cols` order."""
    records = []
    for _, r in df.iterrows():
        row = []
        for c in cols:
            v = r[c]
            if c in ("Revenue", "GP"):
                v = round(float(v), 2)
            elif c in ("Txns", "MonthKey"):
                v = int(v)
            row.append(v)
        records.append(row)
    return records


def build_centres(xls, raw):
    centres_in_raw = sorted(raw["Branch ID"].dropna().unique().tolist())

    loc_master = {}
    if "Location Master" in xls.sheet_names:
        lm = xls.parse("Location Master")
        lm.columns = [str(c).strip() for c in lm.columns]
        for _, r in lm.iterrows():
            centre = str(r.get("Centre Name", "")).strip()
            if not centre or centre.lower() == "grand total":
                continue  # summary row, not a real centre
            loc_master[centre] = {
                "ARM": str(r.get("ARM", "")).strip() or "Unmapped",
                "State": str(r.get("State", "")).strip() or "Unmapped",
                "LocationType": str(r.get("Location Type", "")).strip() or "Service Centre",
            }
    else:
        print("WARNING: no 'Location Master' sheet found — all centres will be Unmapped.")

    # Most common Branch City per centre, from the raw data itself.
    city_by_centre = {}
    if "Branch City" in raw.columns:
        city_counts = raw.groupby(["Branch ID", "Branch City"]).size().reset_index(name="n")
        for centre, sub in city_counts.groupby("Branch ID"):
            top = sub.sort_values("n", ascending=False).iloc[0]
            city_by_centre[centre] = str(top["Branch City"])

    centres = []
    for centre in centres_in_raw:
        meta = loc_master.get(centre, {"ARM": "Unmapped", "State": "Unmapped", "LocationType": "Repair Drop Off"})
        centres.append({
            "Centre": centre,
            "State": meta["State"],
            "ARM": meta["ARM"],
            "LocationType": meta["LocationType"],
            "City": city_by_centre.get(centre, ""),
        })
    return centres


def melt_month_sheet(xls, sheet_name, value_scale=1.0):
    """For Revenue Target / GP Target / CSAT Target sheets: header row is
    row 4 (index 3), first column is Branch ID, remaining columns are months.
    Returns [[MonthKey, Centre, Value], ...] skipping blank cells.
    """
    if sheet_name not in xls.sheet_names:
        print(f"WARNING: sheet '{sheet_name}' not found — skipping.")
        return []
    df = xls.parse(sheet_name, header=3)
    if df.empty or df.columns[0] is None:
        return []
    df = df.rename(columns={df.columns[0]: "Centre"})
    out = []
    for col in df.columns[1:]:
        mk = month_key_from_any(col)
        if mk is None:
            continue
        for _, r in df.iterrows():
            centre = r["Centre"]
            val = r[col]
            if pd.isna(centre) or pd.isna(val):
                continue
            try:
                val = round(float(val), 2)
            except (TypeError, ValueError):
                continue
            out.append([mk, str(centre).strip(), val])
    return out


def build_csat(xls):
    """CSAT sheet layout: <blank/short-label> | Ship-To | Metrics | <month columns...>
    Column A is a short informal label (e.g. "Banjara hills") that does NOT
    match the canonical centre names used everywhere else in the dashboard
    (fact_month, centres, filters all use "Service Banjara Hills" etc, from
    Branch ID / Location Master's Centre Name). "Ship-To" is the column that
    actually holds the matching canonical name here despite the name overlap
    with Location Master's unrelated "Ship to ID" column used for incentives
    — confusing, but confirmed against real data. Using column A instead
    (an earlier version of this function did) silently breaks every CSAT
    lookup dashboard-wide: nothing matches, so CSAT appears empty for every
    centre and every date range, with no error anywhere to point at why.

    "Metrics" also isn't just a spacer column — a sheet-wide "Average" row
    has no centre and Metrics=None (not "CSAT"), so filtering on
    Metrics=="CSAT" is what keeps that summary row out of the per-centre
    data, not just a formality.
    """
    if "CSAT" not in xls.sheet_names:
        print("WARNING: no 'CSAT' sheet found.")
        return []
    df = xls.parse("CSAT")
    if df.empty:
        return []
    if len(df.columns) < 3:
        print("WARNING: 'CSAT' sheet doesn't have the expected "
              "<label> | Ship-To | Metrics | <months...> layout — skipping CSAT.")
        return []
    df = df.rename(columns={df.columns[1]: "Centre", df.columns[2]: "Metric"})
    df = df[df["Metric"].astype(str).str.strip().str.upper() == "CSAT"]
    out = []
    for col in df.columns[3:]:
        mk = month_key_from_any(col)
        if mk is None:
            continue
        for _, r in df.iterrows():
            centre = r["Centre"]
            val = r[col]
            if pd.isna(centre) or pd.isna(val):
                continue
            try:
                pct = round(float(val) * 100.0, 2)
            except (TypeError, ValueError):
                continue
            out.append([mk, str(centre).strip(), pct])
    return out


def build_master_data(xls):
    """Runs the full pipeline and returns the master data dict (everything,
    no access restrictions) — the single source of truth that both the
    plain single-bundle build and the per-tier bundler build from.
    """
    raw = load_raw_data(xls)
    print(f"Loaded {len(raw):,} raw transaction line(s) across "
          f"{raw['MonthKey'].nunique()} month(s) and {raw['Branch ID'].nunique()} centre(s).")

    fact_month = group_sum(raw, ["MonthKey", "Branch ID"]).rename(columns={"Branch ID": "Centre"})
    fact_day = group_sum(raw, ["DateStr", "Branch ID"]).rename(columns={"Branch ID": "Centre"})
    fact_day = fact_day.dropna(subset=["DateStr"])

    # Apply Spare Parts Incentives on top of the transaction-derived GP —
    # see parse_spare_parts_incentives(). This adjusts:
    #   - fact_month (the headline figure used by Executive Overview, GP
    #     tab, trends)
    #   - fact_day, by adding each month's incentive onto that centre's
    #     first day of the month. The KPI cards compute their "exact"
    #     totals from fact_day (so a custom date range can be precise to
    #     the day), so fact_month alone isn't enough — without this, the
    #     headline numbers silently keep excluding the incentive even
    #     though fact_month looks correct. There's no data attributing the
    #     incentive to a specific day, so "posted on day 1 of the month" is
    #     a deliberate, documented modeling choice: it's captured
    #     correctly by any FY/quarter/full-month filter, and only missed
    #     by a custom range that excludes day 1.
    # Every OTHER breakdown (fact_bu, fact_cat, fact_pf, fact_item,
    # fact_txntype, fact_exec) intentionally stays transaction-only —
    # there's no basis to attribute the incentive to a business unit,
    # category, item, or associate either. This is exactly what the
    # dashboard's "GP coverage" caveat banner communicates.
    incentives = parse_spare_parts_incentives(xls)
    covered_months = set()
    fact_month = fact_month.set_index(["MonthKey", "Centre"], drop=False)
    fact_day = fact_day.set_index(["DateStr", "Centre"], drop=False)
    for (mk, centre), incentive_amount in incentives.items():
        covered_months.add(mk)
        if (mk, centre) in fact_month.index:
            raw_gp = float(fact_month.loc[(mk, centre), "GP"])
            new_gp = round(raw_gp + incentive_amount, 2)
            fact_month.loc[(mk, centre), "GP"] = new_gp
        else:
            # Incentive exists for a centre/month with no matching
            # transactions at all — still worth surfacing rather than
            # silently dropping.
            raw_gp = 0.0
            new_gp = round(incentive_amount, 2)
            new_row = pd.DataFrame([{"MonthKey": mk, "Centre": centre,
                                      "Revenue": 0.0, "GP": new_gp, "Txns": 0}])
            fact_month = pd.concat([fact_month.reset_index(drop=True), new_row], ignore_index=True)
            fact_month = fact_month.set_index(["MonthKey", "Centre"], drop=False)

        delta = round(new_gp - raw_gp, 2)
        if abs(delta) > 0.005:
            first_day = f"{mk // 100:04d}-{mk % 100:02d}-01"
            day_key = (first_day, centre)
            if day_key in fact_day.index:
                fact_day.loc[day_key, "GP"] = float(fact_day.loc[day_key, "GP"]) + delta
            else:
                new_day_row = pd.DataFrame([{"DateStr": first_day, "Centre": centre,
                                              "Revenue": 0.0, "GP": delta, "Txns": 0}])
                fact_day = pd.concat([fact_day.reset_index(drop=True), new_day_row], ignore_index=True)
                fact_day = fact_day.set_index(["DateStr", "Centre"], drop=False)
    fact_month = fact_month.reset_index(drop=True)
    fact_day = fact_day.reset_index(drop=True)
    print(f"Applied Spare Parts Incentives to {len(covered_months)} month(s): "
          f"{sorted(covered_months)}")
    # Additive record of JUST the incentive component, for the GP tab's
    # Repair GP / Incentives / Total GP breakdown — fact_month's GP stays
    # the single source of truth for "Total GP" everywhere else; this is
    # purely so the dashboard can show what fraction of it is incentive
    # without having to reverse-engineer that from two other tables.
    gp_incentives = [[mk, centre, round(amount, 2)] for (mk, centre), amount in incentives.items()]

    fact_bu = group_sum(raw, ["MonthKey", "Branch ID", "Business Unit"]).rename(
        columns={"Branch ID": "Centre", "Business Unit": "BU"})
    fact_cat = group_sum(raw, ["MonthKey", "Branch ID", "Category"]).rename(
        columns={"Branch ID": "Centre"})
    fact_pf = group_sum(raw, ["MonthKey", "Branch ID", "Product Family"]).rename(
        columns={"Branch ID": "Centre", "Product Family": "PF"})
    fact_item = group_sum(raw, ["MonthKey", "Branch ID", "Item_Name"]).rename(
        columns={"Branch ID": "Centre", "Item_Name": "Item"})
    fact_txntype = group_sum(raw, ["MonthKey", "Branch ID", "Transaction Type"]).rename(
        columns={"Branch ID": "Centre", "Transaction Type": "TxnType"})
    fact_exec = group_sum(raw, ["Branch ID", "Executive"]).rename(
        columns={"Branch ID": "Centre"})

    month_keys = sorted(fact_month["MonthKey"].unique().tolist())
    months = []
    for mk in month_keys:
        fy, q = fy_quarter(mk)
        months.append({
            "MonthKey": mk,
            "MonthLabel": month_label(mk),
            "FY": fy,
            "Quarter": q,
            "FYQ": f"{fy} {q}",
        })

    return {
        "months": months,
        "gpCoveredMonths": sorted(covered_months),
        "centres": build_centres(xls, raw),
        "fact_month": to_records(fact_month, ["MonthKey", "Centre", "Revenue", "GP", "Txns"]),
        "fact_bu": to_records(fact_bu, ["MonthKey", "Centre", "BU", "Revenue", "GP", "Txns"]),
        "fact_cat": to_records(fact_cat, ["MonthKey", "Centre", "Category", "Revenue", "GP", "Txns"]),
        "fact_pf": to_records(fact_pf, ["MonthKey", "Centre", "PF", "Revenue", "GP", "Txns"]),
        "fact_item": to_records(fact_item, ["MonthKey", "Centre", "Item", "Revenue", "GP", "Txns"]),
        "fact_txntype": to_records(fact_txntype, ["MonthKey", "Centre", "TxnType", "Revenue", "GP", "Txns"]),
        "fact_exec": to_records(fact_exec, ["Centre", "Executive", "Revenue", "GP", "Txns"]),
        "fact_day": to_records(fact_day, ["DateStr", "Centre", "Revenue", "GP", "Txns"]),
        "gp_incentives": gp_incentives,
        "csat": build_csat(xls),
        "target_revenue": melt_month_sheet(xls, "Revenue Target"),
        "target_gp": melt_month_sheet(xls, "GP Target"),
        "target_csat": melt_month_sheet(xls, "CSAT Target"),
        "target_licenses": melt_month_sheet(xls, "License Target"),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", help="Path to the master .xlsx file")
    ap.add_argument("--out", default="data.json", help="Output path for data.json")
    args = ap.parse_args()

    print(f"Reading {args.workbook} ...")
    xls = pd.ExcelFile(args.workbook)
    data = build_master_data(xls)

    import os
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        # allow_nan=False makes json.dump raise loudly here (during the build,
        # where it's visible in the GitHub Action log) instead of silently
        # writing an invalid `NaN` token that would break JSON.parse in every
        # browser and leave the dashboard stuck on its loading screen.
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)

    print(f"Wrote {args.out} "
          f"({sum(len(v) if isinstance(v, list) else 0 for v in data.values()):,} total rows across all tables).")


if __name__ == "__main__":
    main()
