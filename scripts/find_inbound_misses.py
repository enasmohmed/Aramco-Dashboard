#!/usr/bin/env python3
"""
Parse ARAMCO Inbound Report xlsx without pandas (stdlib only) and list Miss shipments
per facility/month using the same SLA rules as dashboard/views.py (_sla_working_days_between_dates).
Usage:
  python3 scripts/find_inbound_misses.py path/to/latest.xlsx [--facility Jeddah] [--month Apr]
"""
from __future__ import annotations

import argparse
import re
import sqlite3
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timedelta

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

EXCEL_BASE = datetime(1899, 12, 30)


def col_to_idx(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c.upper()) - ord("A") + 1)
    return n - 1


def parse_sheet_xml(z: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list[str]]:
    root = ET.fromstring(z.read(sheet_path))
    rows_out: list[list[str]] = []
    for row in root.findall("m:sheetData/m:row", NS):
        cells: dict[int, str] = {}
        for c in row.findall("m:c", NS):
            ref = c.get("r") or ""
            ma = re.match(r"^([A-Z]+)", ref)
            if not ma:
                continue
            ci = col_to_idx(ma.group(1))
            t = c.get("t")
            v_el = c.find("m:v", NS)
            if v_el is None or v_el.text is None:
                continue
            v = v_el.text
            if t == "s":
                v = shared[int(v)]
            cells[ci] = v
        if not cells:
            continue
        last = max(cells)
        line = [""] * (last + 1)
        for i, v in cells.items():
            line[i] = v
        rows_out.append(line)
    return rows_out


def load_first_sheet(path: str) -> list[list[str]]:
    with zipfile.ZipFile(path, "r") as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.findall("m:si", NS):
                parts: list[str] = []
                for t in si.iter():
                    if t.tag.endswith("}t") and t.text:
                        parts.append(t.text)
                shared.append("".join(parts))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = wb.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet")
        rid = sheets[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            raise RuntimeError("Could not resolve first sheet path")
        sheet_path = "xl/" + target.lstrip("/")
        return parse_sheet_xml(z, sheet_path, shared)


def norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())


def find_col(headers: list[str], *needles: str) -> int | None:
    if not headers:
        return None
    low = [norm_name(h) for h in headers]
    for n in needles:
        nn = norm_name(n)
        for i, h in enumerate(low):
            if nn in h or h in nn:
                return i
    return None


def excel_serial_to_dt(x) -> datetime | None:
    if x is None or x == "":
        return None
    try:
        xf = float(x)
        if xf > 1e7 or xf < 30000:
            return None
        return EXCEL_BASE + timedelta(days=xf)
    except (TypeError, ValueError, OverflowError):
        return None


def _is_sla_non_working_day(date_val) -> bool:
    if date_val.weekday() == 4:
        return True
    if date_val.month == 2 and date_val.day == 22:
        return True
    if date_val.month == 3 and 19 <= date_val.day <= 24:
        return True
    return False


def _sla_working_days_between_dates(start_ts, end_ts) -> int:
    if start_ts is None or end_ts is None:
        return 0
    start_d = start_ts.date() if hasattr(start_ts, "date") else start_ts
    end_d = end_ts.date() if hasattr(end_ts, "date") else end_ts
    if start_d > end_d:
        return 0
    if start_d == end_d:
        return 0
    days = 0
    current = start_d
    while current < end_d:
        if not _is_sla_non_working_day(current):
            days += 1
        current += timedelta(days=1)
    return days


def norm_facility(val) -> str | None:
    v = (str(val) or "").strip().lower()
    if not v:
        return None
    if "riyadh" in v or v == "ruh":
        return "Riyadh"
    if "dammam" in v or "damam" in v:
        return "Dammam"
    if "jeddah" in v or "jedd" in v:
        return "Jeddah"
    if "central" in v or "وسط" in v:
        return "Riyadh"
    if "eastern" in v or "east" in v or "شرقي" in v:
        return "Dammam"
    if "western" in v or "west" in v or "غربي" in v:
        return "Jeddah"
    cleaned = re.sub(r"\s+", " ", str(val).strip())
    if not cleaned or cleaned.lower() in {"facility", "region", "site", "location", "warehouse"}:
        return None
    return cleaned


def analyze(path: str, want_facility: str | None, want_month: str | None):
    rows = load_first_sheet(path)
    header_idx = None
    for i, row in enumerate(rows[:25]):
        joined = " ".join(str(x).lower() for x in row if x)
        if "shipment" in joined and ("facility" in joined or "region" in joined) and (
            "create" in joined or "creation" in joined
        ):
            header_idx = i
            break
    if header_idx is None:
        print("Could not find header row", file=sys.stderr)
        sys.exit(1)
    headers = [str(x).strip() for x in rows[header_idx]]
    i_fac = find_col(headers, "Facility", "Region")
    i_ship = find_col(headers, "Shipment_nbr", "Shipment", "Shipment_ID")
    i_create = find_col(headers, "Create shipment", "Create shipemnt", "Create", "Ship_Date")
    i_rcv = find_col(headers, "Received LPN", "Receiving_Complete_Date", "Verified_Date")
    if i_fac is None or i_ship is None or i_create is None or i_rcv is None:
        print("Missing columns", i_fac, i_ship, i_create, i_rcv, file=sys.stderr)
        sys.exit(1)

    data_rows = rows[header_idx + 1 :]
    by_ship: dict[str, dict] = defaultdict(
        lambda: {
            "facility": None,
            "create_min": None,
            "received_max": None,
        }
    )
    for row in data_rows:
        if i_ship >= len(row):
            continue
        ship = str(row[i_ship]).strip()
        if not ship or ship.lower() in {"nan", "none"}:
            continue
        if ship.startswith("250"):
            continue
        fac_raw = row[i_fac] if i_fac < len(row) else ""
        fn = norm_facility(fac_raw)
        if not fn:
            continue
        c = excel_serial_to_dt(row[i_create] if i_create < len(row) else None)
        r = excel_serial_to_dt(row[i_rcv] if i_rcv < len(row) else None)
        st = by_ship[ship]
        st["facility"] = fn
        if c and (st["create_min"] is None or c < st["create_min"]):
            st["create_min"] = c
        if r and (st["received_max"] is None or r > st["received_max"]):
            st["received_max"] = r

    want_month = (want_month or "").strip()[:3].title() if want_month else None
    want_facility = (want_facility or "").strip() or None

    misses: list[tuple[str, str, str, int]] = []
    for ship, st in by_ship.items():
        c, r = st["create_min"], st["received_max"]
        fac = st["facility"]
        if not c or not r or not fac:
            continue
        month = c.strftime("%b")
        if want_month and month != want_month:
            continue
        if want_facility and fac.lower() != want_facility.lower():
            continue
        days = _sla_working_days_between_dates(c, r)
        is_hit = days <= 1
        if not is_hit:
            misses.append((ship, fac, month, days))

    return misses


def upsert_hit(sqlite_path: str, shipment_nbr: str, facility: str):
    conn = sqlite3.connect(sqlite_path)
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM dashboard_inboundshipmentremark WHERE shipment_nbr = ? AND facility = ?",
        (shipment_nbr, facility),
    )
    row = cur.fetchone()
    if row:
        cur.execute(
            "UPDATE dashboard_inboundshipmentremark SET status_override = ?, updated_at = datetime('now') WHERE id = ?",
            ("Hit", row[0]),
        )
    else:
        cur.execute(
            "INSERT INTO dashboard_inboundshipmentremark (shipment_nbr, facility, remark, updated_at, status_override) VALUES (?, ?, '', datetime('now'), ?)",
            (shipment_nbr, facility, "Hit"),
        )
    conn.commit()
    conn.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--facility", default="Jeddah")
    ap.add_argument("--month", default="Apr")
    ap.add_argument("--sqlite", help="If set, upsert Hit for first listed miss")
    args = ap.parse_args()
    misses = analyze(args.xlsx, args.facility, args.month)
    print(f"Misses ({args.facility} {args.month}): {len(misses)}")
    for ship, fac, month, days in misses:
        print(f"  {ship}  {fac}  {month}  working_days={days}")
    if args.sqlite and misses:
        ship, fac, _, _ = misses[0]
        upsert_hit(args.sqlite, ship, fac)
        print(f"Upserted Hit for {ship} @ {fac} in {args.sqlite}")


if __name__ == "__main__":
    main()
