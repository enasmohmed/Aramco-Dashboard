# -*- coding: utf-8 -*-
"""
استيراد بيانات Transportation KPI من ملف إكسل.
يُستخدم من أدمن Transportation KPI عند رفع ملف إكسل.
"""
import re
import pandas as pd


def _norm(val):
    return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())


def _find_col(headers, *names):
    for c in headers:
        cn = _norm(c)
        for n in names:
            if cn == _norm(n) or _norm(n) in cn or cn in _norm(n):
                return c
    return None


def parse_transportation_excel(excel_path):
    """
    يقرأ شيت Transportation من ملف إكسل ويرجع قائمة سجلات جاهزة لـ TransportationKPI.
    كل عنصر: section, month, kpi, region, total, sum_value, hit, miss, achieved_percent, target_percent.
    يرجع أيضاً (rows, errors) حيث errors قائمة رسائل إن وجدت.
    """
    errors = []
    rows = []
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception as e:
        return [], [f"لا يمكن فتح الملف: {e}"]

    transport_sheet = next(
        (
            s
            for s in xls.sheet_names
            if (s or "").lower().strip()
            and (
                "transportation" in (s or "").lower()
                or "transport" in (s or "").lower()
            )
        ),
        None,
    )
    if not transport_sheet:
        return [], ["لم يُعثر على شيت باسم Transportation أو Transport"]

    raw = pd.read_excel(
        excel_path,
        sheet_name=transport_sheet,
        engine="openpyxl",
        header=None,
        dtype=str,
    ).fillna("")

    nrows = raw.shape[0]
    ncols = raw.shape[1]
    kpi_titles_norm = (
        _norm("Delivery Fulfilment"),
        _norm("On Time Delivery"),
        _norm("PODs submission"),
    )
    month_jan_norm = _norm("January")
    month_feb_norm = _norm("February")
    section_in_norm = _norm("Inbound")
    section_out_norm = _norm("Outbound")

    def row_key(row, max_cols=8):
        parts = []
        for c in range(min(max_cols, ncols)):
            if c < len(row):
                parts.append(str(row.iloc[c]).strip())
        return _norm(" ".join(parts))

    def first_cell(row):
        return _norm(str(row.iloc[0]) if len(row) > 0 else "")

    def row_matches(row, *options):
        rk = row_key(row)
        fc = first_cell(row)
        return any(
            _norm(opt) in rk or _norm(opt) in fc or fc == _norm(opt) for opt in options
        )

    sections_order = ["Inbound", "Outbound"]
    by_section = {sec: {"January": [], "February": []} for sec in sections_order}
    current_section = None
    current_month = None
    i = 0

    while i < nrows:
        row = raw.iloc[i]
        fc = first_cell(row)
        rk = row_key(row)

        if section_out_norm in rk or section_out_norm in fc:
            current_section = "Outbound"
            i += 1
            continue
        if section_in_norm in rk or section_in_norm in fc:
            current_section = "Inbound"
            i += 1
            continue
        if "january" in rk or "january" in fc or fc == month_jan_norm:
            current_month = "January"
            i += 1
            continue
        if "february" in rk or "february" in fc or fc == month_feb_norm:
            current_month = "February"
            i += 1
            continue

        is_kpi_row = False
        kpi_title = None
        for kpi in ("Delivery Fulfilment", "On Time Delivery", "PODs submission"):
            kn = _norm(kpi)
            if kn in rk or kn in fc or fc == kn:
                is_kpi_row = True
                kpi_title = kpi
                break
        if not is_kpi_row and ("delivery" in rk or "delivery" in fc) and (
            "fulfilment" in rk
            or "fulfilment" in fc
            or "fulfillment" in rk
            or "fulfillment" in fc
        ):
            is_kpi_row = True
            kpi_title = "Delivery Fulfilment"

        if is_kpi_row and kpi_title and current_section and current_month:
            header_idx = i + 1
            while header_idx < nrows:
                header_row = raw.iloc[header_idx]
                if any(str(header_row.iloc[c]).strip() for c in range(min(ncols, 10))):
                    break
                header_idx += 1
            if header_idx >= nrows:
                i += 1
                continue
            header_row = raw.iloc[header_idx]
            headers = [str(header_row.iloc[c]).strip() or "" for c in range(ncols)]
            while headers and not headers[-1]:
                headers.pop()
            if not headers:
                headers = [
                    str(header_row.iloc[c]).strip() or f"Col_{c}"
                    for c in range(min(15, ncols))
                ]
            col_count = len(headers)
            data_rows = []
            j = header_idx + 1
            while j < nrows:
                data_row = raw.iloc[j]
                first = first_cell(data_row)
                drk = row_key(data_row)
                if row_matches(data_row, "January", "February", "Inbound", "Outbound"):
                    break
                if any(k in drk or k in first for k in kpi_titles_norm):
                    break
                cells = [
                    str(data_row.iloc[c]).strip() if c < len(data_row) else ""
                    for c in range(col_count)
                ]
                if not any(cells):
                    j += 1
                    continue
                data_rows.append(dict(zip(headers, cells)))
                j += 1
            by_section[current_section][current_month].append(
                {"title": kpi_title, "columns": headers, "data": data_rows}
            )
            i = j
            continue
        i += 1

    # تحويل by_section إلى قائمة سجلات TransportationKPI
    region_defaults = ["Fuchs-Yanbu", "Fuchs-Jeddah"]
    kpi_titles = ("Delivery Fulfilment", "On Time Delivery", "PODs submission")

    for section in sections_order:
        for month in ("January", "February"):
            tables = by_section[section][month]
            for k, kpi_name in enumerate(kpi_titles):
                tbl = tables[k] if k < len(tables) else None
                data = (tbl.get("data") or []) if tbl else []
                if len(data) < 2:
                    data = data + [{}] * (2 - len(data))
                row0, row1 = data[0], data[1] if len(data) > 1 else {}
                headers = (tbl.get("columns") or []) if tbl else []
                c_region = _find_col(headers, "region", "REGION")
                c_total = _find_col(headers, "total", "TOTAL")
                c_hit = _find_col(headers, "hit", "HIT")
                c_miss = _find_col(headers, "miss", "MISS")
                c_achieved = _find_col(
                    headers, "achieved", "achived", "ACHIEVED", "ACHIVED"
                )
                c_total_sub = _find_col(
                    headers,
                    "total submitted",
                    "totalsubmitted",
                    "total submitted pods",
                    "total submitted pod",
                    "total submitted pod's",
                    "total submitted pods submission",
                )

                def val(r, col):
                    if not col:
                        return ""
                    return str(r.get(col, "") or "").strip()

                total0 = val(row0, c_total)
                total1 = val(row1, c_total)
                try:
                    sum_val = (
                        str(int(float(total0 or 0)) + int(float(total1 or 0)))
                        if kpi_name == "On Time Delivery"
                        else ""
                    )
                except (ValueError, TypeError):
                    sum_val = ""
                achieved0 = val(row0, c_achieved) or val(row1, c_achieved)
                if achieved0 and "%" in achieved0:
                    achieved0 = achieved0.replace("%", "").strip()
                target = "98"

                for idx, (r, region_label) in enumerate(
                    [(row0, region_defaults[0]), (row1, region_defaults[1])]
                ):
                    region = val(r, c_region) or region_label
                    rows.append(
                        {
                            "section": section,
                            "month": month,
                            "kpi": kpi_name,
                            "region": region,
                            "total": val(r, c_total),
                            "sum_value": sum_val,
                            "hit": val(r, c_hit),
                            "miss": val(r, c_miss),
                            "total_submitted": val(r, c_total_sub),
                            "achieved_percent": achieved0,
                            "target_percent": target,
                        }
                    )

    return rows, errors
