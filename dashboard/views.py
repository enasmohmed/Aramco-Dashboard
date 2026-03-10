# views.py
import datetime
import shutil
import os
import re
from io import BytesIO
from collections import OrderedDict

import pandas as pd
import numpy as np
from django.conf import settings
from django.contrib import messages
from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, HttpResponse
from django.views import View
from .forms import ExcelUploadForm
from django.core.cache import cache

from django.views.decorators.cache import cache_control
import json, traceback, os
from datetime import date, timedelta
from django.db.models import Q
from django.template.loader import render_to_string
from calendar import month_abbr, month_name
import calendar as calendar_module

from django.views.decorators.csrf import csrf_exempt
from django.utils.decorators import method_decorator
from django.utils.text import slugify
from django.utils import timezone

from .models import (
    MeetingPoint,
    InboundShipmentRemark,
    OutboundOrderRemark,
    ExcelSheetCache,
    DashboardDataCache,
)


def make_json_serializable(df):

    def convert_value(x):
        if isinstance(x, (pd.Timestamp, pd.Timedelta)):
            return x.isoformat()
        elif isinstance(x, (datetime.datetime, datetime.date, datetime.time)):
            return x.isoformat()
        elif isinstance(x, (np.int64, np.int32)):
            return int(x)
        elif isinstance(x, (np.float64, np.float32)):
            return float(x)
        elif isinstance(x, (np.ndarray, list, dict)):
            return str(x)
        else:
            return x

    return df.applymap(convert_value)


def _dataframe_to_cache_rows(df):
    """Convert DataFrame to list of dicts suitable for JSONField (ExcelSheetCache)."""
    if df is None or df.empty:
        return []

    def _safe_val(x):
        if x is None or (isinstance(x, float) and (pd.isna(x) or x != x)):
            return None
        if isinstance(x, (pd.Timestamp, datetime.datetime, datetime.date)):
            try:
                return x.isoformat() if hasattr(x, "isoformat") else str(x)
            except Exception:
                return None
        if isinstance(x, (np.integer, np.int64, np.int32)):
            return int(x)
        if isinstance(x, (np.floating, np.float64, np.float32)):
            return None if pd.isna(x) else float(x)
        return x

    rows = []
    for _, row in df.iterrows():
        rows.append({str(k): _safe_val(v) for k, v in row.items()})
    return rows


def _sanitize_for_json(obj):
    """Convert numpy/pandas types to native Python for JsonResponse."""
    if obj is None or isinstance(obj, (bool, str)):
        return obj
    if isinstance(obj, np.ndarray):
        return [_sanitize_for_json(v) for v in obj.tolist()]
    if isinstance(obj, (np.integer, np.int64, np.int32)):
        return int(obj)
    if isinstance(obj, (np.floating, np.float64, np.float32)):
        try:
            v = float(obj)
            if np.isnan(v) or np.isinf(v):
                return None
            return v
        except (ValueError, TypeError):
            return None
    if isinstance(obj, (pd.Timestamp, pd.Timedelta, datetime.datetime, datetime.date)):
        return obj.isoformat() if hasattr(obj, "isoformat") else str(obj)
    if isinstance(obj, (int, float)) and (obj != obj or abs(obj) == float("inf")):
        return None  # NaN or Inf
    try:
        if pd.isna(obj) and not isinstance(obj, (dict, list, tuple)):
            return None
    except (ValueError, TypeError):
        pass
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_for_json(v) for v in obj]
    return obj


def _get_excel_path_for_request(request):
    """يرجع مسار ملف الإكسل المرفوع من الجلسة أو المجلد الافتراضي."""
    if not request:
        return None
    folder = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
    if not os.path.isdir(folder):
        return None
    path = request.session.get("uploaded_excel_path")
    if path and os.path.isfile(path):
        return path
    # الملف الرئيسي لكل التابات (ماعدا Dashboard): all_sheet.xlsx / all_sheet.xlsm
    for name in ["all_sheet.xlsx", "all_sheet.xlsm", "all sheet.xlsx", "all sheet.xlsm"]:
        p = os.path.join(folder, name)
        if os.path.isfile(p):
            return p
    return None


# اسم ملف الداشبورد الثابت (شيت تاني للتاب Dashboard فقط)
DASHBOARD_EXCEL_FILENAME = "Aramco_Tamer3PL_KPI_Dashboard.xlsx"

# داتا Inbound الافتراضية (للكروت والشارت) — نفس فكرة chart_data في rejection
INBOUND_DEFAULT_KPI = {
    "number_of_vehicles": 12,
    "number_of_shipments": 287,
    "number_of_pallets": 1105,
    "total_quantity": 65400,
    "total_quantity_display": "65.4k",
}
# الداتا اللي بتظهر على شارت Pending Shipments (label, value, pct, color)
INBOUND_DEFAULT_PENDING_SHIPMENTS = [
    {"label": "In Transit", "value": "1%", "pct": 1, "color": "#87CEEB"},
    {"label": "Receiving Complete", "value": "96%", "pct": 96, "color": "#2E7D32"},
    {"label": "Verified", "value": "3%", "pct": 3, "color": "#1565C0"},
]

# داتا الشارتات الافتراضية للداشبورد (نفس فكرة chart_data في rejection — لو مفيش إكسل نستخدمها)
DASHBOARD_DEFAULT_CHART_DATA = {
    "outbound_chart_data": {
        "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        "series": [40, 55, 48, 62, 58, 70],
    },
    "returns_chart_data": {
        "categories": ["Mar", "Apr", "May", "Jun", "Jul", "Aug"],
        "series": [280, 320, 300, 350, 380, 400],
    },
    "inventory_capacity_data": {"used": 78, "available": 22},
}


def _read_dashboard_charts_from_excel(excel_path):
    """
    يقرأ داتا الشارتات (Outbound, Returns, Inventory) من ملف الداشبورد لو الشيتات موجودة.
    ترجع ديكت باللي اتقرا فقط (لو مفيش داتا للشارت ترجع None للكاي) — عشان نعمل الشارتات دينامك.
    """
    result = {}
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return result
    sheet_names = [str(s).strip() for s in xls.sheet_names]

    # Outbound: من شيت Outbound_Data أو Outbound — تجميع حسب شهر لو فيه عمود شهر/تاريخ
    for out_name in ["Outbound_Data", "Outbound Data", "Outbound"]:
        if not any(out_name.lower().replace(" ", "") in s.lower().replace(" ", "") for s in sheet_names):
            continue
        sheet_name = next((s for s in sheet_names if out_name.lower() in s.lower()), None)
        if not sheet_name:
            continue
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
            if df.empty or len(df) < 2:
                break
            df.columns = [str(c).strip() for c in df.columns]
            cols_lower = {c.lower(): c for c in df.columns}
            month_col = None
            for c in cols_lower:
                if "month" in c or "date" in c:
                    month_col = cols_lower[c]
                    break
            if month_col:
                df["_m"] = pd.to_datetime(df[month_col], errors="coerce").dt.strftime("%b")
                by_month = df.groupby("_m").size().reindex(
                    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                ).dropna()
                if not by_month.empty:
                    result["outbound_chart_data"] = {
                        "categories": by_month.index.tolist(),
                        "series": by_month.values.tolist(),
                    }
            break
        except Exception:
            break

    # Returns: من شيت Return أو Rejection
    for ret_name in ["Return", "Rejection", "Returns"]:
        sheet_name = next((s for s in sheet_names if ret_name.lower() in s.lower()), None)
        if not sheet_name:
            continue
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
            if df.empty or len(df) < 2:
                break
            df.columns = [str(c).strip() for c in df.columns]
            month_col = next((c for c in df.columns if "month" in c.lower()), None)
            val_col = next(
                (c for c in df.columns if "order" in c.lower() or "booking" in c.lower() or "count" in c.lower()),
                df.columns[1] if len(df.columns) > 1 else None,
            )
            if month_col and val_col:
                summary = df[[month_col, val_col]].dropna()
                if not summary.empty:
                    try:
                        summary[val_col] = pd.to_numeric(summary[val_col].astype(str).str.replace("%", "", regex=False), errors="coerce")
                        summary = summary.dropna(subset=[val_col])
                        categories = summary[month_col].astype(str).tolist()
                        series = summary[val_col].astype(int).tolist()
                        if categories and series:
                            result["returns_chart_data"] = {"categories": categories, "series": series}
                    except Exception:
                        pass
            break
        except Exception:
            break

    # Inventory capacity: من شيت Inventory أو Capacity
    for inv_name in ["Inventory", "Capacity", "Warehouse"]:
        sheet_name = next((s for s in sheet_names if inv_name.lower() in s.lower()), None)
        if not sheet_name:
            continue
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
            if df.empty:
                break
            df.columns = [str(c).strip() for c in df.columns]
            used_col = next((c for c in df.columns if "used" in c.lower() or "utilization" in c.lower()), None)
            if used_col:
                vals = pd.to_numeric(df[used_col], errors="coerce").dropna()
                if len(vals) > 0:
                    used = int(min(100, max(0, vals.mean())))
                    result["inventory_capacity_data"] = {"used": used, "available": 100 - used}
            break
        except Exception:
            break

    return result


DASHBOARD_CACHE_JSON_FILENAME = "dashboard_cache.json"


def _get_file_mtime(path):
    """وقت آخر تعديل للملف (للمقارنة مع الكاش — لو تغيّر نقرأ من الإكسل من جديد)."""
    if not path or not os.path.isfile(path):
        return None
    try:
        return os.path.getmtime(path)
    except OSError:
        return None


def _dashboard_cache_valid(cache_data, excel_path):
    """لو ملف الإكسل اتعدّل بعد آخر كاش، الكاش غير صالح."""
    if not cache_data or not excel_path:
        return False
    stored = cache_data.get("_file_mtime")
    current = _get_file_mtime(excel_path)
    if stored is None and current is None:
        return True
    if stored is None or current is None:
        return False
    return stored == current


def _get_dashboard_cache_json_path():
    """مسار ملف JSON لكاش الداشبورد (يُحدَّث عند الرفع، ويُقرأ عند فتح الداشبورد)."""
    return os.path.join(settings.MEDIA_ROOT, "excel_uploads", DASHBOARD_CACHE_JSON_FILENAME)


def _json_serialize_value(val):
    """تحويل قيمة لصيغة قابلة لحفظها في JSON (مثلاً numpy → int/float)."""
    if isinstance(val, (np.integer, np.int64, np.int32)):
        return int(val)
    if isinstance(val, (np.floating, np.float64, np.float32)):
        return float(val)
    if isinstance(val, (np.ndarray,)):
        return val.tolist()
    if isinstance(val, (pd.Timestamp,)):
        return val.isoformat() if pd.notna(val) else None
    if isinstance(val, dict):
        return {k: _json_serialize_value(v) for k, v in val.items()}
    if isinstance(val, list):
        return [_json_serialize_value(v) for v in val]
    return val


def _dedupe_list_of_dicts(lst):
    """إزالة الصفوف المتكررة من قائمة قاموسات (اعتماداً على المحتوى)."""
    if not lst or not isinstance(lst, list):
        return lst
    seen = set()
    out = []
    for item in lst:
        if not isinstance(item, dict):
            out.append(item)
            continue
        sig = json.dumps(item, sort_keys=True, default=str)
        if sig in seen:
            continue
        seen.add(sig)
        out.append(item)
    return out


def _dedupe_cache_data(data):
    """إزالة التكرار من القوائم داخل بيانات الكاش (pending_shipments، inventory_warehouse_table، returns_region_table، إلخ)."""
    if not data or not isinstance(data, dict):
        return data
    out = dict(data)
    list_keys = ("pending_shipments", "inventory_warehouse_table", "returns_region_table")
    for key in list_keys:
        if key in out and isinstance(out[key], list):
            out[key] = _dedupe_list_of_dicts(out[key])
    if "pod_compliance_chart_data" in out and isinstance(out["pod_compliance_chart_data"], dict):
        for k in ("categories", "series"):
            if k in out["pod_compliance_chart_data"] and isinstance(out["pod_compliance_chart_data"][k], list):
                out["pod_compliance_chart_data"] = dict(out["pod_compliance_chart_data"])
                if k == "series" and isinstance(out["pod_compliance_chart_data"][k], list):
                    out["pod_compliance_chart_data"]["series"] = list(out["pod_compliance_chart_data"]["series"])
    return out


def _load_dashboard_cache_json():
    """تحميل كل كاش الداشبورد من ملف JSON. يُرجع dict مفتاحه مسار الملف وقيمته بيانات الكاش."""
    path = _get_dashboard_cache_json_path()
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_dashboard_cache_to_json(norm_path, cache_data):
    """
    إضافة/تحديث كاش مسار واحد في ملف JSON دون حذف بيانات المسارات الأخرى.
    يُزال التكرار من القوائم قبل الحفظ.
    """
    if not norm_path or not cache_data:
        return
    path = _get_dashboard_cache_json_path()
    dir_path = os.path.dirname(path)
    if dir_path:
        os.makedirs(dir_path, exist_ok=True)
    data = _load_dashboard_cache_json()
    if not isinstance(data, dict):
        data = {}
    cache_clean = _dedupe_cache_data(_json_serialize_value(cache_data))
    data[norm_path] = cache_clean
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _build_dashboard_cache_data(excel_path):
    """
    يبني قاموس بيانات الداشبورد من ملف الإكسل فقط (بدون request).
    يُستخدم لحفظ الكاش عند الرفع وقراءته من الداتابيز لفتح الداشبورد بسرعة.
    """
    if not excel_path or not os.path.exists(excel_path):
        return None
    out = {}
    try:
        inbound_data = _read_inbound_data_from_excel(excel_path)
        if inbound_data:
            out["inbound_kpi"] = inbound_data.get("inbound_kpi")
            out["pending_shipments"] = inbound_data.get("pending_shipments", [])

        charts_from_excel = _read_dashboard_charts_from_excel(excel_path)
        for key, value in charts_from_excel.items():
            if value is not None:
                out[key] = value

        outbound_data = _read_outbound_data_from_excel(excel_path)
        if outbound_data and outbound_data.get("outbound_kpi"):
            out["outbound_kpi"] = outbound_data["outbound_kpi"]
            out["outbound_kpi_keys_from_sheet"] = outbound_data.get("outbound_kpi_keys_from_sheet", [])

        pods_data = _read_pods_data_from_excel(excel_path)
        if pods_data:
            out["pod_compliance_chart_data"] = {
                "categories": pods_data.get("categories", []),
                "series": pods_data.get("series", []),
            }
            if "pod_status_breakdown" in pods_data:
                out["pod_status_breakdown"] = pods_data["pod_status_breakdown"]

        returns_data = _read_returns_data_from_excel(excel_path)
        if returns_data:
            out["returns_kpi"] = returns_data.get("returns_kpi", {})
            if "returns_chart_data" in returns_data:
                out["returns_chart_data"] = returns_data["returns_chart_data"]

        inventory_data = _read_inventory_data_from_excel(excel_path)
        if inventory_data and inventory_data.get("inventory_kpi"):
            out["inventory_kpi"] = inventory_data["inventory_kpi"]

        dashboard_wh = _read_dashboard_warehouse_from_excel(excel_path)
        if dashboard_wh:
            out["inventory_warehouse_table"] = dashboard_wh.get("inventory_warehouse_table", [])
            out["inventory_capacity_data"] = dashboard_wh.get("inventory_capacity_data", {})
        else:
            capacity_data = _read_inventory_snapshot_capacity_from_excel(excel_path)
            if capacity_data:
                out["inventory_capacity_data"] = capacity_data.get("inventory_capacity_data", {})
            warehouse_table = _read_inventory_warehouse_table_from_excel(excel_path)
            if warehouse_table:
                out["inventory_warehouse_table"] = warehouse_table.get("inventory_warehouse_table", [])

        returns_region = _read_returns_region_table_from_excel(excel_path)
        if returns_region:
            out["returns_region_table"] = returns_region.get("returns_region_table", [])
    except Exception:
        return None
    return out if out else None


def _get_dashboard_excel_path(request):
    """
    يرجّع مسار ملف إكسل الداشبورد (Aramco_Tamer3PL_KPI_Dashboard.xlsx) إن وُجد.
    مصدر الداتا لتاب Dashboard فقط؛ باقي التابات من الملف الرئيسي (all_sheet / latest).
    """
    if not request:
        return None
    folder = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
    path = request.session.get("dashboard_excel_path")
    if path and os.path.isfile(path):
        return path
    p = os.path.join(folder, DASHBOARD_EXCEL_FILENAME)
    if os.path.isfile(p):
        try:
            request.session["dashboard_excel_path"] = p
            request.session.save()
        except Exception:
            pass
        return p
    return None


def _is_dashboard_excel_filename(name):
    """يعرف إذا الملف المرفوع هو ملف الداشبورد (شيت تاني)."""
    if not name:
        return False
    n = (name or "").strip().lower()
    return "kpi_dashboard" in n or "aramco_tamer3pl" in n


def _read_inbound_data_from_excel(excel_path):
    """
    يقرأ بيانات Inbound (KPI + Pending Shipments) من ملف الإكسل (نفس ملف التابات all_sheet.xlsx).
    الشيت: "ARAMCO Inbound Report" أو شيت يحتوي "inbound".
    كروت الـ KPI (من شيت ARAMCO Inbound Report):
    - Number of Shipments: عمود Shipment_nbr — عدد القيم المميزة (حذف المتكرر)
    - Number of Pallets (LPNs): عمود LPN — عدد القيم المميزة (حذف المتكرر)
    - Total Quantity: عمود Received QTY — مجموع كل القيم
    Pending Shipments: من عمود Status في نفس الشيت إن وُجد.
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        if "ARAMCO Inbound Report" in (name or "").strip():
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "inbound" in (name or "").lower():
                sheet_name = name
                break
    if not sheet_name:
        sheet_name = xls.sheet_names[0] if xls.sheet_names else None
    if not sheet_name:
        return None
    try:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=None)
    except Exception:
        return None
    if raw.empty or raw.shape[0] < 2:
        return None

    # كشف صف الرؤوس (مثل filter_inbound): لو الصف الأول عنوان مثل "ARAMCO Inbound Report" نبحث عن صف فيه facility + shipment + create + received
    first_col = str(raw.iloc[0, 0]).strip() if raw.shape[1] else ""
    need_header_detect = (
        first_col.startswith("Unnamed:")
        or "ARAMCO" in first_col
        or "inbound report" in first_col.lower()
        or (len(first_col) > 30 and "shipment" not in first_col.lower())
    )
    df = None
    if need_header_detect and raw.shape[0] >= 2:
        header_row_idx = None
        for idx in range(min(10, raw.shape[0])):
            row = raw.iloc[idx]
            cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
            if (
                "facility" in cells or "region" in cells
            ) and ("shipment" in cells or "shipment_nbr" in cells or "shipment nbr" in cells) and (
                "create" in cells or "creation" in cells or "received" in cells or "lpn" in cells
            ):
                header_row_idx = idx
                break
        if header_row_idx is not None:
            headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[header_row_idx].values)]
            df = raw.iloc[header_row_idx + 1:].copy()
            df.columns = headers
            df = df.reset_index(drop=True)
    if df is None:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
    if df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]

    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

    def _find_col(possible_names):
        norm_map = {_norm(c): c for c in df.columns}
        for name in possible_names:
            n = _norm(name)
            if n in norm_map:
                return norm_map[n]
        for col in df.columns:
            cn = _norm(col)
            if any(_norm(x) in cn for x in possible_names):
                return col
        return None

    # Number of Shipments: عمود Shipment_nbr — عدد المميز (حذف المتكرر)
    shipment_col = _find_col([
        "Shipment_nbr", "Shipment nbr", "Shipment Nbr", "Shipment_ID", "Shipment ID",
        "Shipment No", "Shipment Number", "ShipmentNbr",
    ])
    n_shipments = int(df[shipment_col].dropna().astype(str).str.strip().nunique()) if shipment_col else 0

    # Number of Pallets (LPNs): عمود LPN — عدد المميز (حذف المتكرر)
    lpn_col = _find_col(["LPN", "LPNs", "Lpn", "LPN Nbr", "LPN_nbr", "Nbr_LPNs", "Nbr LPNs"])
    n_pallets = int(df[lpn_col].dropna().astype(str).str.strip().nunique()) if lpn_col else 0

    # Total Quantity: عمود Received QTY — مجموع كل القيم
    qty_col = _find_col([
        "Received QTY", "Received_QTY", "Received Qty", "ReceivedQTY",
        "Total_Qty", "Total Qty", "Total Quantity", "Received Quantity",
    ])
    def _to_num(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return 0
        try:
            return float(str(val).replace(",", "").strip())
        except (ValueError, TypeError):
            return 0
    n_qty = int(round(df[qty_col].fillna(0).apply(_to_num).sum())) if qty_col else 0

    if n_qty >= 1000:
        qty_display = f"{n_qty / 1000:.1f}k".rstrip("0").rstrip(".")
        if not qty_display.endswith("k"):
            qty_display += "k"
    else:
        qty_display = str(n_qty)

    inbound_kpi = {
        "number_of_vehicles": 0,
        "number_of_shipments": n_shipments,
        "number_of_pallets": n_pallets,
        "total_quantity": n_qty,
        "total_quantity_display": qty_display,
    }

    # Pending Shipments: من عمود Status في نفس الشيت — In Transit, Receiving Complete, Verified
    date_col = None
    for c in df.columns:
        cl = c.lower()
        if "date" in cl or "timestamp" in cl or "receipt" in cl:
            date_col = c
            break
    pending = []
    status_col = _find_col(["Status", "status", "Shipment Status"])
    STATUS_LABELS = (
        ("in transit", "In Transit", "#87CEEB"),
        ("receiving complete", "Receiving Complete", "#2E7D32"),
        ("verified", "Verified", "#1565C0"),
    )
    if status_col:
        df_status = df.copy()
        # تطبيع Status: حروف صغيرة + إزالة مسافات زائدة لتحمل اختلافات الكتابة
        s = df_status[status_col].fillna("").astype(str).str.strip().str.lower()
        df_status["_status_norm"] = s.str.replace(r"\s+", " ", regex=True)
        if date_col:
            df_status["_date"] = pd.to_datetime(df_status[date_col], errors="coerce")
            df_status = df_status.dropna(subset=["_date"])
            df_status["_day"] = df_status["_date"].dt.normalize()
            # كل يوم: عدد الصفوف (شحنات) لكل حالة، ثم جمع كل الأيام
            count_in_transit = 0
            count_receiving_complete = 0
            count_verified = 0
            for _day, grp in df_status.groupby("_day"):
                count_in_transit += (grp["_status_norm"] == "in transit").sum()
                count_receiving_complete += (grp["_status_norm"] == "receiving complete").sum()
                count_verified += (grp["_status_norm"] == "verified").sum()
        else:
            count_in_transit = (df_status["_status_norm"] == "in transit").sum()
            count_receiving_complete = (df_status["_status_norm"] == "receiving complete").sum()
            count_verified = (df_status["_status_norm"] == "verified").sum()
        total_pending = count_in_transit + count_receiving_complete + count_verified
        if total_pending > 0:
            for key, label, color in STATUS_LABELS:
                if key == "in transit":
                    c = count_in_transit
                elif key == "receiving complete":
                    c = count_receiving_complete
                else:
                    c = count_verified
                pct = round((c / total_pending) * 100)
                pending.append({
                    "label": label,
                    "value": f"{pct}%",
                    "pct": pct,
                    "color": color,
                })
    # لو مفيش داتا Pending من الشيت نرجع قائمة فاضية (الداتا من الشيت فقط)
    if not pending:
        pending = []

    return {"inbound_kpi": inbound_kpi, "pending_shipments": pending}


def _read_outbound_data_from_excel(excel_path):
    """
    يقرأ بيانات Outbound للداشبورد من نفس ملف التابات (all_sheet.xlsx).
    - Released Orders: من شيت ARAMCO Outbound Report، عمود Order Nbr — عدد المميز (حذف المتكرر).
    - Picked Orders: من شيت ARAMCO Outbound Report، عمود Order Nbr مع فلتر Order Status = Shipped — عدد المميز.
    - Number of Pallets (LPNs): من شيت Outbound2، عمود LPN Nbr — عدد المميز (حذف المتكرر).
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None

    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

    def _find_col(df, possible_names):
        norm_map = {_norm(c): c for c in df.columns}
        for name in possible_names:
            n = _norm(name)
            if n in norm_map:
                return norm_map[n]
        for col in df.columns:
            cn = _norm(col)
            if any(_norm(x) in cn for x in possible_names):
                return col
        return None

    released_orders = 0
    picked_orders = 0
    number_of_pallets = 0
    keys_from_sheet = ["released_orders", "picked_orders", "number_of_pallets"]

    # --- شيت ARAMCO Outbound Report: Released Orders + Picked Orders ---
    outbound1_name = None
    for name in xls.sheet_names:
        if "ARAMCO Outbound Report" in (name or "").strip():
            outbound1_name = name
            break
    if not outbound1_name:
        for name in xls.sheet_names:
            if "outbound" in (name or "").lower() and "report" in (name or "").lower():
                outbound1_name = name
                break
    if not outbound1_name:
        for name in xls.sheet_names:
            if "outbound" in (name or "").lower():
                outbound1_name = name
                break

    if outbound1_name:
        try:
            raw1 = pd.read_excel(excel_path, sheet_name=outbound1_name, engine="openpyxl", header=None)
        except Exception:
            raw1 = None
        if raw1 is not None and not raw1.empty and raw1.shape[0] >= 2:
            first_col = str(raw1.iloc[0, 0]).strip() if raw1.shape[1] else ""
            need_header = (
                first_col.startswith("Unnamed:")
                or "ARAMCO" in first_col
                or "outbound report" in first_col.lower()
            )
            df1 = None
            if need_header:
                for idx in range(min(10, raw1.shape[0])):
                    row = raw1.iloc[idx]
                    cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
                    if ("order" in cells and "nbr" in cells) or ("order" in cells and "number" in cells):
                        headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw1.iloc[idx].values)]
                        df1 = raw1.iloc[idx + 1:].copy()
                        df1.columns = headers
                        df1 = df1.reset_index(drop=True)
                        break
            if df1 is None:
                df1 = pd.read_excel(excel_path, sheet_name=outbound1_name, engine="openpyxl", header=0)
            df1.columns = [str(c).strip() for c in df1.columns]

            order_nbr_col = _find_col(df1, ["Order Nbr", "Order Nbr.", "Order Number", "Order No", "Order #", "Order_ID", "Order ID"])
            status_col = _find_col(df1, ["Order Status", "Order_Status", "Status", "OrderStatus"])

            if order_nbr_col:
                order_series = df1[order_nbr_col].dropna().astype(str).str.strip()
                order_series = order_series[order_series != ""]
                released_orders = int(order_series.nunique())

                if status_col:
                    status_norm = df1[status_col].fillna("").astype(str).str.strip().str.lower()
                    shipped_mask = status_norm == "shipped"
                    if shipped_mask.any():
                        picked_orders = int(df1.loc[shipped_mask, order_nbr_col].dropna().astype(str).str.strip().nunique())
                    else:
                        picked_orders = 0
                else:
                    picked_orders = 0
        else:
            try:
                df1 = pd.read_excel(excel_path, sheet_name=outbound1_name, engine="openpyxl", header=0)
            except Exception:
                df1 = None
            if df1 is not None and not df1.empty:
                df1.columns = [str(c).strip() for c in df1.columns]
                order_nbr_col = _find_col(df1, ["Order Nbr", "Order Nbr.", "Order Number", "Order No", "Order #", "Order_ID", "Order ID"])
                status_col = _find_col(df1, ["Order Status", "Order_Status", "Status"])
                if order_nbr_col:
                    order_series = df1[order_nbr_col].dropna().astype(str).str.strip()
                    order_series = order_series[order_series != ""]
                    released_orders = int(order_series.nunique())
                    if status_col:
                        status_norm = df1[status_col].fillna("").astype(str).str.strip().str.lower()
                        shipped_mask = status_norm == "shipped"
                        if shipped_mask.any():
                            picked_orders = int(df1.loc[shipped_mask, order_nbr_col].dropna().astype(str).str.strip().nunique())

    # --- شيت Outbound2: Number of Pallets (LPNs) من عمود LPN Nbr (شيت مختلف عن ARAMCO Outbound Report) ---
    outbound2_name = None
    ob1_lower = (outbound1_name or "").strip().lower()
    for name in xls.sheet_names:
        n = (name or "").strip().lower()
        if n == "outbound2" and n != ob1_lower:
            outbound2_name = name
            break
    if not outbound2_name:
        for name in xls.sheet_names:
            n = (name or "").strip().lower()
            if "outbound" in n and "2" in n and n != ob1_lower:
                outbound2_name = name
                break

    if outbound2_name:
        try:
            df2 = pd.read_excel(excel_path, sheet_name=outbound2_name, engine="openpyxl", header=0)
        except Exception:
            df2 = None
        if df2 is not None and not df2.empty:
            df2.columns = [str(c).strip() for c in df2.columns]
            lpn_col = _find_col(df2, ["LPN Nbr", "LPN_Nbr", "LPN Nbr.", "LPN Number", "LPN No", "LPN"])
            if lpn_col:
                lpn_series = df2[lpn_col].dropna().astype(str).str.strip()
                lpn_series = lpn_series[lpn_series != ""]
                number_of_pallets = int(lpn_series.nunique())

    return {
        "outbound_kpi": {
            "released_orders": int(released_orders),
            "picked_orders": int(picked_orders),
            "number_of_pallets": int(number_of_pallets),
        },
        "outbound_kpi_keys_from_sheet": keys_from_sheet,
    }


def _read_pods_data_from_excel(excel_path):
    """
    يقرأ من شيت PODs_Data: عمود POD_Status (On Time, Pending, Late)،
    Delivery_Date للشهور، POD_ID للعدد. يرجّع داتا لشارت خط: كل شهر ونسبة كل حالة %.
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = str(name).strip().lower().replace(" ", "").replace("_", "")
        if "podsdata" in n or "pods_data" in n or (n == "pods" and "data" in n):
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "pod" in str(name).lower():
                sheet_name = name
                break
    if not sheet_name:
        return None
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
    except Exception:
        return None
    if df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns if c}

    def _col(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lower:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lower[col]
            if k.lower() in cols_lower:
                return cols_lower[k.lower()]
        return None

    status_col = _col("POD_Status", "POD Status", "PODStatus")
    date_col = _col("Delivery_Date", "Delivery Date", "DeliveryDate", "Date")
    pod_id_col = _col("POD_ID", "POD ID", "PODID")
    if not status_col or not date_col:
        return None
    if not pod_id_col:
        pod_id_col = df.columns[0]

    s = df[status_col].fillna("").astype(str).str.strip().str.lower()
    df["_status_norm"] = s.str.replace(r"\s+", " ", regex=True)
    valid_statuses = {"on time", "pending", "late"}
    df = df[df["_status_norm"].isin(valid_statuses)].copy()
    if df.empty:
        return None

    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["_date"])
    df["_month"] = df["_date"].dt.strftime("%b")
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    months_in_data = df["_month"].unique().tolist()
    months_sorted = sorted(months_in_data, key=lambda m: month_order.index(m) if m in month_order else 99)

    series_on_time = []
    series_pending = []
    series_late = []
    for m in months_sorted:
        grp = df[df["_month"] == m]
        on_time = (grp["_status_norm"] == "on time").sum()
        pending = (grp["_status_norm"] == "pending").sum()
        late = (grp["_status_norm"] == "late").sum()
        total = on_time + pending + late
        if total == 0:
            series_on_time.append(0)
            series_pending.append(0)
            series_late.append(0)
        else:
            series_on_time.append(round(100.0 * on_time / total, 1))
            series_pending.append(round(100.0 * pending / total, 1))
            series_late.append(round(100.0 * late / total, 1))

    # تجميع النسب الإجمالية للـ pod_status_breakdown (من نفس الشيت)
    total_on = (df["_status_norm"] == "on time").sum()
    total_pend = (df["_status_norm"] == "pending").sum()
    total_late = (df["_status_norm"] == "late").sum()
    total_all = total_on + total_pend + total_late
    if total_all > 0:
        pct_on = int(round(100.0 * total_on / total_all))
        pct_pend = int(round(100.0 * total_pend / total_all))
        pct_late = int(round(100.0 * total_late / total_all))
    else:
        pct_on = pct_pend = pct_late = 0
    pod_status_breakdown = [
        {"label": "On Time", "pct": pct_on, "color": "#7FB7A6"},
        {"label": "Pending", "pct": pct_pend, "color": "#A8C8EB"},
        {"label": "Late", "pct": pct_late, "color": "#E8A8A2"},
    ]

    return {
        "categories": months_sorted,
        "series": [
            {"name": "On Time", "data": series_on_time},
            {"name": "Pending", "data": series_pending},
            {"name": "Late", "data": series_late},
        ],
        "pod_status_breakdown": pod_status_breakdown,
    }


def _read_returns_data_from_excel(excel_path):
    """
    يقرأ من شيت Return (أو Returns):
    - Total SKUs: من عمود Shipment_nbr — عدد الشحنات المميزة (تحذف المتكرر) = count distinct.
    - Total LPNs: في الداشبورد يُأخذ من Inbound (number_of_pallets).
    إن وُجدت أعمدة Return_Status و Request_Date يُبنى returns_chart_data (On Time / Pending / Late).
    """
    if not excel_path or not os.path.exists(excel_path):
        return None
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = (str(name) or "").strip().lower()
        if n == "return" or n == "returns":
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            n = (str(name) or "").strip().lower().replace(" ", "").replace("_", "")
            if n == "returnsdata" or "returns_data" in n:
                sheet_name = name
                break
    if not sheet_name:
        for name in xls.sheet_names:
            if "return" in (str(name) or "").lower():
                sheet_name = name
                break
    if not sheet_name:
        return None

    # قراءة خام ثم كشف صف الرؤوس (قد يكون الشيت يبدأ بعنوان أو صفوف فارغة)
    try:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=None)
    except Exception:
        return None
    if raw.empty or raw.shape[0] < 1:
        return None

    # شكل الشيت: صف 1 عنوان (مثل ARAMCO Inbound Report)، صف 2 الرؤوس (Facility, Company, Shipment_nbr, LPN, ...)، من صف 3 البيانات
    df = None
    header_row_idx = None
    for idx in range(min(20, raw.shape[0])):
        row = raw.iloc[idx]
        cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
        cells_norm = cells.replace(" ", "").replace("_", "")
        # عمود Shipment_nbr قد يظهر كـ Shipment_nbr أو Shipment Nbr
        if "shipmentnbr" in cells_norm:
            header_row_idx = idx
            break
        if "shipment" in cells and "nbr" in cells_norm:
            header_row_idx = idx
            break
        if "facility" in cells_norm and ("shipment" in cells_norm or "lpn" in cells_norm):
            header_row_idx = idx
            break
        if "shipment" in cells and "lpn" in cells:
            header_row_idx = idx
            break
    if header_row_idx is not None:
        headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[header_row_idx].values)]
        df = raw.iloc[header_row_idx + 1:].copy()
        df.columns = headers
        df = df.reset_index(drop=True)
    if df is None:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
        except Exception:
            return None
    if df is None or df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns if c}

    def _col(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lower:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lower[col]
            if k.lower() in cols_lower:
                return cols_lower[k.lower()]
        return None

    def _col_contains(*parts):
        for col in df.columns:
            c_lower = str(col).lower().replace(" ", "").replace("_", "")
            if all(p.lower().replace(" ", "").replace("_", "") in c_lower for p in parts):
                return col
        return None

    def _col_contains_substring(sub):
        sub = sub.lower().replace(" ", "").replace("_", "")
        for col in df.columns:
            if sub in str(col).lower().replace(" ", "").replace("_", ""):
                return col
        return None

    # Total SKUs = عدد الشحنات المميزة من عمود Shipment_nbr (تحذف المتكرر)
    shipment_nbr_col = _col("Shipment_nbr", "Shipment Nbr", "ShipmentNbr")
    if not shipment_nbr_col:
        shipment_nbr_col = _col_contains("Shipment", "nbr")
    if not shipment_nbr_col:
        shipment_nbr_col = _col_contains_substring("shipment_nbr")
    lpn_col = _col("LPN", "LPN Nbr", "LPNNbr")

    def _distinct_count(series):
        s = series.astype(str).str.strip()
        s = s.replace("", np.nan).replace("nan", np.nan).dropna()
        return int(s.nunique())

    total_skus_kpi = 0
    total_lpns_kpi = 0
    if shipment_nbr_col:
        total_skus_kpi = _distinct_count(df[shipment_nbr_col])

    status_col = _col("Return_Status", "Return Status", "ReturnStatus")
    date_col = _col("Request_Date", "Request Date", "RequestDate", "Date")
    return_id_col = _col("Return_ID", "Return ID", "ReturnID")
    nbr_skus_col = _col("Nbr_SKUs", "Nbr SKUs", "NbrSKUs")
    nbr_items_col = _col("Nbr_Items", "Nbr Items", "NbrItems")

    # الشارت (On Time / Pending / Late) يحتاج status و date
    df_chart = df.copy()
    months_sorted = []
    if status_col and date_col:
        s = df_chart[status_col].fillna("").astype(str).str.strip().str.lower()
        df_chart["_status_norm"] = s.str.replace(r"\s+", " ", regex=True)
        valid_statuses = {"on time", "pending", "late"}
        df_chart = df_chart[df_chart["_status_norm"].isin(valid_statuses)].copy()
        if not df_chart.empty:
            df_chart["_date"] = pd.to_datetime(df_chart[date_col], errors="coerce")
            df_chart = df_chart.dropna(subset=["_date"])
            df_chart["_month"] = df_chart["_date"].dt.strftime("%b")
            month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            months_in_data = df_chart["_month"].unique().tolist()
            months_sorted = sorted(months_in_data, key=lambda m: month_order.index(m) if m in month_order else 99)
    else:
        df_chart = pd.DataFrame()

    # إذا لم يُوجد عمود Shipment_nbr نستخدم المنطق القديم لـ Total SKUs فقط (للتوافق مع شيتات قديمة)
    if not shipment_nbr_col and return_id_col:
        total_skus_kpi = _distinct_count(df[return_id_col])
    elif not shipment_nbr_col and nbr_skus_col:
        total_skus_kpi = int(pd.to_numeric(df[nbr_skus_col], errors="coerce").fillna(0).sum())

    # Total LPNs في الداشبورد يُعيّن من Inbound (number_of_pallets) في _get_dashboard_include_context

    series_on_time = []
    series_pending = []
    series_late = []
    if not df_chart.empty and months_sorted:
        for m in months_sorted:
            grp = df_chart[df_chart["_month"] == m]
            on_time = (grp["_status_norm"] == "on time").sum()
            pending = (grp["_status_norm"] == "pending").sum()
            late = (grp["_status_norm"] == "late").sum()
            total = on_time + pending + late
            if total == 0:
                series_on_time.append(0)
                series_pending.append(0)
                series_late.append(0)
            else:
                series_on_time.append(round(100.0 * on_time / total, 1))
                series_pending.append(round(100.0 * pending / total, 1))
                series_late.append(round(100.0 * late / total, 1))

    return {
        "returns_kpi": {
            "total_skus": total_skus_kpi,
            "total_lpns": total_lpns_kpi,
        },
        "returns_chart_data": {
            "categories": months_sorted,
            "series": [
                {"name": "On Time", "data": series_on_time},
                {"name": "Pending", "data": series_pending},
                {"name": "Late", "data": series_late},
            ],
        },
    }


def _read_inventory_data_from_excel(excel_path):
    """
    يقرأ من شيت Inventory (للداشبورد) ويملأ كاردات:
    - No of Location: Riyadh + Dammam + Jeddah (إما من عمود واحد + Region، أو من أعمدة منفصلة لكل منطقة).
    - Total Qty: نفس المنطق.
    الدالة اللي تحط الأرقام في الكارد: هذه + _get_dashboard_include_context يمرّر النتيجة للتمبلت.
    """
    if not excel_path or not os.path.exists(excel_path):
        return None
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = (name or "").strip().lower()
        if n == "inventory":
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "inventory" in (name or "").lower():
                sheet_name = name
                break
    if not sheet_name:
        return None
    try:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=None)
    except Exception:
        return None
    if raw.empty or raw.shape[0] < 1:
        return None

    # شكل الشيت: بلوكات — العمود A فيه اسم المنطقة (DAMMAM / Riyadh / Jeddah) ثم "No of location" ثم "Total Qty"، والقيم في العمود B
    # مثال: A1=DAMMAM, A2=No of location, B2=769, A3=Total Qty, B3=6958 | A5=Riyadh, A6=No of location, B6=535, ...
    def _parse_block_value(val):
        if pd.isna(val):
            return 0
        s = str(val).strip().replace(",", "")
        try:
            return int(float(s))
        except (ValueError, TypeError):
            return 0

    no_of_location_riyadh = 0
    no_of_location_dammam = 0
    no_of_location_jeddah = 0
    total_qty_riyadh = 0
    total_qty_dammam = 0
    total_qty_jeddah = 0
    total_hit = 0
    total_for_hit_pct = 0
    current_region = None
    nrows = raw.shape[0]
    col_a_idx = 0
    col_b_idx = 1 if raw.shape[1] > 1 else 0
    col_c_idx = 2 if raw.shape[1] > 2 else col_b_idx

    for i in range(nrows):
        a_val = raw.iloc[i, col_a_idx] if col_a_idx < raw.shape[1] else None
        b_val = raw.iloc[i, col_b_idx] if col_b_idx < raw.shape[1] else None
        c_val = raw.iloc[i, col_c_idx] if col_c_idx < raw.shape[1] else None
        a_str = (str(a_val).strip() if pd.notna(a_val) else "").lower()
        if not a_str:
            continue
        # عنوان منطقة
        if a_str == "dammam" or a_str == "damam":
            current_region = "Dammam"
            continue
        if a_str == "riyadh":
            current_region = "Riyadh"
            continue
        if a_str == "jeddah" or a_str == "jedd":
            current_region = "Jeddah"
            continue
        if current_region is None:
            continue
        # صف "No of location" — القيمة في B، وعمود Hit في C إن وُجد
        if "no of location" in a_str or "noof location" in a_str or (a_str.startswith("no ") and "location" in a_str):
            v = _parse_block_value(b_val)
            if current_region == "Riyadh":
                no_of_location_riyadh = v
            elif current_region == "Dammam":
                no_of_location_dammam = v
            else:
                no_of_location_jeddah = v
            total_for_hit_pct += v
            total_hit += _parse_block_value(c_val)
            continue
        # صف "Total Qty" — القيمة في B
        if "total qty" in a_str or "totalqty" in a_str or (a_str.startswith("total") and "qty" in a_str):
            v = _parse_block_value(b_val)
            if current_region == "Riyadh":
                total_qty_riyadh = v
            elif current_region == "Dammam":
                total_qty_dammam = v
            else:
                total_qty_jeddah = v
            continue

    total_no_loc = no_of_location_riyadh + no_of_location_dammam + no_of_location_jeddah
    total_qty = total_qty_riyadh + total_qty_dammam + total_qty_jeddah
    hit_pct = round(100.0 * total_hit / total_for_hit_pct, 2) if total_for_hit_pct > 0 else 0
    # لو لقينا على الأقل قيمة واحدة من البلوكات نرجع النتيجة مباشرة
    if total_no_loc > 0 or total_qty > 0:
        return {
            "inventory_kpi": {
                "no_of_location": total_no_loc,
                "no_of_location_riyadh": no_of_location_riyadh,
                "no_of_location_dammam": no_of_location_dammam,
                "no_of_location_jeddah": no_of_location_jeddah,
                "total_qty": total_qty,
                "total_qty_riyadh": total_qty_riyadh,
                "total_qty_dammam": total_qty_dammam,
                "total_qty_jeddah": total_qty_jeddah,
                "hit_pct": hit_pct,
                "total_skus": 0,
                "total_lpns": 0,
                "utilization_pct": "",
            },
        }

    # شكل جدول عادي: كشف صف الرؤوس
    df = None
    for idx in range(min(15, raw.shape[0])):
        row = raw.iloc[idx]
        cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
        has_no_loc = ("no of location" in cells or "noof location" in cells or
                      ("no" in cells and "location" in cells and "number" in cells))
        has_total_qty = ("total qty" in cells or "totalqty" in cells or
                         ("total" in cells and "qty" in cells))
        has_region = "region" in cells or "facility" in cells or "location" in cells or "area" in cells
        if has_no_loc or has_total_qty or has_region:
            headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[idx].values)]
            df = raw.iloc[idx + 1:].copy()
            df.columns = headers
            df = df.reset_index(drop=True)
            break
    if df is None:
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
        except Exception:
            return None
    if df is None or df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]

    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

    def _find_col(possible_names):
        norm_map = {_norm(c): c for c in df.columns}
        for name in possible_names:
            n = _norm(name)
            if n in norm_map:
                return norm_map[n]
        for col in df.columns:
            cn = _norm(col)
            if any(_norm(x) in cn for x in possible_names):
                return col
        return None

    def _find_col_containing(*parts):
        """يبحث عن عمود اسمه يحتوي كل الأجزاء (بعد التطبيع)."""
        for col in df.columns:
            cn = _norm(col)
            if all(_norm(p) in cn for p in parts):
                return col
        return None

    location_col = _find_col([
        "Region", "Facility", "Location", "Warehouse", "Site", "Area",
        "منطقة", "الموقع", "Facility Code", "Location Name", "KPI",
    ])
    no_of_location_col = _find_col([
        "No of location", "No of Location", "No Of Location",
        "Noof location", "No. of location", "Number of Location", "Number of Locations",
        "No of locations",
    ])
    total_qty_col = _find_col([
        "Total Qty", "Total_Qty", "Total Qty.", "Total Quantity", "TotalQuantity",
        "Total QTY",
    ])

    # أعمدة منفصلة لكل منطقة: No of Location Riyadh, No of Location Dammam, ...
    no_loc_riyadh_col = _find_col_containing("no", "location", "riyadh") or _find_col_containing("no", "location", "central")
    no_loc_dammam_col = _find_col_containing("no", "location", "dammam") or _find_col_containing("no", "location", "eastern")
    no_loc_jeddah_col = _find_col_containing("no", "location", "jeddah") or _find_col_containing("no", "location", "western")
    total_qty_riyadh_col = _find_col_containing("total", "qty", "riyadh") or _find_col_containing("total", "qty", "central")
    total_qty_dammam_col = _find_col_containing("total", "qty", "dammam") or _find_col_containing("total", "qty", "eastern")
    total_qty_jeddah_col = _find_col_containing("total", "qty", "jeddah") or _find_col_containing("total", "qty", "western")

    no_of_location_riyadh = 0
    no_of_location_dammam = 0
    no_of_location_jeddah = 0
    total_qty_riyadh = 0
    total_qty_dammam = 0
    total_qty_jeddah = 0

    def _norm_region(val):
        v = (str(val) or "").strip().lower()
        if not v:
            return None
        if "riyadh" in v or "central" in v or "ruh" in v or "الرياض" in v:
            return "Riyadh"
        if "dammam" in v or "eastern" in v or "damam" in v or "الدمام" in v:
            return "Dammam"
        if "jeddah" in v or "western" in v or "jedd" in v or "جدة" in v:
            return "Jeddah"
        return None

    # الطريقة 1: أعمدة منفصلة لكل منطقة (No of Location Riyadh, Total Qty Dammam, ...)
    if no_loc_riyadh_col or no_loc_dammam_col or no_loc_jeddah_col or total_qty_riyadh_col or total_qty_dammam_col or total_qty_jeddah_col:
        def _sum_col(col):
            if col is None:
                return 0
            return int(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
        if no_loc_riyadh_col:
            no_of_location_riyadh = _sum_col(no_loc_riyadh_col)
        if no_loc_dammam_col:
            no_of_location_dammam = _sum_col(no_loc_dammam_col)
        if no_loc_jeddah_col:
            no_of_location_jeddah = _sum_col(no_loc_jeddah_col)
        if total_qty_riyadh_col:
            total_qty_riyadh = _sum_col(total_qty_riyadh_col)
        if total_qty_dammam_col:
            total_qty_dammam = _sum_col(total_qty_dammam_col)
        if total_qty_jeddah_col:
            total_qty_jeddah = _sum_col(total_qty_jeddah_col)
    # الطريقة 2: عمود Region + عمود No of location + عمود Total Qty
    elif location_col and (no_of_location_col or total_qty_col):
        df["_region_norm"] = df[location_col].apply(_norm_region)
        df_three = df[df["_region_norm"].notna()].copy()
        if not df_three.empty:
            for region in ("Riyadh", "Dammam", "Jeddah"):
                sub = df_three[df_three["_region_norm"] == region]
                if not sub.empty:
                    if no_of_location_col:
                        val = int(pd.to_numeric(sub[no_of_location_col], errors="coerce").fillna(0).sum())
                        if region == "Riyadh":
                            no_of_location_riyadh = val
                        elif region == "Dammam":
                            no_of_location_dammam = val
                        else:
                            no_of_location_jeddah = val
                    if total_qty_col:
                        val = int(pd.to_numeric(sub[total_qty_col], errors="coerce").fillna(0).sum())
                        if region == "Riyadh":
                            total_qty_riyadh = val
                        elif region == "Dammam":
                            total_qty_dammam = val
                        else:
                            total_qty_jeddah = val
    else:
        # الطريقة 3: مفيش عمود منطقة — نجمع العمودين ككل (المجموع الإجمالي فقط)
        if no_of_location_col:
            tot = int(pd.to_numeric(df[no_of_location_col], errors="coerce").fillna(0).sum())
            no_of_location_riyadh = tot  # نعرض الإجمالي تحت "Riyadh" كبديل
        if total_qty_col:
            tot = int(pd.to_numeric(df[total_qty_col], errors="coerce").fillna(0).sum())
            total_qty_riyadh = tot

    no_of_location = no_of_location_riyadh + no_of_location_dammam + no_of_location_jeddah
    total_qty = total_qty_riyadh + total_qty_dammam + total_qty_jeddah

    return {
        "inventory_kpi": {
            "no_of_location": no_of_location,
            "no_of_location_riyadh": no_of_location_riyadh,
            "no_of_location_dammam": no_of_location_dammam,
            "no_of_location_jeddah": no_of_location_jeddah,
            "total_qty": total_qty,
            "total_qty_riyadh": total_qty_riyadh,
            "total_qty_dammam": total_qty_dammam,
            "total_qty_jeddah": total_qty_jeddah,
            "hit_pct": 0,
            "total_skus": 0,
            "total_lpns": 0,
            "utilization_pct": "",
        },
    }


def _read_inventory_snapshot_capacity_from_excel(excel_path):
    """
    يقرأ من شيت Inventory_Snapshot:
    - Used_Space_m3 → Used (مجموع ثم نسبة مئوية).
    - Available_Space_m3 → Available (مجموع ثم نسبة مئوية).
    يرجع inventory_capacity_data: { used: نسبة Used %, available: نسبة Available % }.
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = str(name).strip().lower().replace(" ", "").replace("_", "")
        if "inventorysnapshot" in n or "inventory_snapshot" in n:
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "inventory" in str(name).lower() and "snapshot" in str(name).lower():
                sheet_name = name
                break
    if not sheet_name:
        return None
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
    except Exception:
        return None
    if df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns if c}

    def _col(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lower:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lower[col]
            if k.lower() in cols_lower:
                return cols_lower[k.lower()]
        return None

    used_col = _col("Used_Space_m3", "Used Space m3", "UsedSpace_m3")
    avail_col = _col("Available_Space_m3", "Available Space m3", "AvailableSpace_m3")
    if not used_col or not avail_col:
        return None

    total_used = pd.to_numeric(df[used_col], errors="coerce").fillna(0).sum()
    total_avail = pd.to_numeric(df[avail_col], errors="coerce").fillna(0).sum()
    total = total_used + total_avail
    if total <= 0:
        return {"inventory_capacity_data": {"used": 0, "available": 0}}

    used_pct = round(100.0 * total_used / total, 1)
    available_pct = round(100.0 - used_pct, 1)
    return {
        "inventory_capacity_data": {
            "used": used_pct,
            "available": available_pct,
        },
    }


def _read_inventory_warehouse_table_from_excel(excel_path):
    """
    يقرأ من شيت Inventory_Snapshot جدول الـ Warehouse:
    - Warehouse من عمود Warehouse
    - SKUs من عمود Total_SKUs
    - Available Space من عمود Available_Space_m3
    - Utilization % من عمود Utilization_%
    كل صف كما هو من الشيت بدون جمع.
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = str(name).strip().lower().replace(" ", "").replace("_", "")
        if "inventorysnapshot" in n or "inventory_snapshot" in n:
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "inventory" in str(name).lower() and "snapshot" in str(name).lower():
                sheet_name = name
                break
    if not sheet_name:
        return None
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=0)
    except Exception:
        return None
    if df.empty or len(df) < 1:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns if c}

    def _col(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lower:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lower[col]
            if k.lower() in cols_lower:
                return cols_lower[k.lower()]
        return None

    warehouse_col = _col("Warehouse")
    total_skus_col = _col("Total_SKUs", "Total SKUs", "TotalSKUs")
    avail_space_col = _col("Available_Space_m3", "Available Space m3", "AvailableSpace_m3")
    util_col = _col("Utilization_%", "Utilization %", "UtilizationPct", "Utilization")
    if not warehouse_col:
        return None

    def _val(col, r):
        if not col or col not in r.index:
            return ""
        v = r[col]
        if pd.isna(v):
            return ""
        if isinstance(v, (int, float)):
            return str(int(v)) if v == int(v) else str(v)
        return str(v).strip()

    def _util_pct(col, r):
        if not col or col not in r.index:
            return ""
        v = r[col]
        if pd.isna(v):
            return ""
        try:
            num = float(v)
            if 0 <= num <= 1:
                return f"{round(num * 100, 2)}%"
            return f"{round(num, 2)}%"
        except (TypeError, ValueError):
            s = str(v).strip()
            return f"{s}%" if s and not s.endswith("%") else s

    rows = []
    for _, r in df.iterrows():
        warehouse = "" if pd.isna(r[warehouse_col]) else str(r[warehouse_col]).strip()
        rows.append({
            "warehouse": warehouse,
            "skus": _val(total_skus_col, r),
            "available_space": _val(avail_space_col, r),
            "utilization_pct": _util_pct(util_col, r),
        })
    if not rows:
        return None
    return {"inventory_warehouse_table": rows}


# سعات ثابتة للمناطق (Capacity الأساسي)
WAREHOUSE_CAPACITY = {"Jeddah": 1800, "Dammam": 1575, "Riyadh": 1125}


def _default_warehouse_and_capacity():
    """جدول Warehouse افتراضي + رسمة Capacity (0% Used، 100% Available) عند غياب أعمدة Facility/Location في شيت Dashboard."""
    warehouse_table = []
    total_capacity = 0
    for fac in ("Jeddah", "Dammam", "Riyadh"):
        capacity = WAREHOUSE_CAPACITY.get(fac, 0)
        total_capacity += capacity
        warehouse_table.append({
            "facility": fac,
            "capacity": capacity,
            "utilized": 0,
            "pending": 0,
            "empty": capacity,
            "percentage": 0,
        })
    used_pct = 0.0
    available_pct = 100.0 if total_capacity > 0 else 0.0
    return {
        "inventory_warehouse_table": warehouse_table,
        "inventory_capacity_data": {"used": used_pct, "available": available_pct},
    }


def _read_dashboard_warehouse_from_excel(excel_path):
    """
    يقرأ من شيت اسمه Dashboard:
    - يفلتر بـ Facility (المنطقة): Jeddah, Dammam, Riyadh
    - يفلتر بـ Create Date (حسب الأيام إن لزم)
    - عمود Location: يحذف المتكرر ويحسب العدد = Utilized لكل منطقة
    - Capacity رقم ثابت لكل منطقة، Empty = Capacity - Utilized، Percentage = Utilized/Capacity
    - يرجع: inventory_warehouse_table (الجدول) + inventory_capacity_data (Used/Available للنسبة في الرسمة)
    """
    if not excel_path or not os.path.exists(excel_path):
        return None
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None
    sheet_name = None
    for name in xls.sheet_names:
        n = (str(name) or "").strip().lower()
        if n == "dashboard":
            sheet_name = name
            break
    if not sheet_name:
        for name in xls.sheet_names:
            if "dashboard" in (str(name) or "").lower():
                sheet_name = name
                break
    if not sheet_name:
        return None

    # قراءة خام: الشيت قد يبدأ بعنوان (مثل ARAMCO Stock On Report) ثم صف الرؤوس
    try:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", header=None)
    except Exception:
        return None
    if raw.empty or raw.shape[0] < 2:
        return _default_warehouse_and_capacity()

    # كشف صف الرؤوس: أي صف فيه Facility أو Location أو Batch number أو Production date
    header_row_idx = None
    for idx in range(min(15, raw.shape[0])):
        row = raw.iloc[idx]
        cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
        cells_norm = cells.replace(" ", "").replace("_", "")
        if "facility" in cells_norm or "location" in cells_norm:
            header_row_idx = idx
            break
        if "batch" in cells_norm or "productiondate" in cells_norm or "createdate" in cells_norm:
            header_row_idx = idx
            break
    if header_row_idx is None:
        header_row_idx = 0
    headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[header_row_idx].values)]
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    if df.empty or len(df) < 1:
        return _default_warehouse_and_capacity()

    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in df.columns if c}

    def _col(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lower:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lower[col]
            if k.lower() in cols_lower:
                return cols_lower[k.lower()]
        return None

    def _col_contains(sub):
        sub = sub.lower().replace(" ", "").replace("_", "")
        for col in df.columns:
            if sub in str(col).lower().replace(" ", "").replace("_", ""):
                return col
        return None

    facility_col = _col("Facility", "Region", "Warehouse", "Site", "Area") or _col_contains("Facility") or _col_contains("Region")
    create_date_col = _col("Create Date", "CreateDate", "Create_Date", "Date") or _col_contains("Production date")
    location_col = _col("Location", "Location Name", "LocationName", "Location_Nbr") or _col_contains("Location") or _col("Batch number", "Batch number", "BatchNumber")
    if not facility_col or not location_col:
        return _default_warehouse_and_capacity()

    def _norm_facility(val):
        v = (str(val) or "").strip().lower()
        if not v:
            return None
        if "jeddah" in v or "western" in v or "jedd" in v:
            return "Jeddah"
        if "dammam" in v or "eastern" in v or "damam" in v:
            return "Dammam"
        if "riyadh" in v or "central" in v or "ruh" in v:
            return "Riyadh"
        return None

    df["_facility_norm"] = df[facility_col].apply(_norm_facility)
    df = df[df["_facility_norm"].notna()].copy()
    if df.empty:
        return _default_warehouse_and_capacity()

    # فلترة بـ Create Date: نأخذ كل التواريخ (أو آخر 30 يومًا إذا أردت — هنا نأخذ الكل)
    if create_date_col:
        df["_create_dt"] = pd.to_datetime(df[create_date_col], errors="coerce")
        df = df.dropna(subset=["_create_dt"])

    def _distinct_count(series):
        s = series.astype(str).str.strip().replace("", np.nan).dropna()
        return int(s.nunique())

    warehouse_table = []
    total_utilized = 0
    total_capacity = 0
    for fac in ("Jeddah", "Dammam", "Riyadh"):
        capacity = WAREHOUSE_CAPACITY.get(fac, 0)
        sub = df[df["_facility_norm"] == fac]
        utilized = _distinct_count(sub[location_col]) if not sub.empty else 0
        empty = max(0, capacity - utilized)
        pending = 0
        percentage = round(100.0 * utilized / capacity, 2) if capacity > 0 else 0
        warehouse_table.append({
            "facility": fac,
            "capacity": capacity,
            "utilized": utilized,
            "pending": pending,
            "empty": empty,
            "percentage": percentage,
        })
        total_utilized += utilized
        total_capacity += capacity

    total_empty = total_capacity - total_utilized
    used_pct = round(100.0 * total_utilized / total_capacity, 1) if total_capacity > 0 else 0
    available_pct = round(100.0 * total_empty / total_capacity, 1) if total_capacity > 0 else 0

    return {
        "inventory_warehouse_table": warehouse_table,
        "inventory_capacity_data": {
            "used": used_pct,
            "available": available_pct,
        },
    }


def _read_returns_region_table_from_excel(excel_path):
    """
    يبني returns_region_table من Inventory_Lots + Inventory_Snapshot:
    - Region من عمود Warehouse في Inventory_Lots (فلترة بالـ Warehouse).
    - SKUs: عدد القيم الفريدة لعمود SKU لكل Warehouse بعد الفلترة بالتاريخ.
    - Available: مجموع LPNs لكل Warehouse بعد الفلترة بـ Snapshot_Date (آخر تاريخ).
    - Utilization %: (LPNs للمنطقة والتاريخ) / Capacity_m3 من Inventory_Snapshot لنفس المنطقة، كنسبة مئوية.
    """
    try:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    except Exception:
        return None

    def _find_sheet(*names):
        for want in names:
            want_n = want.lower().replace(" ", "").replace("_", "")
            for s in xls.sheet_names:
                if want_n in str(s).lower().replace(" ", "").replace("_", ""):
                    return s
        return None

    lots_sheet = _find_sheet("Inventory_Lots", "Inventory Lots")
    snapshot_sheet = _find_sheet("Inventory_Snapshot", "Inventory Snapshot")
    if not lots_sheet:
        return None

    try:
        df_lots = pd.read_excel(excel_path, sheet_name=lots_sheet, engine="openpyxl", header=0)
    except Exception:
        return None
    if df_lots.empty:
        return None

    df_lots.columns = [str(c).strip() for c in df_lots.columns]
    cols_lots = {c.lower(): c for c in df_lots.columns if c}

    def _col_lots(*keys):
        for k in keys:
            k_norm = k.lower().replace(" ", "").replace("_", "")
            for col in cols_lots:
                if col.replace(" ", "").replace("_", "") == k_norm:
                    return cols_lots[col]
        return None

    wh_col = _col_lots("Warehouse")
    sku_col = _col_lots("SKU", "Sku")
    lpns_col = _col_lots("LPNs", "LPN")
    snap_col = _col_lots("Snapshot_Date", "Snapshot Date", "SnapshotDate", "Date")
    if not wh_col or not snap_col:
        return None
    if not lpns_col:
        lpns_col = df_lots.columns[1] if len(df_lots.columns) > 1 else None
    if not lpns_col:
        return None

    df_lots["_date"] = pd.to_datetime(df_lots[snap_col], errors="coerce")
    df_lots = df_lots.dropna(subset=["_date"])
    if df_lots.empty:
        return None

    latest_date = df_lots["_date"].max()
    df_filtered = df_lots[df_lots["_date"] == latest_date].copy()

    capacity_by_warehouse = {}
    if snapshot_sheet:
        try:
            df_snap = pd.read_excel(excel_path, sheet_name=snapshot_sheet, engine="openpyxl", header=0)
            if not df_snap.empty:
                df_snap.columns = [str(c).strip() for c in df_snap.columns]
                snap_cols = {c.lower(): c for c in df_snap.columns if c}
                snap_wh = next((snap_cols[c] for c in snap_cols if "warehouse" in c.replace(" ", "").replace("_", "")), None)
                cap_col = next((snap_cols[c] for c in snap_cols if "capacity_m3" in c.replace(" ", "").replace("_", "") or ("capacity" in c and "m3" in c)), None)
                if not cap_col:
                    used_c = next((snap_cols[c] for c in snap_cols if "used_space" in c.replace(" ", "").replace("_", "")), None)
                    avail_c = next((snap_cols[c] for c in snap_cols if "available_space" in c.replace(" ", "").replace("_", "")), None)
                    if used_c and avail_c:
                        df_snap["_cap"] = pd.to_numeric(df_snap[used_c], errors="coerce").fillna(0) + pd.to_numeric(df_snap[avail_c], errors="coerce").fillna(0)
                        cap_col = "_cap"
                if snap_wh and cap_col:
                    for _, r in df_snap.iterrows():
                        w = r.get(snap_wh)
                        if pd.isna(w):
                            continue
                        w = str(w).strip()
                        if not w:
                            continue
                        c = r.get(cap_col)
                        if cap_col == "_cap":
                            val = c
                        else:
                            val = pd.to_numeric(c, errors="coerce")
                        if pd.notna(val) and val > 0:
                            capacity_by_warehouse[w] = float(val)
        except Exception:
            pass

    df_filtered["_lpns_num"] = pd.to_numeric(df_filtered[lpns_col], errors="coerce").fillna(0)
    rows = []
    for wh, grp in df_filtered.groupby(wh_col, dropna=False):
        wh_name = "" if pd.isna(wh) else str(wh).strip()
        skus = grp[sku_col].dropna().astype(str).str.strip() if sku_col else pd.Series(dtype=object)
        skus = skus[skus != ""].nunique() if not skus.empty else 0
        available = int(grp["_lpns_num"].sum())
        cap = capacity_by_warehouse.get(wh_name) or capacity_by_warehouse.get(wh)
        if cap and cap > 0:
            util = round(100.0 * available / cap, 2)
            utilization_pct = f"{util}%"
        else:
            utilization_pct = "—"
        rows.append({
            "region": wh_name,
            "skus": str(int(skus)) if isinstance(skus, (int, float)) else str(skus),
            "available": str(available),
            "utilization_pct": utilization_pct,
        })

    if not rows:
        return None
    return {"returns_region_table": rows}


def get_dashboard_tab_context(request):
    """
    يبني سياق تاب الداشبورد (نفس بيانات Dashboard view).
    إذا وُجدت الفيو في تطبيق dashboard أو inbound يتم استخدامها، وإلا يُرجع سياق افتراضي.
    """
    try:
        for app_label in ["dashboard", "inbound"]:
            try:
                view_module = __import__(f"{app_label}.views", fromlist=["DashboardView"])
                ViewClass = getattr(view_module, "DashboardView", None)
                if ViewClass is not None:
                    view = ViewClass()
                    view.request = request
                    view.object = None
                    return view.get_context_data()
            except (ImportError, AttributeError):
                continue
    except Exception:
        pass
    # سياق افتراضي عند عدم وجود الموديلات/الفيو (مع داتا وهمية لـ Inbound)
    return {
        "title": "Dashboard",
        "breadcrumb": {"title": "Healthcare Dashboard", "parent": "Dashboard", "child": "Default"},
        "is_admin": False,
        "is_employee": False,
        "inbound_data": [],
        "transportation_outbound_data": [],
        "wh_outbound_data": [],
        "returns_data": [],
        "expiry_data": [],
        "damage_data": [],
        "inventory_data": [],
        "pallet_location_availability_data": [],
        "hse_data": [],
        "number_of_shipments": 0,
        "total_vehicles_daily": 0,
        "total_pallets": 0,
        "total_pending_shipments": 0,
        "total_number_of_shipments": 0,
        "total_quantity": 0,
        "total_number_of_line": 0,
        # Inbound KPI + داتا شارت Pending Shipments (من الديكت في الفيو)
        "inbound_kpi": INBOUND_DEFAULT_KPI.copy(),
        "pending_shipments": list(INBOUND_DEFAULT_PENDING_SHIPMENTS),
        "shipment_data": {"bulk": 0, "loose": 0, "cold": 0, "frozen": 0, "ambient": 0},
        "wh_total_released_order": 0,
        "wh_total_piked_order": 0,
        "wh_total_pending_pick_orders": 0,
        "wh_total_number_of_PODs_collected_on_time": 0,
        "wh_total_number_of_PODs_collected_Late": 0,
        "total_orders_items_returned": 0,
        "total_number_of_return_items_orders_updated_on_time": 0,
        "total_number_of_return_items_orders_updated_late": 0,
        "total_SKUs_expired": 0,
        "total_expired_SKUS_disposed": 0,
        "total_nearly_expired_1_to_3_months": 0,
        "total_nearly_expired_3_to_6_months": 0,
        "total_SKUs_expired_calculated": 0,
        "Total_QTYs_Damaged_by_WH": 0,
        "Total_Number_of_Damaged_during_receiving": 0,
        "Total_Araive_Damaged": 0,
        "Total_Locations_match": 0,
        "Total_Locations_not_match": 0,
        "last_shipment": None,
        "Total_Storage_Pallet": 0,
        "Total_Storage_pallet_empty": 0,
        "Total_Storage_Bin": 0,
        "Total_occupied_pallet_location": 0,
        "Total_Storage_Bin_empty": 0,
        "Total_occupied_Bin_location": 0,
        "Total_Incidents_on_the_side": 0,
        "total_no_of_employees": 0,
        "admin_data": [],
        "user_type": "Unknown",
        "years": [],
        "months": list(calendar_module.month_name)[1:],
        "days": list(range(1, 32)),
        "returns_region_table": [
            {"region": "Main warehouse", "skus": "2,538", "available": "1118", "utilization_pct": "71%"},
            {"region": "Dammam DC", "skus": "501", "available": "200", "utilization_pct": "—"},
            {"region": "Riyadh DC", "skus": "3,996", "available": "209", "utilization_pct": "—"},
            {"region": "Jeddah DC", "skus": "7,996", "available": "300", "utilization_pct": "—"},
        ],
    }


@method_decorator(csrf_exempt, name="dispatch")
class UploadExcelViewRoche(View):
    template_name = "index.html"
    excel_file_name = "all sheet.xlsm"
    correct_code = "1234"

    # تابات تحذف من الداشبورد (أضف أسماء الشيتات كما هي في الإكسل)
    EXCLUDE_TABS = []  # مثال: ["Sheet2", "تقارير قديمة", "Backup"]
    # أو: اعرض تابات معينة فقط (لو ضعت قائمة هنا، التابات الأخرى كلها تختفي)
    INCLUDE_ONLY_TABS = (
        None  # مثال: ["Overview", "Dock to stock", "Order General Information"]
    )
    # تابات افتراضية نعرضها بدون الاعتماد على شيت مباشر
    DASHBOARD_TAB_NAME = "Dashboard"
    DEFAULT_EXCEL_FILENAMES = [
        "all sheet.xlsm",
        "all sheet.xlsx",
        "all_sheet.xlsm",
        "all_sheet.xlsx",
    ]

    MONTH_LOOKUP = {}
    MONTH_PREFIXES = set()
    for idx in range(1, 13):
        abbr = month_abbr[idx]
        full = month_name[idx]
        if abbr:
            MONTH_LOOKUP[abbr.lower()] = abbr
            MONTH_PREFIXES.add(abbr.lower())
        if full:
            MONTH_LOOKUP[full.lower()] = abbr
        MONTH_LOOKUP[str(idx)] = abbr
        MONTH_LOOKUP[f"{idx:02d}"] = abbr
    MONTH_LOOKUP["sept"] = "Sep"

    AGGREGATE_COLUMN_KEYWORDS = {
        "total",
        "grand total",
        "overall total",
        "sum",
        "ytd",
        "y.t.d.",
        "avg",
        "average",
        "target",
        "target (%)",
        "target %",
        "target%",
        "cumulative",
    }

    # اسم الملف الافتراضي إذا وُضع في excel_uploads بدون رفع (مثلاً all sheet.xlsm)
    def get_excel_path(self):
        folder_path = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
        os.makedirs(folder_path, exist_ok=True)
        # الملف الرئيسي للتابات: all_sheet.xlsx (مع بعض البدائل في حالة الـ xlsm أو المسافة)
        priority_files = [
            "all_sheet.xlsx",
            "all_sheet.xlsm",
            "all sheet.xlsx",
            "all sheet.xlsm",
            "latest.xlsm",
            "latest.xlsx",
        ] + self.DEFAULT_EXCEL_FILENAMES
        for name in priority_files:
            path = os.path.join(folder_path, name)
            if os.path.exists(path):
                return path
        return os.path.join(folder_path, "latest.xlsx")

    def get_uploaded_file_path(self, request):
        folder = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
        os.makedirs(folder, exist_ok=True)

        # أولوية: ملف الجلسة ثم latest.xlsm ثم latest.xlsx ثم all sheet
        if request:
            saved_path = request.session.get("uploaded_excel_path")
            if saved_path and os.path.exists(saved_path):
                return saved_path
        priority_files = ["latest.xlsm", "latest.xlsx"] + self.DEFAULT_EXCEL_FILENAMES
        for name in priority_files:
            path = os.path.join(folder, name)
            if os.path.exists(path):
                if request:
                    try:
                        request.session["uploaded_excel_path"] = path
                        request.session.save()
                    except Exception:
                        pass
                return path
        return os.path.join(folder, "latest.xlsx")

    def get_sheet_dataframe(self, request, sheet_name):
        """
        يقرأ بيانات الشيت من الكاش (الداتابيز JSON) إن وُجدت للملف الحالي،
        وإلا يقرأ من ملف الإكسل ويحدّث الكاش. يُرجع DataFrame أو None.
        """
        excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
        if not excel_path or not os.path.exists(excel_path):
            return None
        try:
            excel_path_norm = os.path.normpath(os.path.abspath(excel_path))
            cached = ExcelSheetCache.objects.filter(sheet_name=sheet_name).first()
            cached_path = (os.path.normpath(os.path.abspath(cached.source_file_path or "")) if cached and cached.source_file_path else "")
            try:
                path_match = (cached_path == excel_path_norm or
                              (cached_path and os.path.exists(cached_path) and os.path.realpath(excel_path_norm) == os.path.realpath(cached_path)))
            except OSError:
                path_match = (cached_path == excel_path_norm)
            if cached and path_match and cached.data is not None:
                print("✅ الشيت قُرئ من الكاش (الداتابيز) — الموقع صار يفتح بسرعه", flush=True)
                return pd.DataFrame(cached.data)
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            cache_rows = _dataframe_to_cache_rows(df)
            ExcelSheetCache.objects.update_or_create(
                sheet_name=sheet_name,
                defaults={"data": cache_rows, "source_file_path": excel_path_norm},
            )
            return df
        except Exception as e:
            print(f"⚠️ get_sheet_dataframe({sheet_name!r}): {e}")
            try:
                return pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            except Exception:
                return None

    @staticmethod
    def safe_format_value(val):
        if pd.isna(val) or val is pd.NaT:
            return ""
        elif isinstance(val, pd.Timestamp):
            if val.tzinfo is not None:
                val = val.tz_convert(None)
            return val.strftime("%Y-%m-%d %H:%M:%S")
        return val

    # ----------------------------------------------------
    # 🔧 Helper methods for month normalization & filtering
    # ----------------------------------------------------
    def normalize_month_label(self, month_value):
        if month_value is None:
            return None

        raw = str(month_value).strip()
        if not raw:
            return None

        lower = raw.lower()
        if lower in self.MONTH_LOOKUP:
            return self.MONTH_LOOKUP[lower]

        first_three = lower[:3]
        if first_three in self.MONTH_LOOKUP:
            return self.MONTH_LOOKUP[first_three]

        try:
            parsed = pd.to_datetime(raw, errors="coerce")
            if not pd.isna(parsed):
                return parsed.strftime("%b")
        except Exception:
            pass

        return raw[:3].capitalize()

    def _value_matches_month(self, value, month_lower):
        if value is None:
            return False
        normalized = self.normalize_month_label(value)
        return normalized is not None and normalized.lower() == month_lower

    def _column_matches_month(self, column, month_lower):
        if column is None:
            return False
        col_lower = str(column).strip().lower()
        if col_lower == month_lower:
            return True
        if col_lower.startswith(month_lower + " "):
            return True
        if col_lower.endswith(" " + month_lower):
            return True
        if col_lower.startswith(month_lower + "-") or col_lower.endswith(
            "-" + month_lower
        ):
            return True
        if col_lower.startswith(month_lower + "/") or col_lower.endswith(
            "/" + month_lower
        ):
            return True
        if col_lower.startswith(month_lower + "("):
            return True
        if col_lower.split(" ")[0] == month_lower:
            return True
        if col_lower.replace(".", "").startswith(month_lower):
            return True
        return False

    def _is_month_column(self, column):
        if column is None:
            return False
        col_lower = str(column).strip().lower()
        if col_lower in self.MONTH_LOOKUP:
            return True
        first_three = col_lower[:3]
        if first_three in self.MONTH_PREFIXES:
            return True
        col_split = col_lower.replace("/", " ").replace("-", " ").split()
        if col_split and col_split[0][:3] in self.MONTH_PREFIXES:
            return True
        return False

    def _is_aggregate_column(self, column):
        if column is None:
            return False
        col_lower = str(column).strip().lower()
        if col_lower in self.AGGREGATE_COLUMN_KEYWORDS:
            return True
        compact = col_lower.replace(" ", "")
        if compact in {"target%", "target(%)", "total%"}:
            return True
        if col_lower.isdigit():
            try:
                if int(col_lower) >= 1900:
                    return True
            except ValueError:
                pass
        return False

    def _append_missing_month_messages(self, tab_data, missing_months):
        if not missing_months:
            return

        message_table = {
            "title": "Missing Months",
            "columns": ["Message"],
            "data": [
                {"Message": f"No data available for month {month}."}
                for month in missing_months
            ],
        }

        if isinstance(tab_data.get("sub_tables"), list):
            tab_data["sub_tables"] = [
                sub
                for sub in tab_data["sub_tables"]
                if sub.get("title") != "Missing Months"
            ]
            tab_data["sub_tables"].append(message_table)
            return

        # في حال كان التاب عبارة عن جدول واحد فقط، نحوله إلى sub_tables
        columns = tab_data.pop("columns", None)
        data_rows = tab_data.pop("data", None)
        if columns is not None and data_rows is not None:
            existing_table = {
                "title": tab_data.get("name", "Data"),
                "columns": columns,
                "data": data_rows,
            }
            tab_data["sub_tables"] = [existing_table, message_table]
        else:
            tab_data["sub_tables"] = [message_table]

    def apply_month_filter_to_tab(
        self, tab_data, selected_month=None, selected_months=None
    ):
        if not tab_data:
            return None

        selected_months_norm = []
        if selected_months:
            if isinstance(selected_months, str):
                selected_months = [selected_months]
            seen = set()
            for month in selected_months:
                norm = self.normalize_month_label(month)
                if norm and norm.lower() not in seen:
                    seen.add(norm.lower())
                    selected_months_norm.append(norm)

        month_norm = self.normalize_month_label(selected_month)
        month_filters = []
        if selected_months_norm:
            month_filters = selected_months_norm
        elif month_norm:
            month_filters = [month_norm]
        else:
            tab_data.pop("selected_month", None)
            tab_data.pop("selected_months", None)
            return None

        month_filters_lower = [m.lower() for m in month_filters]
        matched_months = set()

        def matches_any_month(column):
            if not month_filters_lower:
                return False
            for month_lower in month_filters_lower:
                if self._column_matches_month(column, month_lower):
                    matched_months.add(month_lower)
                    return True
            return False

        def value_matches_month(value):
            if not month_filters_lower:
                return False
            normalized = self.normalize_month_label(value)
            if not normalized:
                return False
            val_lower = normalized.lower()
            if val_lower in month_filters_lower:
                matched_months.add(val_lower)
                return True
            return False

        def filter_columns(columns):
            filtered = []
            for col in columns:
                if self._is_month_column(col):
                    if matches_any_month(col):
                        filtered.append(col)
                elif self._is_aggregate_column(col) and not self._column_matches_month(
                    col,
                    month_filters_lower[0] if month_filters_lower else "",
                ):
                    continue
                else:
                    filtered.append(col)
            return filtered if filtered else columns

        def filter_rows(data_rows, columns):
            if not data_rows:
                return data_rows

            month_cols = [
                col
                for col in columns
                if str(col).strip().lower() in {"month", "month name", "monthname"}
            ]
            if not month_cols:
                return data_rows

            month_col = month_cols[0]
            scoped_rows = []
            for row in data_rows:
                value = None
                if isinstance(row, dict):
                    value = row.get(month_col)
                if value_matches_month(value):
                    scoped_rows.append(row)
            return scoped_rows if scoped_rows else data_rows

        if "sub_tables" in tab_data and isinstance(tab_data["sub_tables"], list):
            for sub in tab_data["sub_tables"]:
                if not isinstance(sub, dict):
                    continue
                # ✅ الحفاظ على chart_data في sub_table
                sub_chart_data = sub.get("chart_data", [])

                columns = sub.get("columns", [])
                if columns:
                    filtered_columns = filter_columns(columns)
                    if sub.get("data"):
                        new_data = []
                        for row in sub["data"]:
                            if isinstance(row, dict):
                                new_row = {
                                    col: row.get(col, "") for col in filtered_columns
                                }
                            else:
                                new_row = row
                            new_data.append(new_row)
                        sub["data"] = filter_rows(new_data, filtered_columns)
                    sub["columns"] = filtered_columns

                # ✅ إعادة إضافة chart_data إلى sub_table بعد التعديل (حتى لو كانت فارغة)
                sub["chart_data"] = sub_chart_data
        else:
            columns = tab_data.get("columns", [])
            data_rows = tab_data.get("data", [])
            if columns:
                filtered_columns = filter_columns(columns)
                if data_rows:
                    new_rows = []
                    for row in data_rows:
                        if isinstance(row, dict):
                            new_row = {
                                col: row.get(col, "") for col in filtered_columns
                            }
                        else:
                            new_row = row
                        new_rows.append(new_row)
                    tab_data["data"] = filter_rows(new_rows, filtered_columns)
                tab_data["columns"] = filtered_columns

        if "chart_data" in tab_data and isinstance(tab_data["chart_data"], list):
            for chart in tab_data["chart_data"]:
                if not isinstance(chart, dict):
                    continue
                points = chart.get("dataPoints")
                if not points:
                    continue
                filtered_points = []
                for point in points:
                    label_norm = self.normalize_month_label(point.get("label"))
                    if label_norm and label_norm.lower() in month_filters_lower:
                        matched_months.add(label_norm.lower())
                        filtered_points.append(point)
                if filtered_points:
                    chart["dataPoints"] = filtered_points

        if selected_months_norm:
            tab_data["selected_months"] = selected_months_norm
            return selected_months_norm[0]
        else:
            tab_data["selected_month"] = month_filters[0]
            return month_filters[0]

    @method_decorator(cache_control(max_age=3600, public=True), name="get")
    def get(self, request):
        print("🟢 [GET] Loading main dashboard with Overview/All-in-One tabs")
        cache.clear()  # Clear cache on each load

        # مسح بيانات الجلسة أولاً لو الطلب clear_excel (حتى تظهر رسالة رفع الملف)
        action_param = request.GET.get("action", "").strip().lower()
        if action_param == "clear_excel":
            request.session.pop("uploaded_excel_path", None)
            request.session.pop("dashboard_excel_path", None)
            try:
                request.session.save()
            except Exception:
                pass
            from django.shortcuts import redirect
            return redirect(request.path or "/")

        # --------------------------
        # Resolve Excel path — نفتح الصفحة عادي لو في ملف (جلسة أو مجلد)، بدون إجبار على صفحة الرفع
        # --------------------------
        excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
        data_is_uploaded = bool(excel_path and os.path.exists(excel_path))
        if not data_is_uploaded:
            form = ExcelUploadForm()
            return render(
                request, self.template_name, {"form": form, "data_is_uploaded": False}
            )

        # --------------------------
        # Read request parameters
        # --------------------------
        selected_tab = request.GET.get("tab", "").lower() or "all"
        selected_month = request.GET.get("month", "").strip()
        selected_quarter = request.GET.get("quarter", "").strip()
        action = request.GET.get("action", "").lower()
        status = request.GET.get("status")

        print(f"🔹 Selected tab: {selected_tab}")
        print(f"🔹 Selected month: {selected_month}")
        print(f"🔹 Selected quarter: {selected_quarter}")
        print(f"🔹 Action: {action}")

        print("🛰️ Quarter AJAX Triggered:", request.GET.get("quarter"))

        quarter_months = []
        quarter_error = None
        if selected_quarter:
            try:
                quarter_months = self._resolve_quarter_months(selected_quarter)
            except ValueError as exc:
                quarter_error = str(exc)

        effective_month = None if quarter_months else selected_month

        if action == "meeting_points_tab":
            return self.meeting_points_tab(request)

        # ✅ إذا كان الطلب AJAX وبه status فقط (بدون tab)، نعيد قسم Meeting Points فقط
        if (
            request.headers.get("X-Requested-With") == "XMLHttpRequest"
            and request.GET.get("status")
            and not request.GET.get("tab")
        ):
            meeting_html = self.get_meeting_points_section_html(
                request, request.GET.get("status", "all")
            )
            return JsonResponse({"meeting_section_html": meeting_html}, safe=False)

        if action == "export_excel":
            if quarter_error:
                return HttpResponse(quarter_error, status=400)
            return self.export_dashboard_excel(
                request,
                selected_month=effective_month,
                selected_months=quarter_months or None,
            )

        # ====================== طلبات AJAX ======================
        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            print("⚡ [AJAX Request] Received request")

            if quarter_error:
                return JsonResponse({"error": quarter_error})

            tab_filter_map = {
                "dashboard": lambda: self.dashboard_tab(request),
                "all": lambda: self.filter_all_tabs(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "return & refusal": lambda: self.filter_inbound(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                    tab_name="Return & Refusal",
                ),
                "rejections": lambda: self.filter_rejections_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "inbound": lambda: self.filter_dock_to_stock_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "outbound": lambda: self.filter_total_lead_time_performance(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "total lead time performance": lambda: self.filter_total_lead_time_performance(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "pods update": lambda: self.filter_pods_update(
                    request, effective_month
                ),
                "meeting points": lambda: self.meeting_points_tab(request),
                "inventory": lambda: self.filter_inventory(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "Inventory": lambda: self.filter_inventory(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
            }

            # كاش استجابة التاب على السيرفر (تسريع التنقل عند العودة لنفس التاب/الشهر/الربع)
            def _tab_response_cache_key():
                excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
                norm = os.path.normpath(os.path.abspath(excel_path or ""))
                mtime = _get_file_mtime(excel_path) or 0
                q = ",".join(sorted(quarter_months)) if quarter_months else ""
                return f"tab_resp:{norm}:{mtime}:{selected_tab}:{effective_month or ''}:{q}"

            for key, func in tab_filter_map.items():
                if key in selected_tab:
                    print(f"📂 Executing tab filter: {key}")
                    try:
                        cache_key = _tab_response_cache_key()
                        try:
                            cached = cache.get(cache_key)
                            if cached is not None:
                                return JsonResponse(cached, safe=False)
                        except Exception:
                            pass
                        result = func()

                        if isinstance(result, HttpResponse):
                            print(
                                "ℹ️ Filter returned HttpResponse/JsonResponse; returning as-is."
                            )
                            return result

                        payload = None
                        if isinstance(result, (dict, list)):
                            payload = result
                        elif isinstance(result, str):
                            payload = {"detail_html": result}
                        else:
                            payload = {"detail_html": str(result)}
                        try:
                            cache.set(cache_key, payload, timeout=86400)
                        except Exception:
                            pass
                        return JsonResponse(payload, safe=False)

                    except Exception as e:
                        import traceback

                        print("❌ Error while executing tab filter:", key)
                        traceback.print_exc()
                        return JsonResponse(
                            {"error": f"Error in '{key}': {str(e)}"},
                            status=200,
                        )

            # All-in-One (never cached)
            if selected_tab == "all":
                print("🔹 Loading All-in-One tab")
                all_result = self.filter_all_tabs(
                    request=request,
                    selected_month=effective_month,
                    selected_months=quarter_months or None,
                )
                return JsonResponse(all_result, safe=False)

            # Remaining tabs
            if selected_tab in ["rejections", "return & refusal"]:
                return JsonResponse(
                    self.filter_rejections_combined(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            # airport / seaport tabs تم إلغاؤها
            elif selected_tab in [
                "outbound",
                "total lead time performance",
                "total lead time preformance",
            ]:
                return JsonResponse(
                    self.filter_total_lead_time_performance(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            elif selected_tab == "total lead time preformance -r":
                return JsonResponse(
                    self.filter_total_lead_time_roche(request, effective_month),
                    safe=False,
                )
            # data logger tab تم إلغاؤه
            elif "dock to stock - roche" in selected_tab:
                return JsonResponse(
                    self.filter_dock_to_stock_roche(request, effective_month),
                    safe=False,
                )
            elif (selected_tab or "").lower() == "inbound":
                return JsonResponse(
                    self.filter_dock_to_stock_combined(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            elif selected_tab == "pods update":
                return JsonResponse(
                    self.filter_pods_update(request, effective_month), safe=True
                )
            elif "rejection" in selected_tab:
                return JsonResponse(
                    self.filter_rejection_data(request, effective_month), safe=False
                )
            elif "dock to stock" in selected_tab:
                return JsonResponse(
                    self.filter_dock_to_stock_combined(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            elif "meeting points" in selected_tab:
                return self.meeting_points_tab(request)
            elif selected_tab:
                raw_data = self.render_raw_sheet(request, selected_tab)
                return JsonResponse(raw_data, safe=False)
            else:
                return JsonResponse({"error": "⚠️ Please select a tab first."})

        # ====================== الطلب العادي ======================
        # كاش قائمة التابات + الشهور لتسريع فتح الصفحة (بدون فتح الإكسل عند كل زيارة)
        excel_path_norm = os.path.normpath(os.path.abspath(excel_path or ""))
        file_mtime = _get_file_mtime(excel_path) or 0
        initial_cache_key = f"initial_tabs_months:{excel_path_norm}:{file_mtime}"
        try:
            cached_initial = cache.get(initial_cache_key)
            if cached_initial:
                excel_tabs = cached_initial.get("excel_tabs", [])
                all_months = cached_initial.get("all_months", [])
            else:
                cached_initial = None
        except Exception:
            cached_initial = None
        if not cached_initial:
            try:
                xls = pd.ExcelFile(excel_path, engine="openpyxl")
                all_sheets = [s.strip() for s in xls.sheet_names]

                MERGE_SHEETS = ["Urgent orders details", "Outbound details"]
                REJECTION_SHEETS = ["Rejection", "Rejection breakdown"]
                AIRPORT_SHEETS = ["Airport Clearance - Roche", "Airport Clearance - 3PL"]
                SEAPORT_SHEETS = ["Seaport clearance - 3pl", "Seaport clearance - Roche"]
                TOTAL_LEADTIME_SHEETS = [
                    "Total lead time preformance",
                    "Total lead time preformance -R",
                ]
                DOCK_TO_STOCK_SHEETS = ["Dock to stock", "Dock to stock - Roche"]
                EXCLUDE_SHEETS_BASE = ["Sheet2"]
                EXCLUDE_SHEETS_EXTRA = getattr(
                    self.__class__, "EXCLUDE_TABS", []
                )
                EXCLUDE_SHEETS = list(EXCLUDE_SHEETS_BASE) + list(EXCLUDE_SHEETS_EXTRA)

                include_only = getattr(self.__class__, "INCLUDE_ONLY_TABS", None)
                if include_only:
                    include_set = {s.strip() for s in include_only}
                    filtered_tabs = [t for t in all_sheets if t in include_set]
                else:
                    filtered_tabs = [
                        t
                        for t in all_sheets
                        if t not in MERGE_SHEETS
                        and t not in REJECTION_SHEETS
                        and t not in AIRPORT_SHEETS
                        and t not in SEAPORT_SHEETS
                        and t not in TOTAL_LEADTIME_SHEETS
                        and t not in DOCK_TO_STOCK_SHEETS
                        and t not in EXCLUDE_SHEETS
                    ]

                virtual_tabs = [
                    self.DASHBOARD_TAB_NAME,
                    "Inbound",
                    "Outbound",
                    "Return & Refusal",
                    "PODs update",
                    "Inventory",
                    "Meeting Points & Action",
                ]
                if include_only:
                    include_set_v = {s.strip() for s in include_only}
                    filtered_tabs += [v for v in virtual_tabs if v in include_set_v]
                else:
                    filtered_tabs += virtual_tabs

                ordered_tabs = [
                    self.DASHBOARD_TAB_NAME,
                    "Inbound",
                    "Outbound",
                    "Return & Refusal",
                    "PODs update",
                    "Inventory",
                    "Meeting Points & Action",
                ]

                filtered_tabs = [tab for tab in ordered_tabs if tab in filtered_tabs]
                excel_tabs = [{"original": name, "display": name} for name in filtered_tabs]

                all_months_set = set()
                for sheet in xls.sheet_names:
                    try:
                        df = pd.read_excel(excel_path, sheet_name=sheet, engine="openpyxl")
                        df.columns = df.columns.str.strip().str.title()
                        possible_date_cols = [
                            c
                            for c in df.columns
                            if "date" in c.lower() or "month" in c.lower()
                        ]
                        if not possible_date_cols:
                            continue
                        col = possible_date_cols[0]
                        df[col] = pd.to_datetime(df[col], errors="coerce")
                        df["MonthName"] = df[col].dt.strftime("%b")
                        all_months_set.update(df["MonthName"].dropna().unique().tolist())
                    except Exception:
                        continue

                all_months = sorted(
                    all_months_set, key=lambda m: pd.to_datetime(m, format="%b")
                )
                print("📅 [INFO] الشهور المستخرجة من كل الشيتات:", all_months)
                try:
                    cache.set(
                        initial_cache_key,
                        {"excel_tabs": excel_tabs, "all_months": all_months},
                        timeout=86400,
                    )
                except Exception:
                    pass
            except Exception as e:
                print(f"⚠️ [ERROR] تعذر قراءة الشيتات من الملف: {e}")
                excel_tabs = []
                all_months = []

        meeting_points = MeetingPoint.objects.all().order_by("is_done", "-created_at")
        done_count = meeting_points.filter(is_done=True).count()
        total_count = meeting_points.count()

        all_tab_data = self.filter_all_tabs(
            request=request, selected_month=selected_month or None
        )

        render_context = {
            "data_is_uploaded": True,
            "months": all_months,
            "excel_tabs": excel_tabs,
            "active_tab": selected_tab or "all",
            "tab_summaries": [],
            "form": ExcelUploadForm(),
            "meeting_points": meeting_points,
            "done_count": done_count,
            "total_count": total_count,
            "all_tab_data": all_tab_data,
            "raw_tab_data": None,
        }
        try:
            dashboard_ctx = self._get_dashboard_include_context(request)
            render_context["dashboard_missing_data"] = dashboard_ctx.get("dashboard_missing_data", [])
            if (selected_tab or "").lower() == "dashboard":
                render_context.update(dashboard_ctx)
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"⚠️ [Dashboard include context] {e}")
            render_context.setdefault("dashboard_missing_data", [])

        return render(request, self.template_name, render_context)

    def post(self, request):
        print("📥 [DEBUG] تم استدعاء post()")  # ✅ بداية الدالة

        entered_code = request.POST.get("upload_code", "").strip()
        print(f"🔑 [DEBUG] الكود المدخل: {entered_code}")

        # ✅ التحقق من الكود
        if entered_code != self.correct_code:
            print("❌ [DEBUG] الكود غير صحيح!")
            if request.headers.get("x-requested-with") == "XMLHttpRequest":
                return JsonResponse(
                    {"error": "❌ Invalid code. Please try again."}, status=403
                )
            messages.error(request, "❌ Invalid code. Please try again.")
            return redirect(request.path)

        # ✅ التحقق من الملف المرفوع
        form = ExcelUploadForm(request.POST, request.FILES)
        if not form.is_valid():
            print("⚠️ [DEBUG] النموذج غير صالح أو لم يتم رفع ملف.")
            if request.headers.get("x-requested-with") == "XMLHttpRequest":
                return JsonResponse(
                    {"error": "⚠️ Please select an Excel file."}, status=400
                )
            return render(
                request, self.template_name, {"form": form, "data_is_uploaded": False}
            )

        # ✅ حفظ الملف (يدعم .xlsx و .xlsm مثل all sheet.xlsm)
        excel_file = form.cleaned_data["excel_file"]
        folder_path = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
        os.makedirs(folder_path, exist_ok=True)
        file_name = getattr(excel_file, "name", "") or ""
        is_dashboard_file = _is_dashboard_excel_filename(file_name)

        if is_dashboard_file:
            # ✅ ملف الداشبورد (شيت تاني): Aramco_Tamer3PL_KPI_Dashboard.xlsx — للتاب Dashboard فقط
            file_path = os.path.join(folder_path, DASHBOARD_EXCEL_FILENAME)
            print(f"📊 [DEBUG] رفع ملف الداشبورد: {file_name} → {file_path}")
        else:
            # ✅ الملف الرئيسي (all_sheet / latest) — لباقي التابات
            ext = os.path.splitext(file_name)[1] or ".xlsx"
            if ext.lower() not in (".xlsx", ".xlsm"):
                ext = ".xlsx"
            file_path = os.path.join(folder_path, "latest" + ext)

        try:
            if not is_dashboard_file:
                # ✅ حذف أي ملف latest قديم (xlsx أو xlsm) لتفادي بقاء ملف بالامتداد الآخر
                for old_name in ("latest.xlsx", "latest.xlsm"):
                    old_path = os.path.join(folder_path, old_name)
                    if os.path.exists(old_path):
                        try:
                            os.chmod(old_path, 0o644)
                            os.remove(old_path)
                            print(f"🗑️ [DEBUG] تم حذف الملف القديم: {old_path}")
                        except Exception as e:
                            print(f"⚠️ [DEBUG] تحذير حذف {old_name}: {e}")
            if os.path.exists(file_path):
                try:
                    os.chmod(file_path, 0o644)
                    os.remove(file_path)
                    print(f"🗑️ [DEBUG] تم حذف الملف القديم: {file_path}")
                except PermissionError as pe:
                    print(
                        f"⚠️ [DEBUG] تحذير: لا يمكن حذف الملف القديم (PermissionError): {pe}"
                    )
                    temp_path = os.path.join(folder_path, "temp_upload.xlsx")
                    with open(temp_path, "wb+") as destination:
                        for chunk in excel_file.chunks():
                            destination.write(chunk)
                    try:
                        os.replace(temp_path, file_path)
                        print(f"✅ [DEBUG] تم استبدال الملف باستخدام os.replace")
                    except Exception as replace_error:
                        print(
                            f"⚠️ [DEBUG] تحذير: لا يمكن استبدال الملف: {replace_error}"
                        )
                        file_path = temp_path
                except Exception as delete_error:
                    print(f"⚠️ [DEBUG] تحذير: خطأ في حذف الملف القديم: {delete_error}")

            # ✅ حفظ الملف الجديد
            with open(file_path, "wb+") as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            try:
                os.chmod(file_path, 0o644)
            except Exception as chmod_error:
                print(f"⚠️ [DEBUG] تحذير: لا يمكن تغيير صلاحيات الملف: {chmod_error}")

            print(f"✅ [DEBUG] تم حفظ الملف بنجاح في: {file_path}")

            # ✅ حفظ المسار في الجلسة حسب نوع الملف (داشبورد أو رئيسي)
            if is_dashboard_file:
                request.session["dashboard_excel_path"] = file_path
                print(f"💾 [DEBUG] تم حفظ مسار ملف الداشبورد في الجلسة: {file_path}")
            else:
                request.session["uploaded_excel_path"] = file_path
                print(f"💾 [DEBUG] تم حفظ مسار الملف الرئيسي في الجلسة: {file_path}")
                # ✅ ملء كاش الشيتات في الداتابيز (JSON) لتسريع فتح التابات
                try:
                    xls = pd.ExcelFile(file_path, engine="openpyxl")
                    for sheet_name in xls.sheet_names:
                        try:
                            df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
                            cache_rows = _dataframe_to_cache_rows(df)
                            ExcelSheetCache.objects.update_or_create(
                                sheet_name=sheet_name,
                                defaults={"data": cache_rows, "source_file_path": os.path.normpath(os.path.abspath(file_path))},
                            )
                        except Exception as sheet_err:
                            print(f"⚠️ [DEBUG] تخطي شيت '{sheet_name}': {sheet_err}")
                    print(f"✅ [DEBUG] تم تحديث ExcelSheetCache لـ {len(xls.sheet_names)} شيت")
                except Exception as cache_fill_err:
                    print(f"⚠️ [DEBUG] تعبئة الكاش: {cache_fill_err}")
            # ✅ ملء كاش الداشبورد في الداتابيز + ملف JSON (تحديث دون حذف بيانات المسارات الأخرى، مع إزالة التكرار)
            try:
                cache_data = _build_dashboard_cache_data(file_path)
                if cache_data:
                    cache_data["_file_mtime"] = _get_file_mtime(file_path)
                    norm_path = os.path.normpath(os.path.abspath(file_path))
                    DashboardDataCache.objects.update_or_create(
                        source_file_path=norm_path,
                        defaults={"data": _dedupe_cache_data(cache_data)},
                    )
                    _save_dashboard_cache_to_json(norm_path, cache_data)
                    print("✅ [DEBUG] تم تحديث DashboardDataCache وملف JSON")
            except Exception as dash_cache_err:
                print(f"⚠️ [DEBUG] تعبئة كاش الداشبورد: {dash_cache_err}")
            request.session.save()

            # ✅ مسح الكاش بعد رفع ملف جديد
            try:
                cache.clear()
                print(f"🗑️ [DEBUG] تم مسح الكاش")
            except Exception as cache_error:
                print(f"⚠️ [DEBUG] تحذير: لا يمكن مسح الكاش: {cache_error}")

            # ✅ إرجاع response
            if request.headers.get("x-requested-with") == "XMLHttpRequest":
                return JsonResponse(
                    {"success": True, "message": "✅ File uploaded successfully!"}
                )
            messages.success(request, "✅ File uploaded successfully!")
            return redirect(request.path)
        except Exception as e:
            import traceback

            error_trace = traceback.format_exc()
            print(f"❌ [DEBUG] خطأ في حفظ الملف: {e}")
            print(f"❌ [DEBUG] Traceback:\n{error_trace}")
            if request.headers.get("x-requested-with") == "XMLHttpRequest":
                return JsonResponse(
                    {"error": f"❌ Error saving file: {str(e)}"}, status=500
                )
            messages.error(request, f"❌ Error saving file: {str(e)}")
            return redirect(request.path)

    def export_dashboard_excel(
        self, request, selected_month=None, selected_months=None
    ):
        """
        تحميل الملف الأصلي للإكسل (all_sheet) — نفس الملف المستخدم لكل التابات.
        أولوية: ملف الجلسة المرفوع ثم latest ثم all_sheet في المجلد.
        """
        # استخدام نفس مصدر الملف الذي تُقرأ منه كل التابات (all_sheet / ملف مرفوع)
        excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
        if not excel_path or not os.path.exists(excel_path):
            html = (
                "<!DOCTYPE html><html><head><meta charset='utf-8'><title>File not found</title></head><body style='font-family:sans-serif;padding:2rem;'>"
                "<h2>Excel file not found</h2>"
                "<p>Please upload the Excel file first (use <strong>Upload File</strong> on the main page).</p>"
                "<p><a href='javascript:window.close()'>Close this tab</a></p>"
                "</body></html>"
            )
            return HttpResponse(html, status=404, content_type="text/html")

        try:
            # اسم الملف للتنزيل: اسم الملف الأصلي إن أمكن
            download_name = os.path.basename(excel_path)
            if not download_name or download_name == "latest.xlsx":
                download_name = "All_Sheets.xlsx"

            # تحديد نوع المحتوى حسب الامتداد
            ext = os.path.splitext(download_name)[1].lower()
            if ext == ".xlsm":
                content_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
            else:
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            with open(excel_path, "rb") as f:
                file_data = f.read()

            response = HttpResponse(file_data, content_type=content_type)
            response["Content-Disposition"] = (
                f'attachment; filename="{download_name}"'
            )
            return response

        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(f"❌ حدث خطأ عند تحميل الملف: {str(e)}", status=500)

    def render_raw_sheet(self, request, sheet_name):
        """عرض أي شيت كجدول خام إذا مفيش فلتر خاص"""
        print(f"🟢 [DEBUG] ✅ دخل على render_raw_sheet() - التاب: {sheet_name}")

        # 📁 جلب مسار ملف الإكسل
        excel_file_path = self.get_uploaded_file_path(request)
        if not excel_file_path or not os.path.exists(excel_file_path):
            print("⚠️ [ERROR] لم يتم العثور على ملف Excel.")
            return {
                "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                "count": 0,
            }

        try:
            # 📖 قراءة جميع الشيتات
            xls = pd.ExcelFile(excel_file_path, engine="openpyxl")

            # 🔍 البحث عن الشيت بدون حساسية لحالة الأحرف
            matching_sheet = next(
                (
                    s
                    for s in xls.sheet_names
                    if s.lower().strip() == sheet_name.lower().strip()
                ),
                None,
            )

            if not matching_sheet:
                print(
                    f"⚠️ [WARNING] التاب '{sheet_name}' غير موجود. الشيتات المتاحة: {xls.sheet_names}"
                )
                return {
                    "detail_html": f"<p class='text-danger'>❌ Tab '{sheet_name}' does not exist in the file.</p>",
                    "count": 0,
                }

            # 🧾 قراءة الشيت المطابق
            df = pd.read_excel(
                excel_file_path, sheet_name=matching_sheet, engine="openpyxl"
            )

            # 🧹 تنظيف الأعمدة
            df.columns = df.columns.str.strip().str.title()

            # 🗓️ فلترة حسب الشهر إذا تم اختياره
            selected_month = request.GET.get("month")
            if selected_month:
                date_cols = [c for c in df.columns if "Date" in c]
                if date_cols:
                    df[date_cols[0]] = pd.to_datetime(df[date_cols[0]], errors="coerce")
                    df["Month"] = df[date_cols[0]].dt.strftime("%b")
                    df = df[df["Month"] == selected_month]

            # 🧩 طباعة حالة البيانات
            if df.empty:
                print(
                    f"⚠️ [WARNING] الشيت '{matching_sheet}' فاضي أو غير موجود بعد الفلترة!"
                )
            else:
                print(
                    f"✅ [INFO] الشيت '{matching_sheet}' اتقرأ بنجاح وفيه {len(df)} صفوف."
                )
                print(f"📋 [COLUMNS] الأعمدة: {list(df.columns)}")

            # 🔢 تجهيز أول 50 صف فقط للعرض
            data = df.head(50).to_dict(orient="records")
            for row in data:
                for col, val in row.items():
                    row[col] = self.safe_format_value(val)

            # 🧩 توليد HTML من التمبلت
            tab_data = {
                "name": matching_sheet,
                "columns": df.columns.tolist(),
                "data": data,
            }
            month_norm = self.apply_month_filter_to_tab(tab_data, selected_month)

            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm},
            )

            # 📤 إرجاع النتيجة للواجهة
            return {"detail_html": html, "count": len(df), "tab_data": tab_data}

        except Exception as e:
            print(f"❌ [ERROR] أثناء قراءة الشيت '{sheet_name}': {e}")
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while reading sheet: {e}</p>",
                "count": 0,
            }

    def filter_by_month(self, request, selected_month):
        import pandas as pd
        from django.template.loader import render_to_string

        try:
            excel_file_path = self.get_uploaded_file_path(request)
            xls = pd.ExcelFile(excel_file_path, engine="openpyxl")

            # 🧩 تحديد اسم الشيت المطلوب تلقائيًا
            # نحاول نختار شيت يحتوي على "Data logger" أو "Dock to stock"
            possible_sheets = [
                s
                for s in xls.sheet_names
                if any(key in s.lower() for key in ["data logger", "dock to stock"])
            ]

            if not possible_sheets:
                print(
                    "⚠️ لم يتم العثور على أي شيت يحتوي على Data logger أو Dock to stock"
                )
                return {
                    "error": "⚠️ No sheet containing Data logger or Dock to stock was found."
                }

            sheet_name = possible_sheets[0]  # ناخد أول واحد مطابق
            print(f"📄 قراءة الشيت: {sheet_name}")

            df = pd.read_excel(
                excel_file_path, sheet_name=sheet_name, engine="openpyxl"
            )
        except Exception as e:
            return {"error": f"⚠️ Unable to read the tab: {e}"}

        # تنظيف الأعمدة
        df.columns = df.columns.str.strip()

        # التحقق من عمود التاريخ
        if "Month" not in df.columns:
            return {"error": "⚠️ Column 'Month' is missing."}

        # تحويل/تطبيع عمود الشهر لقبول كل الصيغ (تاريخ، اختصار، اسم كامل، رقم 1-12)
        import calendar

        month_raw = df["Month"]
        # حاول تحويله لتاريخ؛ اللي يفشل هنرجّعه نصياً
        parsed = pd.to_datetime(month_raw, errors="coerce")
        month_abbr_from_dates = parsed.dt.strftime("%b")

        # طبّع النصوص: أول 3 حروف من اسم الشهر (Jan/February -> Feb)، والأرقام 1-12 إلى اختصار
        def normalize_month_val(v):
            if pd.isna(v):
                return None
            s = str(v).strip()
            # أرقام
            if s.isdigit():
                n = int(s)
                if 1 <= n <= 12:
                    return calendar.month_abbr[n]
            # أسماء كاملة أو مختصرة
            # جرّب اسم كامل
            for i, mname in enumerate(calendar.month_name):
                if i == 0:
                    continue
                if s.lower() == mname.lower():
                    return calendar.month_abbr[i]
            # جرّب اختصار جاهز أو نص عام -> أول 3 أحرف بحالة Capitalize
            return s[:3].capitalize()

        month_abbr_fallback = month_raw.apply(normalize_month_val)
        # استخدم من التاريخ حيث متاح وإلا fallback
        df["Month"] = month_abbr_from_dates.where(~parsed.isna(), month_abbr_fallback)

        # توحيد تمثيل الشهر المختار (أمان لحالات الإدخال المختلفة)
        selected_month_norm = (
            str(selected_month).strip().capitalize() if selected_month else None
        )

        # حفظ الشهر في الجلسة ليستخدمه باقي التابات عند الاستعلامات اللاحقة
        try:
            if selected_month_norm:
                request.session["selected_month"] = selected_month_norm
        except Exception:
            # في حال عدم توفر الجلسة (مثلاً في طلبات غير مرتبطة بمستخدم)، نتجاوز بهدوء
            pass

        # فلترة الشهر المختار أولاً
        month_df = df[df["Month"] == selected_month_norm]

        if month_df.empty:
            return {
                "error": f"⚠️ لا توجد بيانات متاحة للشهر {selected_month_norm}.",
                "month": selected_month_norm,
                "sheet_name": sheet_name,
            }

        # البحث عن عمود KPI بشكل مرن (ممكن يكون اسمه مختلف)
        kpi_miss_col = None
        possible_kpi_names = [
            "kpi miss in",
            "kpi miss",
            "kpi",
            "miss",
            "clearance handling kpi",
            "transit kpi",
        ]

        for kpi_name in possible_kpi_names:
            kpi_miss_col = next(
                (col for col in df.columns if str(col).strip().lower() == kpi_name),
                None,
            )
            if kpi_miss_col:
                break

        # حساب الإحصائيات
        total = len(month_df.drop_duplicates())

        # لو وجدنا عمود KPI، نحسب Miss
        if kpi_miss_col:
            miss_df = month_df[month_df[kpi_miss_col].astype(str).str.lower() == "miss"]
            miss_count = len(miss_df)
            valid = total - miss_count
        else:
            # لو مفيش عمود KPI، نعرض كل البيانات بدون فلترة Miss
            miss_df = pd.DataFrame()  # جدول فاضي
            miss_count = 0
            valid = total
            print(
                f"⚠️ لم يتم العثور على عمود KPI، سيتم عرض جميع البيانات للشهر {selected_month_norm}"
            )

        # تحويل النتائج إلى HTML (للحفاظ على التوافق مع أي استخدام حالي)
        dedup_html = month_df.to_html(
            classes="table table-bordered table-hover text-center",
            index=False,
            border=0,
        )
        miss_html = miss_df.to_html(
            classes="table table-bordered table-hover text-center text-danger",
            index=False,
            border=0,
        )

        print(
            f"📆 فلترة الشهر {selected_month}: إجمالي={total}, Miss={miss_count}, Valid={valid}"
        )

        hit_pct = int(round((valid / total) * 100)) if total else 0

        # تجهيز البيانات للتمبلت القياسي (جداول + شارت)
        month_df_display = month_df.fillna("").astype(str)
        sub_tables = [
            {
                "title": f"{sheet_name} – {selected_month_norm} (كل السجلات)",
                "columns": month_df_display.columns.tolist(),
                "data": month_df_display.to_dict(orient="records"),
            }
        ]

        if miss_count > 0:
            miss_df_display = miss_df.fillna("").astype(str)
            sub_tables.append(
                {
                    "title": f"{sheet_name} – {selected_month_norm} (السجلات المتأخرة)",
                    "columns": miss_df_display.columns.tolist(),
                    "data": miss_df_display.to_dict(orient="records"),
                }
            )

        summary_table = [
            {"المؤشر": "إجمالي الشحنات", "القيمة": int(total)},
            {"المؤشر": "شحنات صحيحة", "القيمة": int(valid)},
            {"المؤشر": "شحنات Miss", "القيمة": int(miss_count)},
            {"المؤشر": "Hit %", "القيمة": f"{hit_pct}%"},
        ]
        sub_tables.append(
            {
                "title": f"{sheet_name} – {selected_month_norm} (ملخص الأداء)",
                "columns": ["المؤشر", "القيمة"],
                "data": summary_table,
            }
        )

        chart_title = f"{sheet_name} – {selected_month_norm} Performance"
        chart_data = [
            {
                "title": chart_title,
                "type": "column",
                "name": "Valid Shipments",
                "color": "#4caf50",
                "showInLegend": True,
                "dataPoints": [{"label": selected_month_norm, "y": int(valid)}],
                "related_table": sub_tables[0]["title"],
            },
            {
                "title": chart_title,
                "type": "column",
                "name": "Miss Shipments",
                "color": "#f44336",
                "showInLegend": True,
                "dataPoints": [{"label": selected_month_norm, "y": int(miss_count)}],
                "related_table": sub_tables[0]["title"],
            },
            {
                "title": chart_title,
                "type": "line",
                "name": "Hit %",
                "color": "#1976d2",
                "showInLegend": True,
                "dataPoints": [{"label": selected_month_norm, "y": hit_pct}],
                "related_table": sub_tables[-1]["title"],
            },
        ]

        tab_data = {
            "name": f"{sheet_name} ({selected_month_norm})",
            "sub_tables": sub_tables,
            "chart_data": chart_data,
            "chart_title": chart_title,
        }
        month_norm_filtered = self.apply_month_filter_to_tab(
            tab_data, selected_month_norm
        )

        combined_html = render_to_string(
            "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
            {"tab": tab_data, "selected_month": month_norm_filtered},
        )

        return {
            "month": selected_month_norm,
            "selected_month": selected_month_norm,
            "sheet_name": sheet_name,
            "total_shipments": total,
            "miss_count": miss_count,
            "valid_shipments": valid,
            "hit_pct": hit_pct,
            "dedup_html": dedup_html,
            "miss_html": miss_html,
            "html": combined_html,
            "detail_html": combined_html,
            "chart_data": chart_data,
            "chart_title": chart_title,
            "tab_data": tab_data,
        }

    def _resolve_quarter_months(self, selected_quarter):
        if not selected_quarter:
            return []

        import re

        quarter_pattern = re.compile(r"^Q([1-4])(?:[-\s]?(\d{4}))?$", re.IGNORECASE)
        match = quarter_pattern.match(str(selected_quarter).strip())
        if not match:
            raise ValueError(f"⚠️ كورتر غير معروف: {selected_quarter}")

        quarter_number = int(match.group(1))
        quarter_months_map = {
            1: ["Jan", "Feb", "Mar"],
            2: ["Apr", "May", "Jun"],
            3: ["Jul", "Aug", "Sep"],
            4: ["Oct", "Nov", "Dec"],
        }

        months = quarter_months_map.get(quarter_number, [])
        if not months:
            raise ValueError(f"⚠️ لا توجد شهور معرّفة للكوارتر {selected_quarter}.")
        return months

    def filter_by_quarter(self, request, selected_quarter):
        from django.template.loader import render_to_string
        import re

        if not selected_quarter:
            return {"error": "⚠️ Please select a valid quarter."}

        quarter_pattern = re.compile(r"^Q([1-4])(?:[-\s]?(\d{4}))?$", re.IGNORECASE)
        match = quarter_pattern.match(str(selected_quarter).strip())
        if not match:
            return {"error": f"⚠️ Unknown quarter: {selected_quarter}"}

        quarter_number = int(match.group(1))
        quarter_months_map = {
            1: ["Jan", "Feb", "Mar"],
            2: ["Apr", "May", "Jun"],
            3: ["Jul", "Aug", "Sep"],
            4: ["Oct", "Nov", "Dec"],
        }

        display_month_list = quarter_months_map.get(quarter_number, [])
        if not display_month_list:
            return {
                "error": f"⚠️ No months were defined for quarter {selected_quarter}."
            }

        try:
            total_lead_time_result = self.filter_total_lead_time_performance(
                request, selected_months=display_month_list
            )
        except Exception as exc:
            import traceback

            traceback.print_exc()
            total_lead_time_result = {
                "detail_html": f"<p class='text-danger text-center p-4'>⚠️ Error while loading Total Lead Time Performance: {exc}</p>"
            }

        section_html = (
            total_lead_time_result.get("detail_html")
            or total_lead_time_result.get("html")
            or "<p class='text-warning text-center p-4'>⚠️ No data available for this quarter.</p>"
        )

        section_wrapper = f"""
        <section class="quarter-section mb-5">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h4 class="mb-0 text-primary">Total Lead Time Performance – Quarter {selected_quarter}</h4>
                <span class="badge bg-light text-dark px-3 py-2">{', '.join(display_month_list)}</span>
            </div>
            {section_html}
        </section>
        """

        header_html = f"""
        <div class="quarter-header text-center mb-4">
            <h3 class="fw-bold text-primary mb-1">Quarter {selected_quarter}</h3>
            <p class="text-muted mb-0">Months in scope: {', '.join(display_month_list)}</p>
        </div>
        """

        combined_html = (
            f"<div class='quarter-wrapper'>{header_html}{section_wrapper}</div>"
        )

        return {
            "quarter": selected_quarter,
            "months": ", ".join(display_month_list),
            "detail_html": combined_html,
            "html": combined_html,
            "chart_data": total_lead_time_result.get("chart_data", []),
            "chart_title": total_lead_time_result.get("chart_title"),
            "hit_pct": total_lead_time_result.get("hit_pct"),
        }

    def filter_all_tabs(self, request=None, selected_month=None, selected_months=None):
        cache.clear()
        try:
            month_for_filters = selected_month if not selected_months else None
            # ✅ تحديد الفلتر الحالي
            status_filter = "all"
            if request is not None and hasattr(request, "GET"):
                status_filter = request.GET.get("status", "all")

            # ✅ الحصول على مسار ملف Excel
            excel_path = self.get_uploaded_file_path(request)
            if not excel_path or not os.path.exists(excel_path):
                html = render_to_string(
                    "components/ui-kits/tab-bootstrap/components/dashboard-overview.html",
                    {"message": "⚠️ لم يتم العثور على ملف Excel."},
                )
                return {"detail_html": html}

            # ✅ جلب بيانات الـ overview_tab (منها نبدأ النسب)
            overview_data = self.overview_tab(
                request=request,
                selected_month=month_for_filters,
                selected_months=selected_months,
                from_all_in_one=True,
            )

            if not overview_data or "tab_cards" not in overview_data:
                html = render_to_string(
                    "components/ui-kits/tab-bootstrap/components/dashboard-overview.html",
                    {"message": "⚠️ لا توجد بيانات متاحة من overview_tab."},
                )
                return {"detail_html": html}

            # ✅ قائمة التابات المحذوفة (فقط المطلوب إخفاؤها من All-in-One)
            excluded_tabs = [
                "airport clearance",
                "seaport clearance",
                "data logger measurement",
            ]

            clean_tabs = []
            for tab in overview_data.get("tab_cards", []):
                name = tab.get("name", "غير معروف")
                name_lower = name.strip().lower()

                if name_lower in excluded_tabs:
                    continue

                try:
                    hit = float(tab.get("hit_pct", 0))
                except Exception:
                    hit = 0
                hit = int(round(max(0, min(hit, 100))))

                try:
                    target = float(tab.get("target_pct", 100))
                except Exception:
                    target = 100

                chart_data = tab.get("chart_data", []) or []
                chart_type = tab.get("chart_type", "bar")

                clean_tabs.append(
                    {
                        "name": name,
                        "hit_pct": hit,
                        "target_pct": int(target),
                        "count": tab.get("count", 0),
                        "chart_type": chart_type,
                        "chart_data": chart_data,
                    }
                )

            # ✅ ترتيب التابات حسب الأولوية
            desired_order = [
                "Inbound",
                "Outbound",
                "Return & Refusal",
                "PODs update",
                "Inventory",
            ]
            clean_tabs.sort(
                key=lambda x: (
                    desired_order.index(x["name"])
                    if x["name"] in desired_order
                    else len(desired_order)
                )
            )

            # ✅ بيانات الميتنج - جلب كل النقاط (مثل meeting_points_tab)
            meeting_points = MeetingPoint.objects.all().order_by(
                "is_done", "-created_at"
            )

            if status_filter == "done":
                meeting_points = meeting_points.filter(is_done=True)
            elif status_filter == "pending":
                meeting_points = meeting_points.filter(is_done=False)

            meeting_data = [
                {
                    "id": p.id,
                    "description": p.description,
                    "assigned_to": getattr(p, "assigned_to", "") or "",
                    "status": "Done" if p.is_done else "Pending",
                    "created_at": p.created_at,
                    "target_date": p.target_date,
                }
                for p in meeting_points
            ]

            tabs_for_display = clean_tabs

            html = render_to_string(
                "components/ui-kits/tab-bootstrap/components/dashboard-overview.html",
                {
                    "tabs": tabs_for_display,
                    "tabs_json": json.dumps(tabs_for_display),
                    "meeting_data": meeting_data,
                    "status_filter": status_filter,
                },
                request=request,
            )

            return {"detail_html": html}

        except Exception as e:
            traceback.print_exc()
            return {
                "detail_html": f"<div class='alert alert-danger'>⚠️ Error: {e}</div>"
            }

    def get_meeting_points_section_html(self, request, status_filter="all"):
        """
        ✅ دالة مساعدة لإرجاع HTML قسم Meeting Points فقط
        """
        try:
            meeting_points = MeetingPoint.objects.all().order_by(
                "is_done", "-created_at"
            )

            if status_filter == "done":
                meeting_points = meeting_points.filter(is_done=True)
            elif status_filter == "pending":
                meeting_points = meeting_points.filter(is_done=False)

            meeting_data = [
                {
                    "id": p.id,
                    "description": p.description,
                    "assigned_to": getattr(p, "assigned_to", "") or "",
                    "status": "Done" if p.is_done else "Pending",
                    "created_at": p.created_at,
                    "target_date": p.target_date,
                }
                for p in meeting_points
            ]

            # ✅ إرجاع HTML قسم Meeting Points فقط
            html = render_to_string(
                "components/ui-kits/tab-bootstrap/components/meeting_points_section.html",
                {
                    "meeting_data": meeting_data,
                    "status_filter": status_filter,
                },
                request=request,
            )
            return html
        except Exception as e:
            import traceback

            traceback.print_exc()
            return f"<div class='alert alert-danger'>⚠️ Error: {e}</div>"

    def filter_total_lead_time_detail(self, request, selected_month=None):
        try:
            # تحميل الملف من الجلسة
            excel_path = request.session.get("uploaded_excel_path")
            if not excel_path or not os.path.exists(excel_path):
                return {"error": "⚠️ Excel file was not found in the session."}

            # قراءة الشيت المطلوب
            df = pd.read_excel(
                excel_path, sheet_name="Total lead time preformance", engine="openpyxl"
            )
            df.columns = df.columns.str.strip().str.lower()

            # التأكد من الأعمدة المطلوبة
            required_cols = [
                "month",
                "outbound delivery",
                "kpi",
                "reason group",
                "miss reason",
            ]
            for col in required_cols:
                if col not in df.columns:
                    return {"error": f"⚠️ Column '{col}' does not exist in the sheet."}

            # تحويل التاريخ إلى الشهر
            df["month"] = (
                pd.to_datetime(df["month"], errors="coerce")
                .dt.strftime("%b")
                .str.capitalize()
            )

            # استخراج الشهور الموجودة فعليًا في الملف (بترتيب زمني)
            existing_months = df["month"].dropna().unique().tolist()
            existing_months = sorted(
                existing_months, key=lambda x: pd.to_datetime(x, format="%b").month
            )

            if not existing_months:
                return {"error": "⚠️ No valid months were found in the file."}

            # إزالة التكرارات حسب رقم الشحنة
            df = df.drop_duplicates(subset=["outbound delivery"])

            # تنظيف النصوص
            df["reason group"] = df["reason group"].astype(str).str.strip().str.lower()
            df["kpi"] = df["kpi"].astype(str).str.strip().str.lower()

            # بيانات Miss الخاصة بـ 3PL فقط
            df_miss_3pl = df[
                (df["kpi"] == "miss") & (df["reason group"] == "3pl")
            ].copy()

            # 🔹 تنظيف السبب فقط (بدون تغيير الحروف الأصلية)
            df_miss_3pl["miss reason"] = (
                df_miss_3pl["miss reason"]
                .astype(str)
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)  # إزالة المسافات المكررة
            )

            # معالجة اختلاف الحروف أثناء التجميع (case-insensitive grouping)
            df_miss_3pl["_miss_reason_key"] = df_miss_3pl["miss reason"].str.lower()

            # بيانات On Time Delivery (Hit)
            df_hit = df[df["kpi"] != "miss"].copy()

            # تجميع Miss حسب السبب والشهر (باستخدام المفتاح الموحد للحروف)
            miss_grouped = (
                df_miss_3pl.groupby(["_miss_reason_key", "month"], as_index=False)
                .agg(
                    {
                        "miss reason": "first",
                        "month": "first",
                        "_miss_reason_key": "count",
                    }
                )
                .rename(columns={"_miss_reason_key": "count"})
            )

            # Pivot الجدول
            miss_pivot = miss_grouped.pivot_table(
                index="miss reason", columns="month", values="count", fill_value=0
            )

            # تأكد أن كل الشهور الموجودة في الملف موجودة في الجدول
            for m in existing_months:
                if m not in miss_pivot.columns:
                    miss_pivot[m] = 0
            miss_pivot = miss_pivot[existing_months]

            # حساب On Time Delivery لكل شهر
            hit_counts = (
                df_hit.groupby("month").size().reindex(existing_months, fill_value=0)
            )

            # بناء الجدول النهائي
            final_df = miss_pivot.copy()
            final_df.loc["On Time Delivery"] = hit_counts
            final_df = final_df.fillna(0)

            # ترتيب الصفوف بحيث On Time في الأعلى
            final_df = final_df.reindex(
                ["On Time Delivery"]
                + [r for r in final_df.index if r != "On Time Delivery"]
            )

            # إضافة عمود الإجمالي
            final_df["Total"] = final_df.sum(axis=1)

            # صف الإجمالي النهائي
            final_df.loc["TOTAL"] = final_df.sum(numeric_only=True)

            # 🟦 إنشاء جدول HTML
            html_table = """
            <table class='table table-bordered text-center align-middle mb-0'>
                <thead class='table-warning'>
                    <tr><th colspan='{colspan}'>Reason From 3PL Side</th></tr>
                </thead>
                <thead class='table-primary'>
                    <tr>
                        <th>KPI</th>
                        {month_headers}
                        <th>2025</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
            """

            # رؤوس الأعمدة
            month_headers = "".join([f"<th>{m}</th>" for m in existing_months])

            # الصفوف
            rows_html = ""
            for reason, row in final_df.iterrows():
                rows_html += f"<tr><td>{reason}</td>"
                for m in existing_months:
                    rows_html += f"<td>{int(row[m])}</td>"
                rows_html += f"<td class='fw-bold'>{int(row['Total'])}</td></tr>"

            # استبدال القيم في القالب
            html_table = html_table.format(
                colspan=len(existing_months) + 2,
                month_headers=month_headers,
                table_rows=rows_html,
            )

            # وضع الجدول داخل واجهة مرتبة
            html_output = f"""
            <div class='container-fluid'>
                <h5 class='text-center text-primary mb-3'>KPI Summary - 3PL Performance</h5>
                <div class='card shadow'>
                    <div class='card-body'>
                        {html_table}
                    </div>
                </div>
            </div>
            """

            return {"detail_html": html_output, "months": existing_months}

        except Exception as e:
            import traceback

            print("❌ خطأ أثناء تحليل البيانات:", str(e))
            print(traceback.format_exc())
            return {"error": f"⚠️ Error while analyzing data: {e}"}

    def filter_rejection_data(self, request, month=None):
        print("🟣 [DEBUG] filter_rejection_data CALLED ✅ month:", month)

        excel_path = request.session.get("uploaded_excel_path")

        if not excel_path or not os.path.exists(excel_path):
            return {"error": "⚠️ Excel file not found."}

        try:
            df = pd.read_excel(excel_path, sheet_name="Rejection", engine="openpyxl")
            print("🟢 [DEBUG] الأعمدة:", df.columns.tolist())
            print(df.head(3))
        except Exception as e:
            return {"error": f"⚠️ Unable to read the 'Rejection' sheet: {e}"}

        df.columns = df.columns.str.strip().str.title()
        required = ["Month", "Total Number Of Orders", "Booking Orders"]
        if not all(col in df.columns for col in required):
            return {
                "error": "⚠️ Required columns (Month, Total Number Of Orders, Booking Orders) are missing."
            }

        if month:
            df = df[df["Month"].astype(str).str.contains(month, case=False, na=False)]

        if df.empty:
            return {"error": "⚠️ No data available."}

        # ✅ خدي القيم زي ما هي من الإكسل (من العمود Booking Orders)
        summary = df[["Month", "Booking Orders"]].copy()

        # 🧠 تنظيف القيم — شيل علامة % لو موجودة وحوّليها لأرقام
        summary["Booking Orders"] = (
            summary["Booking Orders"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .astype(float)
        )

        # 🎯 البيانات للشارت مباشرة
        chart_data = [
            {"month": row["Month"], "percentage": row["Booking Orders"]}
            for _, row in summary.iterrows()
        ]

        html = df.to_html(
            index=False,
            classes="table table-bordered table-striped text-center align-middle",
            border=0,
        )

        print("📊 DEBUG - chart_data:", chart_data)  # <-- شوفيها في التيرمنال
        return {"detail_html": html, "chart_data": chart_data}

    def filter_dock_to_stock_roche(self, request, selected_month=None):
        print("🟢 [DEBUG] ✅ دخل على filter_dock_to_stock_roche()")

        excel_path = request.session.get("uploaded_excel_path")
        if not excel_path or not os.path.exists(excel_path):
            return {"error": "⚠️ Excel file not found."}

        try:
            import pandas as pd
            from django.template.loader import render_to_string

            sheet_name = "Dock to stock - Roche"
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()

            if df.empty:
                return {"error": "⚠️ Sheet 'Dock to stock - Roche' is empty."}

            # أول عمود هو الشهر
            month_col = df.columns[0]
            # باقي الأعمدة هي الأسباب (KPIs)
            kpi_cols = df.columns[1:]

            # تحويل البيانات بحيث تكون الأسباب صفوف والشهور أعمدة
            melted_df = df.melt(id_vars=[month_col], var_name="KPI", value_name="Value")

            # Pivot فعلي (KPI كصفوف والشهور كأعمدة)
            pivot_df = melted_df.pivot_table(
                index="KPI", columns=month_col, values="Value", aggfunc="sum"
            ).reset_index()
            pivot_df = pivot_df.rename_axis(None, axis=1)

            # ترتيب الأعمدة حسب تسلسل الشهور الموجود في الشيت الأصلي
            month_order = list(df[month_col].unique())
            ordered_cols = ["KPI"] + month_order
            pivot_df = pivot_df.reindex(columns=ordered_cols)

            # ✅ حذف أي عمود اسمه "Total" (اللي بيتولد من الشيت أو من الخطأ)
            if "Total" in pivot_df.columns:
                pivot_df = pivot_df.drop(columns=["Total"])

            # ✅ إضافة عمود "2025" فقط بعد الشهور
            pivot_df["2025"] = pivot_df.iloc[:, 1:].sum(axis=1)

            # ✅ إضافة صف Total (اللي بيكون تحت الجدول)
            total_row = {"KPI": "Total"}
            for col in pivot_df.columns[1:]:  # تجاهل عمود KPI
                total_row[col] = pivot_df[col].sum()
            pivot_df = pd.concat(
                [pivot_df, pd.DataFrame([total_row])], ignore_index=True
            )

            print("✅ [DEBUG] جدول KPI النهائي بعد التعديل:")
            print(pivot_df.to_string(index=False))

            # تجهيز البيانات للعرض
            columns = list(pivot_df.columns)
            table_data = pivot_df.fillna("").to_dict(orient="records")

            tab = {
                "name": "Dock to Stock - Roche",
                "columns": columns,
                "data": table_data,
            }

            month_norm = self.apply_month_filter_to_tab(tab, selected_month)

            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {
                    "tab": tab,
                    "table_title": "Dock to Stock - Roche (KPI Summary)",
                    "selected_month": month_norm,
                },
            )

            return {
                "detail_html": html,
                "chart_title": "Dock to Stock - Roche",
            }

        except Exception as e:
            print(f"❌ [ERROR] {e}")
            return {"error": f"⚠️ Error while reading data: {e}"}

    def filter_dock_to_stock_3pl(
        self, request, selected_month=None, selected_months=None
    ):
        try:
            print("🟢 [DEBUG] ✅ دخل على filter_dock_to_stock_3pl()")
            file_path = self.get_uploaded_file_path(request)
            print(f"📁 [DEBUG] مسار الملف المستخدم: {file_path}")

            if not file_path or not os.path.exists(file_path):
                return {"error": "⚠️ File not found."}

            # 🧩 قراءة الشيت
            df = pd.read_excel(file_path, sheet_name="Dock to stock", engine="openpyxl")
            print(f"📄 [DEBUG] أول 10 صفوف من الشيت Dock to stock:\n{df.head(10)}")

            # ✅ التحقق من وجود الأعمدة المطلوبة
            if "Delv #" not in df.columns or "Month" not in df.columns:
                return {
                    "error": "⚠️ Columns 'Delv #' or 'Month' are missing in the sheet."
                }

            # 🧮 استخراج الشهر من العمود Month
            df["Month"] = pd.to_datetime(df["Month"], errors="coerce").dt.strftime("%b")

            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in seen:
                        seen.add(norm)
                        selected_months_norm.append(norm)

            selected_month_norm = (
                self.normalize_month_label(selected_month)
                if selected_month and not selected_months_norm
                else None
            )
            if selected_months_norm:
                df = df[
                    df["Month"]
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning text-center'>⚠️ No data available for the selected quarter months.</p>",
                        "chart_data": [],
                    }
            elif selected_month_norm:
                df = df[df["Month"].str.lower() == selected_month_norm.lower()]
                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning text-center'>⚠️ No data available for this month.</p>",
                        "chart_data": [],
                    }

            # 🧱 حذف الصفوف اللي مافيهاش شهر
            df = df.dropna(subset=["Month"])

            # ✅ حساب عدد الشحنات الفريدة (hit) لكل شهر من العمود Delv #
            hits_per_month = (
                df.drop_duplicates(subset=["Delv #"])
                .groupby("Month")["Delv #"]
                .count()
                .reset_index(name="Hits")
            )

            print("📊 [DEBUG] نتائج عدد الشحنات الفريدة لكل شهر:")
            print(hits_per_month)

            # ✅ حساب إجمالي الشحنات (Total) لكل شهر قبل حذف المكرر
            total_per_month = (
                df.groupby("Month")["Delv #"]
                .count()
                .reset_index(name="Total_Shipments")
            )

            # ✅ دمج نتائج الـ hits مع الإجمالي
            merged = pd.merge(hits_per_month, total_per_month, on="Month", how="left")

            # ✅ حساب نسبة التارجت لكل شهر
            merged["Target_%"] = (
                merged["Hits"] / merged["Total_Shipments"] * 100
            ).round(2)

            print("📈 [DEBUG] بعد حساب نسبة التارجت:")
            print(merged)

            # ✅ تجهيز جدول KPI بصيغة نهائية
            kpi_name = "On Time Receiving"
            df_kpi = pd.DataFrame({"KPI": [kpi_name]})

            for _, row in merged.iterrows():
                month = row["Month"]
                hits = int(row["Hits"])
                df_kpi[month] = hits

            # ✅ إضافة صف جديد Total
            total_row = {"KPI": "Total"}
            for col in df_kpi.columns[1:]:  # تجاهل عمود KPI
                total_row[col] = df_kpi[col].sum()
            df_kpi = pd.concat([df_kpi, pd.DataFrame([total_row])], ignore_index=True)

            # ✅ إضافة عمود جديد "2025" يمثل مجموع كل الشهور
            df_kpi["2025"] = df_kpi.iloc[:, 1:].sum(axis=1)

            # ✅ إضافة صف جديد لنسبة التارجت
            target_row = {"KPI": "Target (%)"}
            for _, row in merged.iterrows():
                month = row["Month"]
                target_row[month] = row["Target_%"]
            target_row["2025"] = round(merged["Target_%"].mean(), 2)
            df_kpi = pd.concat([df_kpi, pd.DataFrame([target_row])], ignore_index=True)

            print("✅ [DEBUG] جدول KPI النهائي بعد الإضافات:")
            print(df_kpi.to_string(index=False))

            if selected_months_norm:
                desired_cols = ["KPI"] + [
                    m for m in selected_months_norm if m in df_kpi.columns
                ]
                if "2025" in df_kpi.columns:
                    desired_cols.append("2025")
                df_kpi = df_kpi[[col for col in desired_cols if col in df_kpi.columns]]
            elif selected_month_norm:
                keep_cols = ["KPI", selected_month_norm]
                if "2025" in df_kpi.columns:
                    keep_cols.append("2025")
                df_kpi = df_kpi[[col for col in keep_cols if col in df_kpi.columns]]

            # 🧾 تحويل الجدول إلى HTML
            html_table = df_kpi.to_html(
                classes="table table-bordered text-center table-striped", index=False
            )

            # 🔹 الإرجاع لعرض الجدول في الواجهة
            return {
                "detail_html": html_table,
                "chart_data": df_kpi.to_dict(orient="records"),
            }

        except Exception as e:
            print(f"❌ [EXCEPTION] خطأ أثناء تنفيذ الدالة: {e}")
            return {"error": str(e)}

    def filter_total_lead_time_detail(self, request, selected_month=None):
        excel_path = request.session.get("uploaded_excel_path")
        if not excel_path or not os.path.exists(excel_path):
            return {
                "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                "count": 0,
            }

        try:
            # قراءة الشيت المطلوب
            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next(
                (
                    s
                    for s in xls.sheet_names
                    if "total lead time preformance" in s.lower()
                ),
                None,
            )
            if not sheet_name:
                return {
                    "detail_html": "<p class='text-danger'>❌ Tab 'Total lead time preformance' does not exist in the file.</p>",
                    "count": 0,
                }

            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.str.strip().str.lower()

            # التحقق من الأعمدة المطلوبة
            required_cols = [
                "month",
                "outbound delivery",
                "kpi",
                "reason group",
                "miss reason",
            ]
            if not all(col in df.columns for col in required_cols):
                html = render_to_string(
                    "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                    {
                        "tabs": [
                            {
                                "name": sheet_name,
                                "columns": df.columns.tolist(),
                                "data": df.head(50).to_dict(orient="records"),
                            }
                        ]
                    },
                )
                return {"detail_html": html, "count": len(df)}

            # تحويل التاريخ إلى شهر
            df["month"] = (
                pd.to_datetime(df["month"], errors="coerce")
                .dt.strftime("%b")
                .str.capitalize()
            )

            # استخراج الشهور الموجودة فعليًا
            existing_months = sorted(
                df["month"].dropna().unique().tolist(),
                key=lambda x: pd.to_datetime(x, format="%b").month,
            )
            if not existing_months:
                return {
                    "detail_html": "<p class='text-danger'>⚠️ No valid months were found in the file.</p>",
                    "count": 0,
                }

            # تنظيف النصوص
            df["reason group"] = df["reason group"].astype(str).str.strip().str.lower()
            df["kpi"] = df["kpi"].astype(str).str.strip().str.lower()
            df["miss reason"] = (
                df["miss reason"]
                .astype(str)
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)
            )

            # بيانات Miss الخاصة بـ 3PL فقط
            df_miss_3pl = df[
                (df["kpi"] == "miss") & (df["reason group"] == "3pl")
            ].copy()
            df_miss_3pl["_reason_key"] = df_miss_3pl["miss reason"].str.lower()

            # بيانات Hit (On Time Delivery)
            df_hit = df[df["kpi"] != "miss"].copy()

            # تجميع Miss حسب السبب والشهر
            miss_grouped = df_miss_3pl.groupby(
                ["_reason_key", "month"], as_index=False
            ).agg({"miss reason": "first"})
            miss_grouped["count"] = (
                df_miss_3pl.groupby(["_reason_key", "month"]).size().values
            )

            miss_pivot = miss_grouped.pivot_table(
                index="miss reason", columns="month", values="count", fill_value=0
            )

            # إضافة أعمدة الشهور الناقصة
            for m in existing_months:
                if m not in miss_pivot.columns:
                    miss_pivot[m] = 0
            miss_pivot = miss_pivot[existing_months]

            # حساب On Time Delivery
            hit_counts = (
                df_hit.groupby("month").size().reindex(existing_months, fill_value=0)
            )

            # بناء الجدول النهائي
            final_df = miss_pivot.copy()
            final_df.loc["On Time Delivery"] = hit_counts
            final_df = final_df.fillna(0)

            # تحويل كل القيم لأعداد صحيحة
            final_df = final_df.astype(int)

            # إضافة عمود الإجمالي (2025 بدل TOTAL)
            final_df["2025"] = final_df.sum(axis=1)

            # صف الإجمالي النهائي
            total_row = final_df.sum(numeric_only=True)
            total_row.name = "TOTAL"
            final_df = pd.concat([final_df, pd.DataFrame([total_row])])

            # ترتيب الأعمدة
            final_df.reset_index(inplace=True)
            # final_df.rename(columns={"miss reason": "KPI"}, inplace=True)
            final_df.rename(columns={"index": "KPI"}, inplace=True)

            # ✅ تجهيز البيانات للتمبلت الديناميكي
            tab_data = {
                "name": "KPI Summary - 3PL Performance",
                "sub_tables": [
                    {
                        "title": "Reason From 3PL Side",
                        "columns": ["KPI"] + existing_months + ["2025"],
                        "data": final_df.to_dict(orient="records"),
                    }
                ],
            }

            month_norm = self.apply_month_filter_to_tab(tab_data, selected_month)
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm},
            )

            return {"detail_html": html, "count": len(df), "tab_data": tab_data}

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while reading data: {e}</p>",
                "count": 0,
            }

    def filter_total_lead_time_roche(self, request, selected_month=None):
        """
        🔹 قراءة شيت "Total lead time preformance -R" من التمبلت المرفوع
        🔹 استخراج أسباب التأخير وترتيبها حسب الشهور
        🔹 عرضها بتصميم الجدول الموحد
        """
        print("🟢 [DEBUG] ✅ دخل على filter_total_lead_time_roche()")

        excel_path = request.session.get("uploaded_excel_path")
        if not excel_path or not os.path.exists(excel_path):
            return {"error": "⚠️ Excel file not found."}

        try:
            # فتح ملف الإكسل
            xls = pd.ExcelFile(excel_path, engine="openpyxl")

            # 🔍 البحث عن الشيت المطلوب
            sheet_name = next(
                (s for s in xls.sheet_names if "preformance -r" in s.lower()), None
            )
            if not sheet_name:
                return {
                    "error": "⚠️ No sheet containing 'Total lead time preformance -R' was found."
                }

            # قراءة الشيت
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.str.strip()

            # التحقق من وجود الأعمدة المطلوبة
            if "Month" not in df.columns:
                return {"error": "⚠️ Column named 'Month' was not found in the sheet."}

            # ترتيب الشهور بالترتيب الزمني
            month_order = [
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Oct",
                "Nov",
                "Dec",
            ]
            df["Month"] = pd.Categorical(
                df["Month"], categories=month_order, ordered=True
            )
            df = df.sort_values("Month")

            # تحويل البيانات إلى شكل طويل (Melt)
            df_melted = df.melt(id_vars=["Month"], var_name="KPI", value_name="Count")

            # تجميع البيانات حسب السبب والشهر
            pivot_df = (
                df_melted.groupby(["KPI", "Month"])["Count"]
                .sum()
                .unstack(fill_value=0)
                .reindex(columns=month_order, fill_value=0)
            )

            # إضافة عمود الإجمالي السنوي
            pivot_df["2025"] = pivot_df.sum(axis=1)

            # صف الإجمالي الكلي
            total_row = pivot_df.sum(numeric_only=True)
            total_row.name = "TOTAL"
            pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])])

            # ✅ إعادة تسمية العمود الأول إلى KPI
            pivot_df.reset_index(inplace=True)
            pivot_df.rename(columns={"index": "KPI"}, inplace=True)

            # حذف الشهور الفارغة تمامًا (بدون بيانات)
            pivot_df = pivot_df.loc[:, (pivot_df != 0).any(axis=0)]

            # ✅ تجهيز بيانات الجدول لتمبلت الـ HTML
            tab = {
                "name": "Total Lead Time Performance - Roche Side",
                "columns": list(pivot_df.columns),
                "data": pivot_df.to_dict(orient="records"),
            }

            month_norm = self.apply_month_filter_to_tab(tab, selected_month)
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {
                    "tab": tab,
                    "table_title": "Roche Lead Time 2025",
                    "selected_month": month_norm,
                },
            )

            return {
                "detail_html": html,
                "message": "✅ تم عرض بيانات Roche Lead Time بنجاح.",
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {"error": f"⚠️ Error while reading Roche Lead Time data: {e}"}

    def filter_outbound(self, request, selected_month=None):
        """
        🔹 عرض تاب Outbound بخطوات أفقية من تمبلت خارجي
        """
        try:
            # ✅ الخطوات مع ألوان وخلفيات مختلفة
            raw_steps = [
                {
                    "title": "GI Issue<br>Pick & Pack",
                    "icon": "bi-receipt",
                    "bg": "#9fc0e4",
                    "text_color": "#fff",
                    "border": "4px solid #9fc0e4",
                    "sub_color": "#eee",
                },
                {
                    "title": "Prepare Docs<br>Invoice, PO and Market place",
                    "icon": "bi-box-seam",
                    "bg": "#e8f1fb",
                    "text_color": "#007fa3",
                    "border": "4px solid #9fc0e4",
                    "sub_color": "#000",
                },
                {
                    "title": "Dispatch Time<br>from Docs Ready till left from WH",
                    "icon": "bi-arrow-left-right",
                    "bg": "#9fc0e4",
                    "text_color": "#fff",
                    "border": "4px solid #9fc0e4",
                    "sub_color": "#eee",
                },
                {
                    "title": "Delivery<br>Deliver to Customer",
                    "icon": "bi-file-earmark-text",
                    "bg": "#e8f1fb",
                    "text_color": "#007fa3",
                    "border": "4px solid #9fc0e4",
                    "sub_color": "#000",
                },
            ]

            steps = []
            for step in raw_steps:
                # نقسم النص على <br>
                parts = step["title"].split("<br>")
                styled_title = ""
                for i, part in enumerate(parts):
                    # لو دا السطر الأخير → نستخدم sub_color
                    color = (
                        step["sub_color"] if i == len(parts) - 1 else step["text_color"]
                    )
                    styled_title += f"<span class='step-line d-block' style='color:{color};'>{part.strip()}</span>"

                steps.append(
                    {
                        "title": styled_title,
                        "icon": step["icon"],
                        "bg": step["bg"],
                        "text_color": step["text_color"],
                        "border": step["border"],
                    }
                )

            # ✅ تمرير البيانات إلى التمبلت
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/workflow.html",
                {
                    "table_title": "Outbound workflow",
                    "table_text": "Process Stages",
                    "table_span": "Way Of Calculation",
                    "table_text_bottom": "The KPI was calculated based full lead time Order creation to deliver the order to the customer Based on SLA for each city",
                    "process_steps_text": "=NETWORKDAYS(Order Date, Delivery Date,7)-1",
                    "steps": steps,
                    "workflow_type": "outbound",
                },
            )

            return {
                "detail_html": html,
                "message": "✅ Outbound steps displayed successfully.",
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {"error": f"⚠️ Error while rendering the Outbound tab: {e}"}

    def filter_outbound_shipments(
        self,
        request,
        selected_month=None,
        selected_months=None,
        selected_facility=None,
        selected_order_nbr=None,
    ):
        """
        🔹 يقرأ من شيت ARAMCO Outbound Report: Order Nbr, Create Order, Shipped date, Facility، إلخ.
        🔹 ترتيب الفلترة: Facility → Create Order (شهر/تاريخ) → Order Nbr.
        🔹 Hit/Miss: عدد الأسطر لكل Order Nbr؛ ≤50 سطر → مسموح يوم واحد؛ >50 سطر → مسموح يومين.
           الأيام المحسوبة = أيام عمل بين Create Order و Shipped date (يوم الجمعة لا يُحسب).
           Hit لو الأيام المستخدمة ≤ المسموح، وإلا Miss.
        🔹 جدول التفاصيل: صف واحد لكل طلب؛ فلتر Facility ثم Month ثم Order Nbr (واجهة + خادم).
        """
        try:
            import os

            # الملف الرئيسي لكل التابات (all_sheet / latest) مع أولوية للملف المرفوع في الجلسة
            excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")

            # طباعة أسماء الشيتات في الملف عشان نعرف مين موجود
            print("\n[Outbound] أسماء الشيتات في ملف الإكسل:", xls.sheet_names)

            # المقارنة (Create Order و Shipped date) تُقرأ من شيت "ARAMCO Outbound Report" — نبحث عنه أولاً
            outbound1_name = next(
                (s for s in xls.sheet_names if (s or "").strip().lower() == "aramco outbound report"),
                None,
            )
            if not outbound1_name:
                outbound1_name = next(
                    (s for s in xls.sheet_names if "aramco outbound report" in (s or "").lower()),
                    None,
                )
            # لو مفيش ARAMCO Outbound Report نرجع لشيت Outbound1 أو أي شيت outbound
            if not outbound1_name:
                for name in xls.sheet_names:
                    low = (name or "").lower().strip()
                    if "outbound" in low and ("1" in low or "one" in low or low == "outbound1"):
                        outbound1_name = name
                        break
                if not outbound1_name:
                    outbound1_name = next(
                        (s for s in xls.sheet_names if "outbound" in (s or "").lower() and "2" not in (s or "").lower()),
                        None,
                    )

            outbound2_name = None
            for name in xls.sheet_names:
                low = (name or "").lower().strip()
                if "outbound" in low and ("2" in low or "two" in low or low == "outbound2"):
                    outbound2_name = name
                    break
            if not outbound2_name:
                outbound2_name = next(
                    (s for s in xls.sheet_names if "outbound" in (s or "").lower() and "1" not in (s or "").lower()),
                    None,
                )
            # لو لسه مفيش Outbound2: نجرب أي شيت فيه عمود Packed Timestamp (أو Packed) + Order Nbr
            if not outbound2_name and outbound1_name:
                for sheet in xls.sheet_names:
                    if sheet == outbound1_name:
                        continue
                    try:
                        probe = pd.read_excel(
                            excel_path, sheet_name=sheet, engine="openpyxl", nrows=2
                        )
                        probe.columns = probe.columns.astype(str).str.strip()
                        has_order = any(
                            "order" in c.lower() and "nbr" in c.lower()
                            for c in probe.columns
                        ) or any("order nbr" in c.lower() for c in probe.columns)
                        has_packed = any("packed" in c.lower() for c in probe.columns)
                        if has_order and has_packed:
                            outbound2_name = sheet
                            print(
                                f"[Outbound] تم استخدام الشيت '{sheet}' كـ Outbound2 (فيه Order Nbr + Packed)"
                            )
                            break
                    except Exception:
                        continue
            if not outbound2_name and outbound1_name:
                print(
                    "[Outbound] ⚠️ لم يتم العثور على شيت Outbound2. تأكدي أن اسم الشيت يحتوي على 'outbound' و '2' أو أن فيه عمود Packed Timestamp و Order Nbr."
                )

            if not outbound1_name:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Sheet 'ARAMCO Outbound Report' (or Outbound1) not found. المقارنة بين Create Order و Shipped date تتم من هذا الشيت.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                }

            # قراءة الشيت — المقارنة بين Create Order و Shipped date تتم من هنا (ARAMCO Outbound Report أو Outbound1)
            df1 = pd.read_excel(
                excel_path, sheet_name=outbound1_name, engine="openpyxl"
            )
            df1.columns = df1.columns.astype(str).str.strip()

            # لو أول عمود شكله "Unnamed" أو فيه عنوان التقرير (زي ARAMCO Outbound Report)
            # يبقى الأغلب إن أول صف هو عنوان والترويسات في صف تاني → نعيد قراءة الشيت ونكتشف صف الترويسات
            first_col_ob = str(df1.columns[0]).strip() if len(df1.columns) else ""
            if first_col_ob.startswith("Unnamed:") or "outbound report" in first_col_ob.lower():
                raw_ob = pd.read_excel(
                    excel_path, sheet_name=outbound1_name, engine="openpyxl", header=None
                )
                if raw_ob.empty or raw_ob.shape[0] < 2:
                    raw_ob.columns = raw_ob.columns.astype(str).str.strip()
                    df1 = raw_ob.copy()
                else:
                    header_row_idx_ob = None
                    for idx in range(min(10, raw_ob.shape[0])):
                        row = raw_ob.iloc[idx]
                        cells = " ".join(
                            str(c).strip().lower() for c in row.dropna().astype(str)
                        )
                        # نعتبر صف ترويسات لو فيه Order + (Create أو Shipped) (شيت ARAMCO Outbound Report)
                        if (
                            "order" in cells
                            and ("create" in cells or "shipped" in cells or "ship" in cells)
                            and ("status" in cells or "facility" in cells or "customer" in cells)
                        ):
                            header_row_idx_ob = idx
                            break
                    if header_row_idx_ob is not None:
                        df1 = raw_ob.iloc[header_row_idx_ob + 1 :].copy()
                        headers_ob = [
                            str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}"
                            for i, c in enumerate(raw_ob.iloc[header_row_idx_ob].values)
                        ]
                        df1.columns = headers_ob
                        df1 = df1.reset_index(drop=True)
                    else:
                        # fallback: استخدم أول صف كترويسة وتجاهل الصف ده من الداتا
                        headers_ob = [
                            str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}"
                            for i, c in enumerate(raw_ob.iloc[0].values)
                        ]
                        df1 = raw_ob.copy()
                        df1.columns = headers_ob
                        df1 = df1.iloc[1:].reset_index(drop=True)

            df1.columns = df1.columns.astype(str).str.strip()

            def find_col(df, candidates):
                for c in df.columns:
                    if str(c).strip().lower() in [x.lower() for x in candidates]:
                        return c
                for cand in candidates:
                    for c in df.columns:
                        if cand.lower() in str(c).lower():
                            return c
                return None

            def find_col_containing(df, *keywords):
                k_lower = [k.lower() for k in keywords]
                for c in df.columns:
                    col_lower = str(c).strip().lower()
                    if all(k in col_lower for k in k_lower):
                        return c
                return None

            # Outbound1 / ARAMCO Outbound Report — إيجاد الأعمدة (زي Inbound: بحث مرن)
            order_nbr_col = find_col(
                df1, ["Order Nbr", "Order Nbr.", "Order Number", "Order No", "Order #"]
            )
            if not order_nbr_col:
                order_nbr_col = find_col_containing(df1, "order", "nbr") or find_col_containing(df1, "order", "no")
            customer_col = find_col(df1, ["Customer Name", "Customer"])
            if not customer_col:
                customer_col = find_col_containing(df1, "customer", "name")
            # Create Order (من ARAMCO Outbound Report) = تاريخ إنشاء الطلب
            create_ts_col = find_col(
                df1,
                [
                    "Create Order",
                    "Create Timestamp",
                    "Create Date",
                    "Order Date",
                    "Created",
                ],
            )
            if not create_ts_col:
                create_ts_col = find_col_containing(df1, "create", "order") or find_col_containing(df1, "create", "date") or find_col_containing(df1, "create", "timestamp")
            city_col = find_col(df1, ["Customer City", "City"])
            order_type_col = find_col(df1, ["Order Type", "Type"])
            status_col = find_col(df1, ["Status"])
            if not status_col:
                status_col = find_col_containing(df1, "status")
            # Shipped date = تاريخ الشحن للمقارنة مع Create Order
            ship_date_col = find_col(
                df1,
                [
                    "Shipped Date",
                    "Ship Date",
                    "Shipment Date",
                    "Order checked",
                    "Order checked?",
                ],
            )
            if not ship_date_col:
                ship_date_col = find_col_containing(df1, "shipped", "date") or find_col_containing(df1, "ship", "date")
            facility_code_col = find_col(
                df1, ["Facility Code", "Facility Code.", "Facility"]
            )
            if not facility_code_col:
                facility_code_col = find_col_containing(df1, "facility")

            required_ob1 = [
                order_nbr_col,
                customer_col,
                create_ts_col,
                status_col,
                ship_date_col,
            ]
            if not all(required_ob1):
                missing = []
                if not order_nbr_col:
                    missing.append("Order Nbr")
                if not customer_col:
                    missing.append("Customer Name")
                if not create_ts_col:
                    missing.append("Create Timestamp / Create Order")
                if not status_col:
                    missing.append("Status")
                if not ship_date_col:
                    missing.append("Ship Date / Shipped date")

                actual_cols = ", ".join(str(c) for c in df1.columns.tolist()[:20])
                if len(df1.columns) > 20:
                    actual_cols += ", …"

                return {
                    "detail_html": (
                        f"<p class='text-danger'>⚠️ Outbound1: missing required columns: {', '.join(missing)}.</p>"
                        f"<p class='text-muted small mt-2'>الأعمدة المقروءة من الشيت '{outbound1_name}': {actual_cols}</p>"
                    ),
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                }

            rename_ob1 = {
                order_nbr_col: "Order Nbr",
                customer_col: "Customer Name",
                create_ts_col: "Create Timestamp",
                status_col: "Status",
                ship_date_col: "Ship Date",
            }
            if city_col and city_col in df1.columns:
                rename_ob1[city_col] = "Customer City"
            if order_type_col and order_type_col in df1.columns:
                rename_ob1[order_type_col] = "Order Type"
            if facility_code_col and facility_code_col in df1.columns:
                rename_ob1[facility_code_col] = "Facility Code"
            df1 = df1.rename(columns=rename_ob1)
            if "Customer City" not in df1.columns:
                df1["Customer City"] = ""
            if "Order Type" not in df1.columns:
                df1["Order Type"] = ""
            if "Facility Code" not in df1.columns:
                df1["Facility Code"] = ""

            # عمود Create Order و Ship Date: تحويل قوي لتاريخ (Excel serial، نصوص، صيغ متعددة)
            def _parse_outbound_date(val):
                if val is None or (isinstance(val, float) and (pd.isna(val) or val != val)):
                    return pd.NaT
                if hasattr(val, "date") and hasattr(val, "time"):
                    return pd.Timestamp(val)
                if isinstance(val, (int, float)):
                    try:
                        return pd.to_datetime(val, origin="1899-12-30", unit="D")
                    except Exception:
                        return pd.NaT
                s = str(val).strip()
                if not s or s.lower() in ("nan", "nat", "none", "#n/a", ""):
                    return pd.NaT
                if s.startswith("0 ") and len(s) > 2:
                    s = s[2:].strip()
                # Excel serial أحياناً يقرأ كنص (مثل "45321.0")
                try:
                    n = float(s)
                    if 10000 < n < 100000:
                        return pd.to_datetime(n, origin="1899-12-30", unit="D")
                except (ValueError, TypeError):
                    pass
                out = pd.to_datetime(s, errors="coerce", dayfirst=True)
                if pd.isna(out):
                    for fmt in (
                        "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d",
                        "%m/%d/%Y", "%m/%d/%Y %H:%M:%S", "%d-%b-%Y", "%d-%b-%y", "%b %d %Y", "%b %d, %Y",
                        "%d/%b/%Y", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y", "%d %b %Y",
                    ):
                        try:
                            out = pd.to_datetime(s, format=fmt, errors="coerce")
                            if pd.notna(out):
                                break
                        except Exception:
                            continue
                return out

            for dt_col in ["Create Timestamp", "Ship Date"]:
                if dt_col not in df1.columns:
                    continue
                raw_dates = df1[dt_col].copy()
                # أولاً: pandas قد يكون قرأ التواريخ تلقائياً من Excel
                df1[dt_col] = pd.to_datetime(df1[dt_col], errors="coerce")
                # ثانياً: أي قيمة لا تزال NaT نحاول تحويلها من القيمة الخام (نص، Excel serial، إلخ)
                nat_mask = df1[dt_col].isna()
                if nat_mask.any():
                    parsed = raw_dates.loc[nat_mask].apply(_parse_outbound_date)
                    df1.loc[nat_mask, dt_col] = parsed.values
                # ثالثاً: لو لا تزال NaT وقد تكون أرقام Excel serial كنص
                still_nat = df1[dt_col].isna()
                if still_nat.any():
                    raw_vals = raw_dates.loc[still_nat].astype(str)
                    numeric = pd.to_numeric(raw_vals.str.replace(",", "."), errors="coerce")
                    valid_serial = (numeric.notna()) & (numeric > 10000) & (numeric < 100000)
                    if valid_serial.any():
                        try:
                            conv = pd.to_datetime(numeric[valid_serial].values, origin="1899-12-30", unit="D")
                            df1.loc[still_nat, dt_col] = df1.loc[still_nat, dt_col].astype("object")
                            idx_ok = numeric.index[valid_serial]
                            df1.loc[idx_ok, dt_col] = conv
                            df1[dt_col] = pd.to_datetime(df1[dt_col], errors="coerce")
                        except Exception:
                            pass
            # تشخيص: عدد التواريخ الصالحة وعينة
            _create_ok = df1["Create Timestamp"].notna().sum()
            _ship_ok = df1["Ship Date"].notna().sum()
            print(
                f"\n[Outbound] تواريخ: Create Timestamp صالح={_create_ok}/{len(df1)}, Ship Date صالح={_ship_ok}/{len(df1)}"
            )
            if _create_ok > 0:
                print("  عينة Create Timestamp:", df1["Create Timestamp"].dropna().head(3).tolist())
            if _ship_ok > 0:
                print("  عينة Ship Date:", df1["Ship Date"].dropna().head(3).tolist())

            df1["Order Nbr"] = df1["Order Nbr"].astype(str).str.strip()
            df1["Status"] = df1["Status"].astype(str).str.strip()

            # مفتاح ربط موحّد (يحل اختلاف التنسيق مثل 001 vs 1)
            def _order_key(ser):
                def _norm(v):
                    s = str(v).strip()
                    try:
                        return str(int(float(s)))
                    except (ValueError, TypeError):
                        return s

                return ser.astype(str).str.strip().apply(_norm)

            # Outbound2: Packed Timestamp + key to join (Order Nbr)
            packed_series = None
            if outbound2_name:
                df2 = pd.read_excel(
                    excel_path, sheet_name=outbound2_name, engine="openpyxl"
                )
                df2.columns = df2.columns.astype(str).str.strip()
                order_nbr_col2 = find_col(
                    df2,
                    ["Order Nbr", "Order Nbr.", "Order Number", "Order No", "Order #"],
                )
                packed_col = find_col(
                    df2, ["Packed Timestamp", "Packed Date", "Packed", "Packed Time"]
                )
                if order_nbr_col2 and packed_col:
                    df2 = df2[[order_nbr_col2, packed_col]].copy()
                    df2.columns = ["Order Nbr", "Packed Timestamp"]
                    df2["Order Nbr"] = df2["Order Nbr"].astype(str).str.strip()
                    df2["Packed Timestamp"] = pd.to_datetime(
                        df2["Packed Timestamp"], errors="coerce"
                    )
                    # إزالة التكرار في Order Nbr عشان الـ map يشتغل (نحتفظ بأول صف لكل Order Nbr)
                    df2_unique = df2.drop_duplicates(subset=["Order Nbr"], keep="first")
                    packed_series = df2_unique.set_index("Order Nbr")[
                        "Packed Timestamp"
                    ]
                    df1["Packed Timestamp"] = df1["Order Nbr"].map(packed_series)
                    # طباعة عينة من Outbound2 للتأكد من الداتا
                    print("\n[Outbound2] عينة من الشيت — Order Nbr و Packed Timestamp:")
                    for idx, r in df2.head(8).iterrows():
                        print(
                            f"  Order Nbr: {r['Order Nbr']!r}  |  Packed: {r['Packed Timestamp']}"
                        )
                    print(f"  عدد صفوف Outbound2: {len(df2)}\n")
                    # لو معظم القيم فاضية، نجرب المفتاح الموحّد (مثلاً 1 و 001 يتطابقان)
                    if df1["Packed Timestamp"].notna().sum() < len(df1) // 2:
                        df2["_ok"] = _order_key(df2["Order Nbr"])
                        packed_by_ok = df2.drop_duplicates(
                            subset=["_ok"], keep="first"
                        ).set_index("_ok")["Packed Timestamp"]
                        df1["_ok"] = _order_key(df1["Order Nbr"])
                        df1["Packed Timestamp"] = df1["Packed Timestamp"].fillna(
                            df1["_ok"].map(packed_by_ok)
                        )
                        df1.drop(columns=["_ok"], inplace=True, errors="ignore")
                else:
                    df1["Packed Timestamp"] = pd.NaT
            else:
                df1["Packed Timestamp"] = pd.NaT

            if "Packed Timestamp" not in df1.columns:
                df1["Packed Timestamp"] = pd.NaT

            # طباعة في الترمينال: نتيجة الربط مع Outbound2
            packed_filled = df1["Packed Timestamp"].notna().sum()
            print("\n" + "=" * 70)
            print("Outbound — نتيجة الربط (Packed Timestamp من Outbound2)")
            print("=" * 70)
            print(f"  إجمالي الصفوف (Outbound1): {len(df1)}")
            print(f"  صفوف فيها Packed Timestamp غير فاضي: {packed_filled}")
            print(f"  صفوف فاضية (مفيش ربط): {len(df1) - packed_filled}")
            if outbound2_name:
                print(f"  شيت Outbound2 المستخدم: {outbound2_name}")
            else:
                print("  ⚠️ مفيش شيت Outbound2 تم استخدامه")
            print("  عينة من df1 (Order Nbr | Packed Timestamp | Ship Date):")
            for i, row in df1.head(10).iterrows():
                pt = row.get("Packed Timestamp")
                pt_str = str(pt)[:19] if pd.notna(pt) and pt is not pd.NaT else "(فاضي)"
                sd = row.get("Ship Date")
                sd_str = str(sd)[:19] if pd.notna(sd) and sd is not pd.NaT else "(فاضي)"
                print(
                    f"    Order Nbr: {row.get('Order Nbr')!r}  |  Packed: {pt_str}  |  Ship: {sd_str}"
                )
            print("=" * 70 + "\n")

            # لو لسه فاضي: نجرب نأخذ Packed من Outbound1 لو العمود موجود فيه
            if df1["Packed Timestamp"].isna().all():
                packed_in_ob1 = find_col(
                    df1, ["Packed Timestamp", "Packed Date", "Packed", "Packed Time"]
                )
                if packed_in_ob1 and packed_in_ob1 in df1.columns:
                    df1["Packed Timestamp"] = pd.to_datetime(
                        df1[packed_in_ob1], errors="coerce"
                    )

            # ========== Outbound Hit/Miss (working days excl. Friday) ==========
            # الفرق بأيام العمل بين Create Order و Shipped date (يوم الجمعة لا يُحسب).
            # ≤50 سطر (صف) للطلب → مسموح يوم واحد؛ >50 سطر → مسموح يومين. Hit لو الأيام المستخدمة ≤ المسموح.
            def _working_days_between_exclude_friday(create_ts, ship_ts):
                if create_ts is None or ship_ts is None or pd.isna(create_ts) or pd.isna(ship_ts):
                    return 0
                try:
                    start_d = create_ts.date() if hasattr(create_ts, "date") else create_ts
                    end_d = ship_ts.date() if hasattr(ship_ts, "date") else ship_ts
                except Exception:
                    return 0
                if start_d > end_d:
                    return 0
                days = 0
                current = start_d
                while current <= end_d:
                    if current.weekday() != 4:  # 4 = Friday — لا نحسب يوم الجمعة
                        days += 1
                    current += timedelta(days=1)
                return days

            line_count_per_order = df1.groupby("Order Nbr").size()
            create_min = df1.groupby("Order Nbr")["Create Timestamp"].min()
            ship_max = df1.groupby("Order Nbr")["Ship Date"].max()
            order_days_used = {}
            order_allowed_days = {}
            order_line_count = {}
            order_is_hit = {}
            for o in df1["Order Nbr"].unique():
                if not o or (isinstance(o, float) and pd.isna(o)):
                    continue
                line_count = int(line_count_per_order.get(o, 0))
                allowed_days = 1 if line_count <= 50 else 2
                create_ts = create_min.get(o)
                ship_ts = ship_max.get(o)
                if pd.isna(create_ts) or pd.isna(ship_ts):
                    order_is_hit[o] = False
                    order_days_used[o] = 0
                else:
                    days_used = _working_days_between_exclude_friday(create_ts, ship_ts)
                    order_days_used[o] = days_used
                    order_is_hit[o] = days_used <= allowed_days
                order_allowed_days[o] = allowed_days
                order_line_count[o] = line_count

            df1["_LineCount"] = df1["Order Nbr"].map(order_line_count)
            df1["_AllowedDays"] = df1["Order Nbr"].map(order_allowed_days)
            df1["_DaysUsed"] = df1["Order Nbr"].map(order_days_used)
            df1["Cycle Days"] = df1["_DaysUsed"].replace(np.nan, 0)
            df1["Cycle Hours"] = df1["Cycle Days"] * 24
            df1["Days_Used"] = df1["_DaysUsed"]
            df1["is_hit"] = df1["Order Nbr"].map(lambda o: order_is_hit.get(o, False))
            df1["HIT or MISS"] = np.where(df1["is_hit"], "Hit", "Miss")
            df1.loc[df1["Create Timestamp"].isna() | df1["Ship Date"].isna(), "HIT or MISS"] = "Miss"

            # لو كل الطلبات Miss غالباً التواريخ ما اتحولتش صح (Create Order أو Ship Date فاضية)
            _n_hit = (df1["HIT or MISS"] == "Hit").sum()
            _n_miss = (df1["HIT or MISS"] == "Miss").sum()
            _n_create_ok = df1["Create Timestamp"].notna().sum()
            _n_ship_ok = df1["Ship Date"].notna().sum()
            outbound_all_miss_dates_issue = (
                _n_miss > 0 and _n_hit == 0 and (_n_create_ok == 0 or _n_ship_ok == 0)
            )
            if outbound_all_miss_dates_issue:
                print(
                    "[Outbound] ⚠️ كل الطلبات Miss — تحقق من قراءة التواريخ: "
                    f"Create Timestamp صالح={_n_create_ok}, Ship Date صالح={_n_ship_ok}"
                )
            # تشخيص النتيجة النهائية
            n_orders = df1["Order Nbr"].nunique()
            n_orders_both_dates = sum(
                1 for o in df1["Order Nbr"].unique()
                if not (pd.isna(create_min.get(o)) or pd.isna(ship_max.get(o)))
            )
            print(
                f"[Outbound] نتيجة Hit/Miss: طلبات={n_orders}, منها لها تاريخان صالحان={n_orders_both_dates}, "
                f"Hit={_n_hit} (صف)، Miss={_n_miss} (صف)"
            )

            # الشهر من Create Order (عمود Create Timestamp) — زي Inbound: يوم/شهر/سنة
            df1["Month"] = df1["Create Timestamp"].dt.strftime("%b").replace("NaT", "").fillna("")

            # تطبيع الفاسيليتي إلى Riyadh / Dammam / Jeddah لاستخدامها في جداول المناطق
            def _norm_facility(val):
                v = (str(val) or "").strip().lower()
                if "riyadh" in v or v == "ruh":
                    return "Riyadh"
                if "dammam" in v or "damam" in v:
                    return "Dammam"
                if "jeddah" in v or "jdda" in v or "jedd" in v:
                    return "Jeddah"
                return None

            if "Facility Code" in df1.columns:
                df1["_FacilityNorm"] = df1["Facility Code"].apply(_norm_facility)
            else:
                df1["_FacilityNorm"] = None

            # فلتر Facility أولاً (من الطلب أو من المعامل)
            _fac = (str(selected_facility).strip() if selected_facility else "") or (
                (request.GET.get("outbound_facility") or "").strip()
                if request and getattr(request, "GET", None)
                else ""
            )
            if _fac:
                fn = df1["_FacilityNorm"].fillna("").astype(str).str.strip()
                fc = df1["Facility Code"].fillna("").astype(str).str.strip().str.lower()
                df1 = df1[(fn == _fac) | (fc == _fac.lower())]

            # فلتر الشهر (Create Order)
            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in selected_months_norm:
                        selected_months_norm.append(norm)
            selected_month_norm = (
                self.normalize_month_label(selected_month)
                if selected_month and not selected_months_norm
                else None
            )
            if selected_months_norm:
                df1 = df1[
                    df1["Month"]
                    .fillna("")
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
            elif selected_month_norm:
                df1 = df1[
                    df1["Month"].fillna("").str.lower() == selected_month_norm.lower()
                ]

            # فلتر Order Nbr (اختياري — لعرض طلب معين)
            _order_nbr = (selected_order_nbr or "").strip() if selected_order_nbr else ""
            if not _order_nbr and request and getattr(request, "GET", None):
                _order_nbr = (request.GET.get("outbound_order_nbr") or "").strip()
            if _order_nbr:
                df1 = df1[
                    df1["Order Nbr"].astype(str).str.strip().str.lower().str.contains(_order_nbr.lower(), na=False)
                ]

            if df1.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ No outbound records for the selected period.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                }

            # فقط صفوف فيها شهر صالح (عشان الشارت والـ KPI يظهروا صح)
            df_summary = df1[
                (df1["Month"].notna()) & (df1["Month"].astype(str).str.strip() != "")
            ].copy()
            if df_summary.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ No valid month values in outbound data.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                }

            def month_order_value(label):
                if not label:
                    return 999
                label = str(label).strip()[:3].title()
                for idx in range(1, 13):
                    if month_abbr[idx] == label:
                        return idx
                return 999

            total_per_month = (
                df_summary.groupby("Month")["Order Nbr"]
                .nunique()
                .reset_index(name="Total_Shipments")
            )
            hits_df = (
                df_summary[df_summary["is_hit"]]
                .groupby("Month")["Order Nbr"]
                .nunique()
                .reset_index(name="Hits")
            )
            summary_df = total_per_month.merge(hits_df, on="Month", how="left")
            summary_df["Hits"] = summary_df["Hits"].fillna(0).astype(int)
            summary_df["Misses"] = summary_df["Total_Shipments"] - summary_df["Hits"]
            summary_df["Hit %"] = (
                summary_df["Hits"]
                / summary_df["Total_Shipments"].replace(0, np.nan)
                * 100
            )
            summary_df["Hit %"] = summary_df["Hit %"].fillna(0).round(2)
            
            # حساب عدد الـ Facilities الفريدة لكل شهر
            facility_per_month = (
                df_summary.groupby("Month")["Facility Code"]
                .nunique()
                .reset_index(name="Facility_Count")
            )
            summary_df = summary_df.merge(facility_per_month, on="Month", how="left")
            summary_df["Facility_Count"] = summary_df["Facility_Count"].fillna(0).astype(int)
            
            summary_df = summary_df.sort_values(
                by="Month", key=lambda col: col.map(month_order_value)
            )

            months_with_miss = summary_df[summary_df["Misses"] > 0]["Month"].tolist()
            months_with_hit_only_ob = summary_df[summary_df["Misses"] == 0][
                "Month"
            ].tolist()
            # ترتيب الشهور في الجداول والشارت يكون زمنيًا فقط (Jan → Dec)،
            # وقائمة الأشهر التي بها Miss تُستخدم في العنوان فقط.
            ordered_months = summary_df["Month"].tolist()

            kpi_rows = []
            for _, row in summary_df.iterrows():
                m = row["Month"]
                kpi_rows.append(
                    {
                        "Month": m,
                        "Total Shipments": int(row["Total_Shipments"]),
                        "Hit (≤2d)": int(row["Hits"]),
                        "Miss (>2d)": int(row["Misses"]),
                        "Hit %": float(row["Hit %"]),
                        "Facility Count": int(row["Facility_Count"]),
                    }
                )

            pivot_cols = ["KPI"] + ordered_months
            if len(ordered_months) >= 2:
                pivot_cols.append("2025")

            hit_pct_row = {"KPI": "Hit %"}
            total_row = {"KPI": "Total Shipments"}
            hit_row = {"KPI": "Hit (≤2d)"}
            miss_row = {"KPI": "Miss (>2d)"}
            for m in ordered_months:
                r = next((x for x in kpi_rows if x["Month"] == m), None)
                if r:
                    total_val = int(r["Total Shipments"])
                    hit_val = int(r["Hit (≤2d)"])
                    miss_val = int(r["Miss (>2d)"])
                    total_row[m] = total_val
                    hit_row[m] = hit_val
                    miss_row[m] = miss_val
                    hit_pct_row[m] = (
                        int(round(hit_val / total_val * 100)) if total_val > 0 else 0
                    )
            if "2025" in pivot_cols:
                total_2025 = sum(r["Total Shipments"] for r in kpi_rows)
                hit_2025 = sum(r["Hit (≤2d)"] for r in kpi_rows)
                hit_pct_row["2025"] = (
                    int(round(hit_2025 / total_2025 * 100)) if total_2025 > 0 else 0
                )
                total_row["2025"] = int(sum(r["Total Shipments"] for r in kpi_rows))
                hit_row["2025"] = int(sum(r["Hit (≤2d)"] for r in kpi_rows))
                miss_row["2025"] = int(sum(r["Miss (>2d)"] for r in kpi_rows))

            # Total Shipments آخر صف في الجدول
            summary_data_pivot = [hit_pct_row, hit_row, miss_row, total_row]
            summary_columns = pivot_cols
            summary_data = summary_data_pivot

            # إحصائيات على مستوى الشحنة (Order Nbr): Hit / Miss (Pending معروضة كـ Miss)
            orders_df = df1.drop_duplicates(subset=["Order Nbr"], keep="first")
            overall_total = int(orders_df.shape[0])
            overall_hits = int((orders_df["HIT or MISS"] == "Hit").sum())
            overall_miss = int((orders_df["HIT or MISS"] == "Miss").sum())
            overall_failed = overall_miss
            overall_hit_pct = (
                round((overall_hits / overall_total) * 100, 2) if overall_total else 0
            )
            overall_miss_pct = (
                round((overall_miss / overall_total) * 100, 2) if overall_total else 0
            )
            overall_failed_pct = (
                round((overall_failed / overall_total) * 100, 2) if overall_total else 0
            )

            # سيتم بناء بيانات الشارت لاحقًا من جداول المناطق (Riyadh / Dammam / Jeddah)
            chart_data = []

            months_with_miss_label = (
                " — Months with Miss: " + ", ".join(months_with_miss)
                if months_with_miss
                else " — All months Hit"
            )
            summary_table = {
                "id": "sub-table-outbound-hit-summary",
                "title": "Outbound KPI ≤ 2d (working days excl. Friday)" + months_with_miss_label,
                "columns": summary_columns,
                "data": summary_data,
                "chart_data": chart_data,
                "months_with_miss": months_with_miss,
                "months_with_hit_only": months_with_hit_only_ob,
            }

            # ====== جداول KPI لكل مدينة (Riyadh / Dammam / Jeddah) بنفس فكرة Inbound FS ======
            FACILITIES_OB = ["Riyadh", "Dammam", "Jeddah"]
            facility_tables = []
            facility_stats = {}

            for f in FACILITIES_OB:
                fdf = df1[df1["_FacilityNorm"] == f].copy()
                if fdf.empty:
                    facility_stats[f] = {
                        "total": 0, "hit": 0, "miss": 0, "hit_pct": 0,
                        "by_month": {}, "ordered_months": [], "months_with_miss": [], "rows": [],
                    }
                    continue

                # صف واحد لكل Order Nbr (مثل Inbound)
                f_orders = fdf.drop_duplicates(subset=["Order Nbr"], keep="first")
                rows_f = []
                for _, r in f_orders.iterrows():
                    o = r.get("Order Nbr")
                    rows_f.append({
                        "Facility Code": f,
                        "Order Nbr": o,
                        "Line Count": int(r.get("_LineCount") or 0),
                        "Allowed Days": int(r.get("_AllowedDays") or 0),
                        "Days": int(r.get("_DaysUsed") or 0),
                        "HIT or MISS": r.get("HIT or MISS", "Miss"),
                        "Month": r.get("Month", ""),
                    })
                f_total_per_month = (
                    fdf.groupby("Month")["Order Nbr"]
                    .nunique()
                    .reset_index(name="Total_Shipments")
                )
                f_hits_df = (
                    fdf[fdf["is_hit"]]
                    .groupby("Month")["Order Nbr"]
                    .nunique()
                    .reset_index(name="Hits")
                )
                f_summary_df = f_total_per_month.merge(
                    f_hits_df, on="Month", how="left"
                )
                f_summary_df["Hits"] = f_summary_df["Hits"].fillna(0).astype(int)
                f_summary_df["Misses"] = (
                    f_summary_df["Total_Shipments"] - f_summary_df["Hits"]
                )
                f_summary_df["Hit %"] = (
                    f_summary_df["Hits"]
                    / f_summary_df["Total_Shipments"].replace(0, np.nan)
                    * 100
                )
                f_summary_df["Hit %"] = f_summary_df["Hit %"].fillna(0).round(2)

                f_summary_df = f_summary_df.sort_values(
                    by="Month", key=lambda col: col.map(month_order_value)
                )

                by_month = {}
                total_all = 0
                hits_all = 0
                misses_all = 0
                for _, row in f_summary_df.iterrows():
                    m = row["Month"]
                    total_val = int(row["Total_Shipments"])
                    hits_val = int(row["Hits"])
                    misses_val = int(row["Misses"])
                    hit_pct_val = (
                        int(round((hits_val / total_val) * 100)) if total_val > 0 else 0
                    )
                    by_month[m] = {
                        "total": total_val,
                        "hit": hits_val,
                        "miss": misses_val,
                        "hit_pct": hit_pct_val,
                    }
                    total_all += total_val
                    hits_all += hits_val
                    misses_all += misses_val

                hit_pct_all = (
                    round((hits_all / total_all) * 100, 2) if total_all else 0
                )
                ordered_months = list(f_summary_df["Month"].tolist())
                months_with_miss = [
                    m for m in ordered_months if by_month.get(m, {}).get("miss", 0) > 0
                ]

                facility_stats[f] = {
                    "total": total_all,
                    "hit": hits_all,
                    "miss": misses_all,
                    "hit_pct": hit_pct_all,
                    "by_month": by_month,
                    "ordered_months": ordered_months,
                    "months_with_miss": months_with_miss,
                    "rows": rows_f,
                }

            # الشهور الموحدة لكل المناطق
            all_months_ob = set()
            for s in facility_stats.values():
                all_months_ob.update((s.get("by_month") or {}).keys())
            ordered_months_overall_ob = sorted(
                all_months_ob, key=lambda x: month_order_value(x)
            )

            # شارت Hit %: لكل شهر ٣ أعمدة (Riyadh / Dammam / Jeddah)
            # ألوان الأعمدة (بالترتيب): أول لون #9084ad، الثاني #e8f1fb، الثالث #0d6efd
            facility_colors_ob = {
                "Riyadh": "#9084ad",
                "Dammam": "#e8f1fb",
                "Jeddah": "#538fe7",
            }
            chart_data = []
            if ordered_months_overall_ob:
                for f in FACILITIES_OB:
                    stats_f = facility_stats.get(f, {})
                    by_month = stats_f.get("by_month", {}) or {}
                    data_points = [
                        {
                            "label": m,
                            "y": by_month.get(m, {}).get("hit_pct", 0),
                        }
                        for m in ordered_months_overall_ob
                    ]
                    chart_data.append(
                        {
                            "type": "column",
                            "name": f"{f} Hit %",
                            "color": facility_colors_ob.get(f, "#007fa3"),
                            "valueSuffix": "%",
                            "related_table": "sub-table-outbound-facilities-hit",
                            "dataPoints": data_points,
                        }
                    )

            # جدول واحد موحّد: الصفوف = المناطق، الأعمدة = (Jan Hit / Jan Miss / Feb Hit / Feb Miss / ...)
            if ordered_months_overall_ob:
                pivot_cols = ["KPI"]
                for m in ordered_months_overall_ob:
                    pivot_cols.append(f"{m} Hit")
                    pivot_cols.append(f"{m} Miss")

                summary_rows = []
                for f in FACILITIES_OB:
                    stats_f = facility_stats.get(f, {})
                    by_month = stats_f.get("by_month", {}) or {}
                    row = {"KPI": f}
                    for m in ordered_months_overall_ob:
                        month_stats = by_month.get(m, {}) or {}
                        row[f"{m} Hit"] = month_stats.get("hit", 0)
                        row[f"{m} Miss"] = month_stats.get("miss", 0)
                    summary_rows.append(row)

                facility_tables.append(
                    {
                        "id": "sub-table-outbound-facilities-hit",
                        "title": "Outbound KPI ≤ 2d (working days excl. Friday) — Hit by Facility",
                        "columns": pivot_cols,
                        "data": summary_rows,
                        "chart_data": [],
                        "full_width": False,
                        "facility_name": "All Facilities",
                    }
                )

            # جدول Outbound Shipments Detail: صف واحد لكل طلب (Order Nbr) — + عمود Remark قابل للتحرير من الأدمن
            detail_columns_ob = ["Facility Code", "Order Nbr", "Line Count", "Allowed Days", "Days", "HIT or MISS", "Month", "Remark"]
            detail_rows_ob = []
            for f in FACILITIES_OB:
                for r in facility_stats.get(f, {}).get("rows", []):
                    detail_rows_ob.append({
                        "Facility Code": r.get("Facility Code", f),
                        "Order Nbr": r.get("Order Nbr"),
                        "Line Count": r.get("Line Count"),
                        "Allowed Days": r.get("Allowed Days"),
                        "Days": r.get("Days"),
                        "HIT or MISS": r.get("HIT or MISS", "Miss"),
                        "Month": r.get("Month", ""),
                        "Remark": "",
                    })

            # ربط ملاحظات Outbound من الأدمن (OutboundOrderRemark) حسب Order Nbr + Facility
            if detail_rows_ob:
                unique_pairs_ob = list(
                    dict.fromkeys(
                        [
                            (str(r.get("Order Nbr", "")).strip(), r.get("Facility Code"))
                            for r in detail_rows_ob
                            if r.get("Order Nbr") and r.get("Facility Code")
                        ]
                    )
                )
                if unique_pairs_ob:
                    remarks_qs_ob = OutboundOrderRemark.objects.filter(
                        order_nbr__in=[p[0] for p in unique_pairs_ob],
                        facility__in=[p[1] for p in unique_pairs_ob],
                    )
                    remarks_map_ob = {
                        (obj.order_nbr.strip(), obj.facility.strip()): (obj.remark or "")
                        for obj in remarks_qs_ob
                    }
                else:
                    remarks_map_ob = {}

                for row in detail_rows_ob:
                    on = str(row.get("Order Nbr", "")).strip()
                    fc = row.get("Facility Code", "")
                    row["Remark"] = remarks_map_ob.get((on, fc), "") if on and fc else ""
            facility_code_options = sorted(
                set(str(row.get("Facility Code", "")).strip() for row in detail_rows_ob if row.get("Facility Code") and str(row.get("Facility Code", "")).strip())
            )
            # تحويل الشهر لـ str وتصفية nan/float عشان sorted لا يرمي TypeError
            _month_set = set()
            for row in detail_rows_ob:
                m = row.get("Month")
                if m is None or (isinstance(m, float) and (pd.isna(m) or m != m)):
                    continue
                s = str(m).strip()
                if s:
                    _month_set.add(s)
            month_options = sorted(_month_set) if _month_set else []
            hit_miss_options = ["Hit", "Miss"]

            detail_table = {
                "id": "sub-table-outbound-detail",
                "title": "Outbound Shipments Detail",
                "columns": detail_columns_ob,
                "data": detail_rows_ob,
                "chart_data": [],
                "full_width": True,
                "filter_options": {
                    "facility_codes": facility_code_options,
                    "months": month_options,
                    "hit_miss": hit_miss_options,
                },
            }

            outbound_warning_html = ""
            if outbound_all_miss_dates_issue:
                outbound_warning_html = (
                    "<p class='alert alert-warning mb-2'>"
                    "⚠️ <strong>كل الطلبات تظهر Failed (100%)</strong> — السبب الأرجح: لم يتم قراءة تواريخ "
                    "<strong>Create Order</strong> أو <strong>Ship Date</strong> من الشيت بشكل صحيح. "
                    "تأكد أن العمودين موجودين في شيت ARAMCO Outbound Report وأن التاريخ بصيغة يوم/شهر/سنة (مثلاً 3/02/2026)."
                    "</p>"
                )

            return {
                "detail_html": outbound_warning_html,
                "sub_tables": [summary_table] + facility_tables + [detail_table],
                "chart_data": chart_data,
                "stats": {
                    "total": overall_total,
                    "hit": overall_hits,
                    "miss": overall_miss,
                    "failed": overall_failed,
                    "hit_pct": overall_hit_pct,
                    "miss_pct": overall_miss_pct,
                    "failed_pct": overall_failed_pct,
                    "target": 99,
                },
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error processing outbound shipments: {e}</p>",
                "sub_tables": [],
                "chart_data": [],
                "stats": {},
            }

    def filter_inventory(self, request, selected_month=None, selected_months=None):
        """
        تاب Inventory: يقرأ شيت "Inventory" — كروت أربعة (Total/Successful/Failed/Target)، شارت Hit %،
        وجدول تفاصيل مثل Inbound Shipments Detail مع بادج على Hit/Miss.
        """
        try:
            import os

            excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                    "tab_data": {"name": "Inventory"},
                }

            df = self.get_sheet_dataframe(request, "Inventory")
            if df is None or df.empty:
                xls = pd.ExcelFile(excel_path, engine="openpyxl")
                sheet_name = next(
                    (s for s in xls.sheet_names if (s or "").strip().lower() == "inventory"),
                    None,
                )
                if not sheet_name:
                    sheet_name = next(
                        (s for s in xls.sheet_names if "inventory" in (s or "").lower()),
                        None,
                    )
                if not sheet_name:
                    return {
                        "detail_html": "<p class='text-warning'>⚠️ Sheet 'Inventory' was not found.</p>",
                        "sub_tables": [],
                        "chart_data": [],
                        "stats": {},
                        "tab_data": {"name": "Inventory"},
                    }
                df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
            if df.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Inventory sheet is empty.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                    "stats": {},
                    "tab_data": {"name": "Inventory"},
                }

            def _norm(s):
                return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

            def _find_col(names):
                col_map = {_norm(c): c for c in df.columns}
                for n in names:
                    if _norm(n) in col_map:
                        return col_map[_norm(n)]
                for c in df.columns:
                    if any(_norm(x) in _norm(c) for x in names):
                        return c
                return None

            region_col = _find_col([
                "Region", "Area", "Facility", "Warehouse", "Location",
                "Site", "City", "Branch", "KPI", "Location Name",
                "منطقة", "الموقع", "المنطقة", "موقع", "فرع",
            ])
            if not region_col and len(df.columns) > 0:
                first_col = df.columns[0]
                first_vals = df[first_col].dropna().astype(str).str.strip().str.lower()
                if not first_vals.empty:
                    sample = " ".join(first_vals.head(5).tolist())
                    if any(x in sample for x in ("riyadh", "dammam", "jeddah", "central", "eastern", "western")):
                        region_col = first_col
                else:
                    region_col = first_col

            hit_miss_col = _find_col(["HIT or MISS", "Hit or Miss", "Hit or MISS", "Hit/Miss", "Status", "Hit - Miss"])
            if not hit_miss_col and region_col:
                for c in df.columns:
                    if c == region_col:
                        continue
                    cl = _norm(c)
                    if "hit" in cl or "miss" in cl:
                        hit_miss_col = c
                        break

            def _norm_region(val):
                v = (str(val) or "").strip().lower()
                if "riyadh" in v or v == "ruh":
                    return "Riyadh"
                if "dammam" in v or "damam" in v:
                    return "Dammam"
                if "jeddah" in v or "jedd" in v:
                    return "Jeddah"
                if "central" in v:
                    return "Riyadh"
                if "eastern" in v or "east" in v:
                    return "Dammam"
                if "western" in v or "west" in v:
                    return "Jeddah"
                return None

            FACILITIES = ["Riyadh", "Dammam", "Jeddah"]
            if hit_miss_col:
                df["_hm"] = df[hit_miss_col].astype(str).str.strip().str.lower()
                df["HIT or MISS"] = df["_hm"].map(lambda x: "Hit" if x == "hit" else ("Miss" if x == "miss" else (x.title() if x else "")))
                df = df.drop(columns=["_hm"], errors="ignore")
            else:
                df["HIT or MISS"] = ""

            if region_col:
                df["_RegionNorm"] = df[region_col].apply(_norm_region)
            else:
                df["_RegionNorm"] = None

            def _safe_val(val):
                if val is None or (isinstance(val, float) and (pd.isna(val) or val != val)):
                    return ""
                if isinstance(val, (pd.Timestamp, datetime.datetime)):
                    try:
                        return val.strftime("%Y-%m-%d %H:%M") if hasattr(val, "strftime") else str(val)
                    except Exception:
                        return str(val)
                return val

            columns = [c for c in df.columns if not str(c).startswith("_")]
            if "Status" not in columns and "HIT or MISS" in df.columns:
                columns.append("HIT or MISS")
            if hit_miss_col and hit_miss_col != "HIT or MISS" and hit_miss_col in columns:
                columns = [c for c in columns if c != hit_miss_col]
            columns = [c for c in columns if str(c).strip().lower() != "results"]
            col_rename = {}
            for c in columns:
                if "(Shortages or Excess)" in str(c):
                    col_rename[c] = str(c).replace("(Shortages or Excess)", "").strip()
            col_rename["HIT or MISS"] = "Status"
            if col_rename:
                columns = [col_rename.get(c, c) for c in columns]
            order_end = ["Total", "Hit", "Miss", "Status"]
            tail = [c for c in order_end if c in columns]
            if "Hit" not in tail:
                tail.insert(1, "Hit")
            others = [c for c in columns if c not in order_end]
            columns = others + tail

            def _strip_shortages(val):
                if val is None or not isinstance(val, str):
                    return val
                return val.replace("(Shortages or Excess)", "").strip() or val

            def _to_num(val):
                if val is None or val == "":
                    return None
                try:
                    return float(str(val).replace(",", "").strip())
                except (ValueError, TypeError):
                    return None

            rows = []
            rows_by_region = {"Riyadh": [], "Dammam": [], "Jeddah": []}
            for _, r in df.iterrows():
                row = {}
                for c in df.columns:
                    if str(c).startswith("_") or str(c).strip().lower() == "results":
                        continue
                    if hit_miss_col and c == hit_miss_col:
                        continue
                    out_key = col_rename.get(c, c)
                    if out_key not in columns:
                        continue
                    v = r.get(c, r.get(c) if c in r else "")
                    v = _safe_val(v)
                    if out_key == "Status" or "miss" in str(c).lower():
                        v = _strip_shortages(str(v))
                    row[out_key] = v
                if "Total" in row and "Status" in row:
                    total_num = _to_num(row["Total"])
                    hm_val = row["Status"]
                    hm_num = _to_num(hm_val)
                    if total_num is not None and hm_num is not None and total_num == hm_num:
                        row["Status"] = "Hit"
                if "Hit" not in row and "Status" in row:
                    row["Hit"] = row["Total"] if str(row.get("Status", "")).strip().lower() == "hit" else ""
                rows.append(row)
                rn = r.get("_RegionNorm")
                if rn in rows_by_region:
                    rows_by_region[rn].append(row)

            total = len(rows)
            hit_count = sum(1 for row in rows if str(row.get("Status", "")).strip().lower() == "hit")
            miss_count = sum(1 for row in rows if str(row.get("Status", "")).strip().lower() == "miss")
            hit_pct = round((hit_count / total) * 100, 2) if total else 0
            target = 99
            stats = {"total": total, "hit": hit_count, "miss": miss_count, "hit_pct": hit_pct, "target": target}

            chart_data = []
            if region_col and df["_RegionNorm"].notna().any():
                facility_colors = {"Riyadh": "#9084ad", "Dammam": "#e8f1fb", "Jeddah": "#538fe7"}
                for f in FACILITIES:
                    region_rows = rows_by_region.get(f, [])
                    tot_f = len(region_rows)
                    hit_f = sum(1 for row in region_rows if str(row.get("Status", "")).strip().lower() == "hit")
                    pct_f = round((hit_f / tot_f) * 100, 2) if tot_f else 0
                    chart_data.append({
                        "type": "column",
                        "name": f"{f} Hit %",
                        "color": facility_colors.get(f, "#74c0fc"),
                        "valueSuffix": "%",
                        "dataPoints": [{"label": "Hit %", "y": pct_f}],
                    })

            facility_codes = sorted(df[region_col].dropna().astype(str).str.strip().unique().tolist()) if region_col else ["All"]
            months = []
            month_col = _find_col(["Month", "month", "الشهر"])
            if month_col and month_col in df.columns:
                months = sorted(df[month_col].dropna().astype(str).str.strip().unique().tolist())
            sub_tables = [{
                "id": "sub-table-inventory-detail",
                "title": "Inventory",
                "columns": columns,
                "data": rows,
                "chart_data": [],
                "full_width": True,
                "facility_name": None,
                "filter_options": {
                    "facility_codes": facility_codes,
                    "months": months,
                    "hit_miss": ["Hit", "Miss"],
                    "facility_column": region_col or "",
                },
            }]
            tab_data = {
                "name": "Inventory",
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "stats": stats,
            }
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": None},
            )
            return {
                "detail_html": html,
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "stats": stats,
                "tab_data": tab_data,
            }
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error processing Inventory: {e}</p>",
                "sub_tables": [],
                "chart_data": [],
                "stats": {},
                "tab_data": {"name": "Inventory"},
            }

    def filter_inbound(self, request, selected_month=None, selected_months=None, tab_name=None):
        """
        Inbound tab: reads from ARAMCO Inbound Report. Excludes Shipment_nbr starting with 250 (Return).
        Return & Refusal tab: same sheet (ARAMCO Inbound Report), includes only Shipment_nbr starting with 250.
        Filter order (Return): Facility → Create shipment D&T (month/date) → Shipment_nbr (optional).
        Hit/Miss: ≤50 lines = 1 day allowed, >50 = 2 days. Return uses working days (excl. Friday) between
        Create shipment D&T and Received LPN D&T; Inbound uses calendar days.
        """
        try:
            import os

            # الملف الرئيسي لكل التابات (all_sheet / latest) مع أولوية للملف المرفوع في الجلسة
            excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = None
            is_return_tab = (tab_name or "").strip().lower() == "return & refusal"
            # Return tab: always read from ARAMCO Inbound Report (same sheet, filter to Shipment_nbr starting with 250 only)
            if is_return_tab:
                for s in xls.sheet_names:
                    if "ARAMCO Inbound Report" in (s or "").strip():
                        sheet_name = s
                        break
                if not sheet_name:
                    sheet_name = next((s for s in xls.sheet_names if "inbound" in (s or "").lower()), None)
            if not sheet_name and not is_return_tab:
                sheet_name = next((s for s in xls.sheet_names if (s or "").strip().lower() == "return & refusal"), None)
                if not sheet_name:
                    sheet_name = next((s for s in xls.sheet_names if "return" in (s or "").lower() and "refusal" in (s or "").lower()), None)
                if not sheet_name:
                    sheet_name = next((s for s in xls.sheet_names if "return" in (s or "").lower()), None)
            if not sheet_name:
                for s in xls.sheet_names:
                    if "ARAMCO Inbound Report" in (s or "").strip():
                        sheet_name = s
                        break
            if not sheet_name:
                sheet_name = next((s for s in xls.sheet_names if "inbound" in (s or "").lower()), None)
            if not sheet_name:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Sheet 'ARAMCO Inbound Report' or 'Inbound' (or 'Return' for Return tab) was not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            if df.empty:
                msg = "Return sheet is empty." if is_return_tab else "Inbound sheet is empty."
                return {
                    "detail_html": f"<p class='text-warning'>⚠️ {msg}</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            df.columns = df.columns.astype(str).str.strip()

            # إذا الصف الأول عنوان (مثل "ARAMCO Inbound Report" أو "Return & Refusal") والرؤوس في صف تالي، نكتشف صف الرؤوس
            first_col = str(df.columns[0]).strip() if len(df.columns) else ""
            need_header_detect = (
                first_col.startswith("Unnamed:")
                or first_col == "ARAMCO Inbound Report"
                or (first_col and "inbound report" in first_col.lower())
                or (is_return_tab and first_col and ("return" in first_col.lower() or "refusal" in first_col.lower()))
            )
            if need_header_detect:
                raw = pd.read_excel(
                    excel_path, sheet_name=sheet_name, engine="openpyxl", header=None
                )
                if raw.empty or raw.shape[0] < 2:
                    raw.columns = raw.columns.astype(str).str.strip()
                else:
                    header_row_idx = None
                    for idx in range(min(10, raw.shape[0])):
                        row = raw.iloc[idx]
                        cells = " ".join(str(c).strip().lower() for c in row.dropna().astype(str))
                        if (
                            "facility" in cells
                            and ("shipment" in cells or "shipment_nbr" in cells)
                            and ("create" in cells or "creation" in cells)
                            and ("received" in cells or "lpn" in cells)
                        ):
                            header_row_idx = idx
                            break
                        if is_return_tab and header_row_idx is None and (
                            ("facility" in cells or "region" in cells) and ("shipment" in cells or "order" in cells or "return" in cells)
                        ):
                            header_row_idx = idx
                            break
                    if header_row_idx is not None:
                        df = raw.iloc[header_row_idx + 1 :].copy()
                        headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[header_row_idx].values)]
                        df.columns = headers
                        df = df.reset_index(drop=True)
                    else:
                        df = raw.copy()
                        headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"Col_{i}" for i, c in enumerate(raw.iloc[0].values)]
                        df.columns = headers
                        df = df.iloc[1:].reset_index(drop=True)

            df.columns = df.columns.astype(str).str.strip()

            def normalize_name(val):
                return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())

            def find_column(possible_names):
                normalized_map = {normalize_name(col): col for col in df.columns}
                for name in possible_names:
                    norm = normalize_name(name)
                    if norm in normalized_map:
                        return normalized_map[norm]
                for col in df.columns:
                    col_norm = normalize_name(col)
                    if any(normalize_name(n) in col_norm for n in possible_names):
                        return col
                return None

            # مرادفات كثيرة لأن أسماء الأعمدة في الإكسل تختلف (مسافات، شرطات، رموز)
            col_facility = find_column([
                "Facility", "facility", "Facility Code", "facility code", "FacilityCode",
                "Site", "Warehouse", "Location", "Facility Name", "facility name",
                # شيت ARAMCO Inbound Report الجديد يستخدم Region بدل Facility
                "Region",
            ])
            col_shipment = find_column([
                "Shipment_nbr", "Shipment nbr", "Shipment Nbr", "Shipment No", "Shipment Number",
                "shipment number", "Shipment ID", "ShipmentID", "Shipment #", "Shipment#",
                "Shipment No.", "ShipmentNbr", "Shipment Nbr.",
                # شيت ARAMCO Inbound Report: Shipment_ID
                "Shipment_ID",
            ])
            col_create = find_column([
                "Create shipment D&T", "Create Shipment D&T", "Create shipment D&T", "Create Shipment D&T",
                "Create Timestamp", "create timestamp", "Creation Date", "Create Date", "Created Date",
                "Shipment Create Date", "Create Date & Time", "Create D&T", "Create DT",
                "Create shipement D&T", "Create shipemnt D&T", "Create Shipement D&T",
                # شيت ARAMCO Inbound Report: نعتبر Ship_Date هو تاريخ الإنشاء
                "Ship_Date",
            ])
            col_received = find_column([
                "Received LPN D&T", "Received LPN D&T", "Last LPN Rcv TS", "last lpn rcv ts",
                "Received LPN Date", "LPN Received Date", "Receipt Date", "Received Date",
                "Last LPN Receive", "LPN Rcv TS", "Received D&T", "Received DT",
                "Received LPN D&T", "Received LPN DT",
                # شيت ARAMCO Inbound Report: نستخدم Receiving_Complete_Date أو Verified_Date
                "Receiving_Complete_Date", "Verified_Date",
            ])
            col_lpn = find_column(["LPN", "LPN Nbr", "LPN_nbr", "LPNs", "LPN Number", "LPN No"])

            # بحث ثاني: أي عمود اسمه يحتوي على الكلمات المفتاحية (لو الأسماء مختلفة جداً)
            def find_column_containing(*keywords):
                k_norm = [normalize_name(k) for k in keywords]
                for col in df.columns:
                    c = normalize_name(col)
                    if all(k in c for k in k_norm):
                        return col
                return None

            if not col_facility:
                col_facility = find_column_containing("facility")
            if not col_shipment:
                col_shipment = find_column_containing("shipment", "nbr") or find_column_containing("shipment", "no") or find_column_containing("shipment", "number") or find_column_containing("shipment", "id")
            if not col_create:
                col_create = find_column_containing("create", "shipment") or find_column_containing("create", "timestamp") or find_column_containing("create", "date") or find_column_containing("creation", "date")
            if not col_received:
                col_received = find_column_containing("received", "lpn") or find_column_containing("lpn", "rcv") or find_column_containing("lpn", "receive") or find_column_containing("last", "lpn")

            if not col_facility or not col_shipment or not col_create or not col_received:
                missing = []
                if not col_facility:
                    missing.append("Facility")
                if not col_shipment:
                    missing.append("Shipment_nbr")
                if not col_create:
                    missing.append("Create shipment D&T")
                if not col_received:
                    missing.append("Received LPN D&T")
                actual_cols = ", ".join(str(c) for c in df.columns.tolist()[:20])
                if len(df.columns) > 20:
                    actual_cols += ", …"
                return {
                    "detail_html": (
                        f"<p class='text-danger'>⚠️ Missing required columns: {', '.join(missing)}</p>"
                        f"<p class='text-muted small mt-2'>الأعمدة الموجودة في الشيت: {actual_cols}</p>"
                    ),
                    "sub_tables": [],
                    "chart_data": [],
                }

            df = df.rename(columns={
                col_facility: "Facility",
                col_shipment: "Shipment_nbr",
                col_create: "Create_shipment_DT",
                col_received: "Received_LPN_DT",
            })
            if col_lpn:
                df = df.rename(columns={col_lpn: "LPN"})
            else:
                df["LPN"] = range(len(df))

            df["Facility"] = df["Facility"].astype(str).str.strip()
            df["Shipment_nbr"] = df["Shipment_nbr"].astype(str).str.strip()
            # Return tab: only Shipment_nbr starting with 250. Inbound: exclude 250 (Return shipments)
            if is_return_tab:
                df = df[df["Shipment_nbr"].str.startswith("250", na=False)].copy()
                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning'>⚠️ No Return rows (only shipments starting with 250 are shown).</p>",
                        "sub_tables": [],
                        "chart_data": [],
                    }
            else:
                exclude_250 = getattr(self, "INBOUND_EXCLUDE_250", True)
                if exclude_250:
                    df = df[~df["Shipment_nbr"].str.startswith("250", na=False)].copy()
                    if df.empty:
                        return {
                            "detail_html": "<p class='text-warning'>⚠️ No Inbound rows (shipments starting with 250 are excluded as Return).</p>",
                            "sub_tables": [],
                            "chart_data": [],
                        }
            df["Create_shipment_DT"] = pd.to_datetime(
                df["Create_shipment_DT"], errors="coerce", dayfirst=True
            )
            df["Received_LPN_DT"] = pd.to_datetime(
                df["Received_LPN_DT"], errors="coerce", dayfirst=True
            )

            FACILITIES = ["Riyadh", "Dammam", "Jeddah"]

            def norm_facility(val):
                v = (val or "").strip().lower()
                if not v:
                    return None
                # قيم مباشرة باسم المدينة أو كودها
                if "riyadh" in v or v == "ruh":
                    return "Riyadh"
                if "dammam" in v or "damam" in v:
                    return "Dammam"
                if "jeddah" in v or "jedd" in v:
                    return "Jeddah"
                # قيم Region (Central / Eastern / Western) نربطها بالمدن الرئيسية
                if "central" in v or "وسط" in v:
                    return "Riyadh"
                if "eastern" in v or "east" in v or "شرقي" in v:
                    return "Dammam"
                if "western" in v or "west" in v or "غربي" in v:
                    return "Jeddah"
                return None

            df["_FacilityNorm"] = df["Facility"].map(norm_facility)
            df = df[df["_FacilityNorm"].notna()].copy()
            df["Month"] = df["Create_shipment_DT"].dt.strftime("%b").fillna("")

            if df.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ No rows with Facility in (Riyadh, Dammam, Jeddah).</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            # Return tab: server-side filters — Facility → Create shipment D&T (month) → Shipment_nbr (optional)
            if is_return_tab and request and getattr(request, "GET", None):
                _fac = (request.GET.get("return_facility") or "").strip()
                if _fac:
                    fn = df["_FacilityNorm"].fillna("").astype(str).str.strip()
                    fc = df["Facility"].fillna("").astype(str).str.strip().str.lower()
                    df = df[(fn == _fac) | (fc == _fac.lower())]
                selected_months_norm = []
                if selected_months:
                    if isinstance(selected_months, str):
                        selected_months = [selected_months]
                    for m in selected_months:
                        norm = self.normalize_month_label(m)
                        if norm and norm not in selected_months_norm:
                            selected_months_norm.append(norm)
                month_norm = self.normalize_month_label(selected_month) if selected_month and not selected_months_norm else None
                if selected_months_norm:
                    df = df[df["Month"].fillna("").str.lower().isin([m.lower() for m in selected_months_norm])]
                elif month_norm:
                    df = df[df["Month"].fillna("").str.lower() == month_norm.lower()]
                _ship = (request.GET.get("return_shipment_nbr") or "").strip()
                if _ship:
                    df = df[df["Shipment_nbr"].astype(str).str.strip().str.lower().str.contains(_ship.lower(), na=False)]
                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning'>⚠️ No Return data for the selected filters.</p>",
                        "sub_tables": [],
                        "chart_data": [],
                    }

            def calendar_days_between(create_ts, received_ts):
                """
                Calendar days between Create shipment D&T and Received LPN D&T (Inbound).
                """
                if create_ts is None or received_ts is None or pd.isna(create_ts) or pd.isna(received_ts):
                    return 0
                start_d = create_ts.date() if hasattr(create_ts, "date") else create_ts
                end_d = received_ts.date() if hasattr(received_ts, "date") else received_ts
                if start_d > end_d:
                    return 0
                return (end_d - start_d).days

            def working_days_between_exclude_friday(create_ts, received_ts):
                """
                Working days between Create and Received, Friday not counted (Return tab).
                منطق الأيام يشبه Inbound: الفرق بعدد الأيام بين التاريخين (start → end),
                لكن نستبعد يوم الجمعة من العد.
                مثال: 9 Feb → 10 Feb = يوم واحد (لو اليومين ليسا جمعة).
                """
                if create_ts is None or received_ts is None or pd.isna(create_ts) or pd.isna(received_ts):
                    return 0
                try:
                    start_d = create_ts.date() if hasattr(create_ts, "date") else create_ts
                    end_d = received_ts.date() if hasattr(received_ts, "date") else received_ts
                except Exception:
                    return 0
                if start_d > end_d:
                    return 0
                days = 0
                current = start_d
                # نعد الأيام من start حتى قبل end (زي الفرق بالتواريخ)، مع استبعاد الجمعة
                while current < end_d:
                    if current.weekday() != 4:  # Friday
                        days += 1
                    current += timedelta(days=1)
                return days

            def month_order_value(label):
                if not label:
                    return 999
                label = (label or "").strip()[:3].title()
                for idx in range(1, 13):
                    if month_abbr[idx] == label:
                        return idx
                return 999

            def compute_facility_kpi(facility_name):
                fdf = df[df["_FacilityNorm"] == facility_name].copy()
                if fdf.empty:
                    return {"total": 0, "hit": 0, "miss": 0, "hit_pct": 0, "miss_pct": 0, "rows": [], "by_month": {}, "months_with_miss": [], "ordered_months": []}

                # عدد الـ Lines (الصفوف) لكل Shipment_nbr: ≤50 سطر = مسموح يوم واحد للاستلام (Hit لو استلمت في يوم)، >50 سطر = مسموح يومين (Miss لو أخذت يومين أو أكثر)
                line_count_per_shipment = fdf.groupby("Shipment_nbr").size()
                create_min = fdf.groupby("Shipment_nbr")["Create_shipment_DT"].min()
                received_max = fdf.groupby("Shipment_nbr")["Received_LPN_DT"].max()

                hits = 0
                misses = 0
                rows = []
                ship_month_hit = []
                for ship_id in fdf["Shipment_nbr"].unique():
                    if not ship_id:
                        continue
                    line_count = int(line_count_per_shipment.get(ship_id, 0))
                    # ≤50 سطر → الشحنة أخذت يوم كامل للاستلام → لو الفرق يوم واحد أو أقل = Hit. >50 سطر → أخذت يومين أو أكثر = Miss لو الفرق > يومين
                    allowed_days = 1 if line_count <= 50 else 2
                    create_ts = create_min.get(ship_id)
                    received_ts = received_max.get(ship_id)
                    if pd.isna(create_ts) or pd.isna(received_ts):
                        continue
                    month = create_ts.strftime("%b") if pd.notna(create_ts) else ""
                    # Return: working days (excl. Friday). Inbound: calendar days.
                    days_used = (
                        working_days_between_exclude_friday(create_ts, received_ts)
                        if is_return_tab
                        else calendar_days_between(create_ts, received_ts)
                    )
                    is_hit = days_used <= allowed_days
                    if is_hit:
                        hits += 1
                    else:
                        misses += 1
                    ship_month_hit.append((month, is_hit))
                    rows.append({
                        "Shipment_nbr": ship_id,
                        "Create_shipment_DT": create_ts,
                        "Received_LPN_DT": received_ts,
                        "Line Count": line_count,
                        "Allowed Days": allowed_days,
                        "Days": days_used,
                        "HIT or MISS": "Hit" if is_hit else "Miss",
                        "Month": month,
                    })

                by_month = {}
                for month, is_hit in ship_month_hit:
                    if not month:
                        continue
                    if month not in by_month:
                        by_month[month] = {"total": 0, "hit": 0, "miss": 0}
                    by_month[month]["total"] += 1
                    if is_hit:
                        by_month[month]["hit"] += 1
                    else:
                        by_month[month]["miss"] += 1
                # Hit % = (total hit / total shipment) * 100; Miss % = (total miss / total shipment) * 100
                for m in by_month:
                    t = by_month[m]["total"]
                    by_month[m]["hit_pct"] = round((by_month[m]["hit"] / t) * 100, 2) if t else 0
                    by_month[m]["miss_pct"] = round((by_month[m]["miss"] / t) * 100, 2) if t else 0

                ordered_months = sorted(by_month.keys(), key=lambda x: month_order_value(x))
                months_with_miss = [m for m in ordered_months if by_month[m]["miss"] > 0]
                total = hits + misses
                hit_pct = round((hits / total) * 100, 2) if total else 0   # Hit % = total hit / total shipment * 100
                miss_pct = round((misses / total) * 100, 2) if total else 0  # Miss % = total miss / total shipment * 100
                return {
                    "total": total, "hit": hits, "miss": misses, "hit_pct": hit_pct, "miss_pct": miss_pct, "rows": rows,
                    "by_month": by_month, "months_with_miss": months_with_miss, "ordered_months": ordered_months,
                }

            facility_stats = {f: compute_facility_kpi(f) for f in FACILITIES}

            # الشهور المجمعة عبر كل الفاسيليتيز (للاستخدام في الشارت والجدول)
            all_months = set()
            for s in facility_stats.values():
                by_month = s.get("by_month", {}) or {}
                all_months.update(by_month.keys())
            ordered_months_overall = sorted(all_months, key=lambda x: month_order_value(x))

            # الشارت: Hit % فقط (من إجمالي الشحنات)
            facility_colors = {
                "Riyadh": "#9084ad",
                "Dammam": "#e8f1fb",
                "Jeddah": "#538fe7",
            }
            chart_data = []
            if ordered_months_overall:
                for f in FACILITIES:
                    stats_f = facility_stats.get(f, {})
                    by_month = stats_f.get("by_month", {}) or {}
                    data_points_hit = [
                        {"label": m, "y": by_month.get(m, {}).get("hit_pct", 0)}
                        for m in ordered_months_overall
                    ]
                    chart_data.append(
                        {
                            "type": "column",
                            "name": f"{f} Hit %",
                            "color": facility_colors.get(f, "#74c0fc"),
                            "valueSuffix": "%",
                            "related_table": "sub-table-inbound-facilities-hit",
                            "dataPoints": data_points_hit,
                        }
                    )

            # جدول واحد كبير لكل المناطق:
            # الأعمدة = KPI (المناطق) + لكل شهر عمودين (Hit / Miss)
            # الصفوف = منطقة واحدة في كل صف (Riyadh / Dammam / Jeddah)
            facility_tables = []
            if ordered_months_overall:
                pivot_cols = ["KPI"]
                for m in ordered_months_overall:
                    pivot_cols.append(f"{m} Hit")
                    pivot_cols.append(f"{m} Miss")
                summary_rows = []

                for f in FACILITIES:
                    s = facility_stats[f]
                    by_month = s.get("by_month", {}) or {}
                    row = {"KPI": f}
                    for m in ordered_months_overall:
                        b = by_month.get(m, {}) or {}
                        row[f"{m} Hit"] = b.get("hit", 0)
                        row[f"{m} Miss"] = b.get("miss", 0)
                    summary_rows.append(row)

                facility_tables.append(
                    {
                        "id": "sub-table-inbound-facilities-hit",
                        "title": "Return KPI ≤ 2d (working days excl. Friday) — Hit by Facility" if is_return_tab else "Inbound KPI ≤ 24h — Hit % by Facility",
                        "columns": pivot_cols,
                        "data": summary_rows,
                        "chart_data": [],
                        "full_width": False,
                        "facility_name": "All Facilities",
                        "is_first_facility": True,
                    }
                )

            # Hit % = (total hit / total shipment) * 100; Miss % = (total miss / total shipment) * 100
            overall_total = sum(facility_stats[f]["total"] for f in FACILITIES)
            overall_hits = sum(facility_stats[f]["hit"] for f in FACILITIES)
            overall_miss = overall_total - overall_hits
            overall_hit_pct = round((overall_hits / overall_total) * 100, 2) if overall_total else 0   # total hit / total shipment * 100
            overall_miss_pct = round((overall_miss / overall_total) * 100, 2) if overall_total else 0   # total miss / total shipment * 100
            aggregated_kpi_rows = [
                {"KPI": "Hit %", "2025": round(overall_hit_pct, 2)},
                {"KPI": "Miss %", "2025": round(overall_miss_pct, 2)},
                {"KPI": "Hit (≤2d)", "2025": overall_hits},
                {"KPI": "Miss (>2d)", "2025": overall_miss},
                {"KPI": "Total Shipments", "2025": overall_total},
            ]
            aggregated_kpi_table = {
                "id": "inbound-aggregated-kpi",
                "title": "Return (250 only)" if is_return_tab else "Inbound (All)",
                "columns": ["KPI", "2025"],
                "data": aggregated_kpi_rows,
                "chart_data": [],
                "full_width": False,
                "facility_name": None,
                "is_first_facility": False,
            }

            # Detail table: one row per shipment — filters Facility / Month / Shipment_nbr / Hit or Miss
            # (Return tab shows Create shipment D&T and Received LPN D&T for date comparison as well)
            detail_columns = [
                "Facility Code",
                "Shipment_nbr",
                "Create shipment D&T",
                "Received LPN D&T",
                "Line Count",
                "Allowed Days",
                "Days",
                "HIT or MISS",
                "Month",
                "Remark",
            ]
            detail_rows = []
            for fac in FACILITIES:
                for r in facility_stats.get(fac, {}).get("rows", []):
                    detail_rows.append({
                        "Facility Code": fac,
                        "Shipment_nbr": r.get("Shipment_nbr"),
                        "Create shipment D&T": r.get("Create_shipment_DT"),
                        "Received LPN D&T": r.get("Received_LPN_DT"),
                        "Line Count": r.get("Line Count"),
                        "Allowed Days": r.get("Allowed Days"),
                        "Days": r.get("Days"),
                        "HIT or MISS": r.get("HIT or MISS"),
                        "Month": r.get("Month"),
                    })
            # Remark from InboundShipmentRemark (shared with Inbound)
            if detail_rows:
                unique_pairs = list(dict.fromkeys([(str(r.get("Shipment_nbr", "")).strip(), r.get("Facility Code")) for r in detail_rows if r.get("Shipment_nbr") and r.get("Facility Code")]))
                if unique_pairs:
                    remarks_qs = InboundShipmentRemark.objects.filter(
                        shipment_nbr__in=[p[0] for p in unique_pairs],
                        facility__in=[p[1] for p in unique_pairs],
                    )
                    remarks_map = {(obj.shipment_nbr.strip(), obj.facility.strip()): (obj.remark or "") for obj in remarks_qs}
                else:
                    remarks_map = {}
                for row in detail_rows:
                    sn = str(row.get("Shipment_nbr", "")).strip()
                    fc = row.get("Facility Code", "")
                    row["Remark"] = remarks_map.get((sn, fc), "") if sn and fc else ""

            facility_options = sorted(df["Facility"].dropna().unique().astype(str).tolist())
            month_options = sorted(set(r.get("Month") for r in detail_rows if r.get("Month")))
            detail_table = {
                "id": "sub-table-return-detail" if is_return_tab else "sub-table-inbound-detail",
                "title": "Return Shipments Detail (250 only)" if is_return_tab else "Inbound Shipments Detail",
                "columns": detail_columns,
                "data": detail_rows,
                "chart_data": [],
                "full_width": True,
                "filter_options": {
                    "facility_codes": facility_options,
                    "months": month_options,
                    "statuses": ["Hit", "Miss"],
                    "hit_miss": ["Hit", "Miss"],
                },
            }

            sub_tables = [aggregated_kpi_table] + facility_tables + [detail_table]

            from django.template.loader import render_to_string

            display_name = (tab_name or "Inbound").strip()
            tab_data = {
                "name": display_name,
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "stats": {
                    "total": overall_total,
                    "hit": overall_hits,
                    "miss": overall_miss,
                    "hit_pct": overall_hit_pct,
                    "miss_pct": overall_miss_pct,
                    "target": 99,
                },
            }
            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in seen:
                        seen.add(norm)
                        selected_months_norm.append(norm)
            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data,
                None if selected_months_norm else selected_month,
                selected_months_norm or None,
            )
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm_tab},
            )

            return {
                "detail_html": html,
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "stats": {
                    "total": overall_total,
                    "hit": overall_hits,
                    "miss": overall_miss,
                    "hit_pct": overall_hit_pct,
                    "miss_pct": overall_miss_pct,
                },
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing inbound data: {e}</p>",
                "sub_tables": [],
                "chart_data": [],
            }

    # Merge sheets from Excel
    def filter_pods_update(self, request, selected_month=None, selected_months=None):
        """
        تاب PODs: قراءة من شيت PODs (الشيت الجديد).

        المنطق الجديد (بدون الاعتماد على الشيت القديم):
        - فلترة أولاً بعمود Actual GI Date & Time (بالتاريخ / الشهر).
        - ثم فلترة بـ Delivery (من الدروب داون في الجدول).
        - لكل صف (Delivery في يوم معين) نحسب الفرق بين Actual GI Date & Time و POD REC DATE.
        - الفرق يُحسب بأيام عمل فقط (بدون يوم الجمعة).
        - Hit = خلال 7 أيام عمل أو أقل، Miss = أكثر من 7 أيام عمل.
        - لا توجد كروت KPI فوق التاب؛ فقط:
          * جدول KPI (Hit % / Hit / Miss / Total Shipments) لكل شهر + صف إجمالي.
          * شارت Hit % للشهور فقط (بدون مناطق).
          * جدول تفاصيل يعرض Actual GI Date & Time / Delivery / POD REC DATE / Days / Hit or Miss / Month.
        """
        import pandas as pd
        from django.template.loader import render_to_string
        import os
        from datetime import datetime, timedelta

        try:
            excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {"error": "⚠️ Excel file not found."}

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next(
                (s for s in xls.sheet_names if "pod" in (s or "").lower()),
                None,
            )
            if not sheet_name:
                return {"error": "⚠️ Sheet 'PODs' was not found."}

            df = pd.read_excel(
                excel_path,
                sheet_name=sheet_name,
                engine="openpyxl",
                dtype=str,
                header=0,
            ).fillna("")
            df.columns = df.columns.astype(str).str.strip()

            def _norm(val):
                return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())

            def _find_col(dframe, names):
                nmap = {_norm(c): c for c in dframe.columns}
                for name in names:
                    n = _norm(name)
                    if n in nmap:
                        return nmap[n]
                for col in dframe.columns:
                    if any(_norm(n) in _norm(col) for n in names):
                        return col
                return None

            def _find_col_containing(dframe, *keywords):
                """بحث إضافي: أي عمود يحتوي الكلمات المطلوبة بعد التطبيع."""
                ks = [_norm(k) for k in keywords]
                for col in dframe.columns:
                    c = _norm(col)
                    if all(k in c for k in ks):
                        return col
                return None

            # الأعمدة في الشيت الجديد
            col_gi = _find_col(
                df,
                [
                    "Actual GI Date & Time",
                    "Actual GI Date and Time",
                    "Actual GI Date",
                    "Actual GI D&T",
                    "Actual GI DT",
                    "GI Date & Time",
                    "GI Date Time",
                ],
            )
            col_delivery = _find_col(
                df,
                [
                    "Delivery",
                    "Delivery Number",
                    "Delivery No",
                    "Delivery Nbr",
                ],
            )
            col_pod_rec = _find_col(
                df,
                [
                    "POD REC DATE",
                    "POD Rec Date",
                    "POD Date",
                    "POD REC D&T",
                    "POD Date & Time",
                ],
            )

            # fallback لو الأسماء مختلفة قليلاً
            if not col_gi:
                col_gi = _find_col_containing(df, "actual", "gi", "date")
            if not col_pod_rec:
                col_pod_rec = _find_col_containing(df, "pod", "rec", "date")

            if not col_gi:
                return {"error": "⚠️ Required column 'Actual GI Date & Time' not found in PODs sheet."}
            if not col_pod_rec:
                return {"error": "⚠️ Required column 'POD REC DATE' not found in PODs sheet."}

            # تحويل التواريخ
            df["_gi_dt"] = pd.to_datetime(df[col_gi], errors="coerce")
            df["_pod_dt"] = pd.to_datetime(df[col_pod_rec], errors="coerce")

            # أيام عمل بين Actual GI Date & Time و POD REC DATE (بدون الجمعة)
            def business_days_between(start, end):
                if pd.isna(start) or pd.isna(end):
                    return None
                try:
                    s = pd.Timestamp(start).date()
                    e = pd.Timestamp(end).date()
                except Exception:
                    return None
                if s > e:
                    return None
                days = 0
                current = s
                # نعد الأيام من start حتى قبل end (زي الفرق بالتواريخ) مع استبعاد الجمعة
                while current < e:
                    if current.weekday() != 4:  # 4 = Friday
                        days += 1
                    current += timedelta(days=1)
                return days

            df["_work_days"] = df.apply(
                lambda row: business_days_between(row["_gi_dt"], row["_pod_dt"]),
                axis=1,
            )

            def _days_to_str(d):
                if d is None or (isinstance(d, float) and pd.isna(d)):
                    return ""
                try:
                    return str(int(float(d)))
                except (ValueError, TypeError):
                    return ""

            df["Days"] = df["_work_days"].apply(_days_to_str)

            def _hit_or_miss(d):
                if d is None or (isinstance(d, float) and pd.isna(d)):
                    return "Pending"
                try:
                    return "Hit" if float(d) <= 7 else "Miss"
                except (ValueError, TypeError):
                    return "Pending"

            df["Hit or Miss"] = df["_work_days"].apply(_hit_or_miss)
            df["Month"] = df["_gi_dt"].dt.strftime("%b").fillna("")

            # فلترة أولاً بـ Actual GI Date (الشهر)
            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm.lower() not in seen:
                        seen.add(norm.lower())
                        selected_months_norm.append(norm)

            if selected_months_norm:
                df = df[
                    df["Month"]
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
            elif selected_month:
                month_norm = self.normalize_month_label(selected_month)
                if month_norm:
                    df = df[df["Month"].str.lower() == month_norm.lower()]

            # فلترة بـ OBD Number (اختياري: نعرض كل الصفوف، الفلتر في الواجهة)
            # لا نستبعد أي صف هنا؛ OBD Number يكون متاحاً كفلتر في جدول التفاصيل

            if df.empty:
                return {
                    "detail_html": "<p class='text-warning text-center p-4'>⚠️ No data available for selected period.</p>",
                    "count": 0,
                    "hit_pct": 0,
                }

            month_order = [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ]
            months_raw = df["Month"].dropna().unique().tolist()
            # استبعاد القيم الفارغة أو غير المعرّفة (حتى لا يظهر عمود بدون اسم وقيم أصفار)
            months = [
                str(m).strip() for m in months_raw
                if m is not None and str(m).strip() and str(m).strip() in month_order
            ]
            months = sorted(months, key=lambda m: month_order.index(m))
            if not months:
                return {
                    "detail_html": "<p class='text-warning text-center p-4'>⚠️ No months found in data.</p>",
                    "count": 0,
                    "hit_pct": 0,
                }

            # إحصائيات عامة
            hit_count = len(df[df["Hit or Miss"] == "Hit"])
            miss_count = len(df[df["Hit or Miss"] == "Miss"])
            total_shipments = hit_count + miss_count
            hit_pct = (
                round((hit_count / total_shipments * 100), 2)
                if total_shipments > 0
                else 0
            )
            miss_pct = (
                round((miss_count / total_shipments * 100), 2)
                if total_shipments > 0
                else 0
            )

            # جدول KPI: Hit %, Hit, Miss, Total — أعمدة الشهور + 2025 (إجمالي السنة لكل صف)
            hit_by_month = []
            miss_by_month = []
            total_by_month = []
            hit_pct_by_month = []
            for month in months:
                df_m = df[df["Month"] == month]
                h = len(df_m[df_m["Hit or Miss"] == "Hit"])
                m = len(df_m[df_m["Hit or Miss"] == "Miss"])
                t = h + m
                hit_by_month.append(h)
                miss_by_month.append(m)
                total_by_month.append(t)
                hit_pct_by_month.append(
                    int(round((h / t * 100))) if t > 0 else 0
                )

            ytd_hit = sum(hit_by_month)
            ytd_miss = sum(miss_by_month)
            ytd_total = ytd_hit + ytd_miss
            ytd_hit_pct = int(round((ytd_hit / ytd_total * 100))) if ytd_total > 0 else 0

            # الأعمدة: KPI + الشهور (Jan / Feb / Mar ...) + عمود إجمالي 2025
            kpi_columns = ["KPI"] + months + ["2025"]
            kpi_rows = [
                {
                    "KPI": "Hit %",
                    **{m: hit_pct_by_month[i] for i, m in enumerate(months)},
                    "2025": ytd_hit_pct,
                },
                {
                    "KPI": "Hit",
                    **{m: hit_by_month[i] for i, m in enumerate(months)},
                    "2025": ytd_hit,
                },
                {
                    "KPI": "Miss",
                    **{m: miss_by_month[i] for i, m in enumerate(months)},
                    "2025": ytd_miss,
                },
                {
                    "KPI": "Total Shipments",
                    **{m: total_by_month[i] for i, m in enumerate(months)},
                    "2025": ytd_total,
                },
            ]

            # شارت: Hit % لكل شهر
            chart_data = [
                {
                    "type": "column",
                    "name": "Hit %",
                    "color": "#007fa3",
                    "valueSuffix": "%",
                    "related_table": "sub-table-pods-kpi",
                    "dataPoints": [
                        {"label": m, "y": hit_pct_by_month[i]}
                        for i, m in enumerate(months)
                    ],
                },
            ]

            sub_tables = [
                {
                    "id": "sub-table-pods-kpi",
                    "title": "PODs KPI — Hit & Miss",
                    "columns": kpi_columns,
                    "data": kpi_rows,
                    "chart_data": chart_data,
                }
            ]

            # جدول التفاصيل: Actual GI Date & Time / Delivery / POD REC DATE / Days / Hit or Miss / Month (+ أي أعمدة أخرى)
            detail_df = df.copy()

            def _fmt_date(x):
                if pd.isna(x) or x is pd.NaT:
                    return ""
                try:
                    return pd.Timestamp(x).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    return ""

            # صياغة الأعمدة الزمنية للعرض
            detail_df[col_gi] = detail_df["_gi_dt"].apply(_fmt_date)
            detail_df[col_pod_rec] = detail_df["_pod_dt"].apply(_fmt_date)

            # إسقاط الأعمدة المساعدة
            detail_df = detail_df.drop(
                columns=[c for c in ["_gi_dt", "_pod_dt", "_work_days"] if c in detail_df.columns],
                errors="ignore",
            )

            base_cols = [col_gi]
            if col_delivery:
                base_cols.append(col_delivery)
            base_cols.append(col_pod_rec)
            for add in ["Days", "Hit or Miss", "Month"]:
                if add not in base_cols:
                    base_cols.append(add)

            # تأكد أن الأعمدة موجودة في الداتا
            detail_columns = [c for c in base_cols if c in detail_df.columns]
            if not detail_columns:
                detail_columns = list(detail_df.columns)

            # ترتيب حسب Actual GI Date
            sort_col = col_gi if col_gi in detail_df.columns else detail_columns[0]
            detail_df = detail_df.sort_values(sort_col, ascending=False)
            detail_rows_raw = detail_df.head(500)[detail_columns].to_dict(orient="records")

            def _to_blank(val):
                if val is None:
                    return ""
                if isinstance(val, float) and (pd.isna(val) or (val != val)):
                    return ""
                s = str(val).strip()
                if s.lower() in ("nan", "nat", "none", "<nat>"):
                    return ""
                return s

            detail_rows = [
                {k: _to_blank(v) for k, v in row.items()} for row in detail_rows_raw
            ]

            delivery_options = (
                sorted(
                    detail_df[col_delivery]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace("", None)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if col_delivery and col_delivery in detail_df.columns
                else []
            )
            status_options = (
                sorted(
                    detail_df["Hit or Miss"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace("", None)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Hit or Miss" in detail_df.columns
                else []
            )
            month_options = (
                sorted(
                    detail_df["Month"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace("", None)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Month" in detail_df.columns
                else []
            )

            detail_table = {
                "id": "sub-table-pods-detail",
                "title": "PODs — Excel  Details",
                "columns": detail_columns,
                "data": detail_rows,
                "chart_data": [],
                # لا نعرض جدول التفاصيل كجدول رئيسي (إخفاؤه من التاب حسب طلب المستخدم)
                "full_width": False,
                "filter_options": {
                    "deliveries": delivery_options,
                    "delivery_column": col_delivery or "Delivery",
                    "statuses": status_options,
                    "months": month_options,
                },
            }
            # لا نضيف detail_table إلى sub_tables لعدم عرضه في التاب
            # sub_tables.append(detail_table)

            # كروت KPI
            stats = {
                "total_shipments": total_shipments,
                "hit_pct": hit_pct,
                "miss_pct": miss_pct,
                "target": 100,
            }

            tab_data = {
                "name": "PODs Update",
                "sub_tables": sub_tables,
                "chart_data": sub_tables[0].get("chart_data", []) if sub_tables else [],
                "chart_title": "PODs Hit % by Month",
                "hit_pct": hit_pct,
                "target_pct": 100,
                "stats": stats,
            }

            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data, selected_month, selected_months_norm or None
            )
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm_tab},
            )

            return {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "PODs Closed % Performance",
                "hit_pct": hit_pct,
                "target_pct": 100,
                "count": total_shipments,
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {"error": f"⚠️ Error processing PODs: {e}"}

    def filter_rejections_combined(
        self, request, selected_month=None, selected_months=None
    ):
        """
        تاب Return & Refusal: عرض جدول Return فقط من شيت Inbound (Shipment Type = RMA).
        بدون Rejection / Rejection breakdown / شارت — جدول فقط بعرض الصفحة.
        """
        import pandas as pd
        import os
        from django.template.loader import render_to_string

        try:
            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_names = [s.strip() for s in xls.sheet_names]
            sub_tables = []
            chart_data = []
            return_chart_data = []

            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in seen:
                        seen.add(norm)
                        selected_months_norm.append(norm)

            # ✅ جدول Return فقط من شيت Inbound: فلتر Shipment Type = RMA
            return_columns_display = [
                "Shipment Nbr",
                "Shipment Type",
                "Status",
                "Create Timestamp",
                "Arrival Date",
                "Offloading Date",
                "Last LPN Rcv TS",
            ]

            def _normalize_col(val):
                return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())

            def _find_col(df, possible_names):
                norm_map = {_normalize_col(c): c for c in df.columns}
                for name in possible_names:
                    n = _normalize_col(name)
                    if n in norm_map:
                        return norm_map[n]
                for col in df.columns:
                    if any(
                        _normalize_col(name) in _normalize_col(col)
                        for name in possible_names
                    ):
                        return col
                return None

            inbound_sheet = next(
                (s for s in sheet_names if "inbound" in s.lower()), None
            )
            if inbound_sheet:
                try:
                    df_in = pd.read_excel(
                        excel_path,
                        sheet_name=inbound_sheet,
                        engine="openpyxl",
                        dtype=str,
                        header=0,
                    ).fillna("")
                    df_in.columns = df_in.columns.astype(str).str.strip()

                    col_ship_nbr = _find_col(
                        df_in, ["shipment nbr", "shipment number", "shipment no"]
                    )
                    col_ship_type = _find_col(
                        df_in, ["shipment type", "shipmenttype", "type"]
                    )
                    col_status = _find_col(df_in, ["status", "shipment status"])
                    col_create = _find_col(
                        df_in, ["create timestamp", "created timestamp"]
                    )
                    col_arrival = _find_col(
                        df_in, ["arrival date", "arrival timestamp"]
                    )
                    col_offload = _find_col(df_in, ["offloading date", "offload date"])
                    col_last_lpn = _find_col(
                        df_in, ["last lpn rcv ts", "last lpn receive ts"]
                    )

                    if col_ship_type is not None:
                        df_in = df_in[
                            df_in[col_ship_type].astype(str).str.strip().str.upper()
                            == "RMA"
                        ]
                    else:
                        df_in = df_in.iloc[0:0]

                    if not df_in.empty and all(
                        [
                            col_ship_nbr,
                            col_status,
                            col_create,
                            col_arrival,
                            col_offload,
                            col_last_lpn,
                        ]
                    ):
                        rename = {
                            col_ship_nbr: "Shipment Nbr",
                            col_status: "Status",
                            col_create: "Create Timestamp",
                            col_arrival: "Arrival Date",
                            col_offload: "Offloading Date",
                            col_last_lpn: "Last LPN Rcv TS",
                        }
                        if col_ship_type is not None:
                            rename[col_ship_type] = "Shipment Type"
                        df_in = df_in.rename(columns=rename)

                        for c in return_columns_display:
                            if c not in df_in.columns:
                                df_in[c] = ""

                        ts_create = pd.to_datetime(
                            df_in["Create Timestamp"], errors="coerce"
                        )
                        ts_last = pd.to_datetime(
                            df_in["Last LPN Rcv TS"], errors="coerce"
                        )
                        hours = (ts_last - ts_create).dt.total_seconds() / 3600.0
                        df_in["_is_hit"] = (hours <= 24) & (hours.notna())
                        df_in["_month"] = ts_create.dt.strftime("%b")

                        col_fac = _find_col(
                            df_in, ["facility", "facility code", "منشأة"]
                        )
                        if col_fac and col_fac in df_in.columns:
                            def _norm_fac(v):
                                v = str(v).strip().lower()
                                if "central" in v or "riyadh" in v or "الرياض" in v:
                                    return "Riyadh"
                                if "eastern" in v or "dammam" in v or "الدمام" in v:
                                    return "Dammam"
                                if "western" in v or "jeddah" in v or "جدة" in v:
                                    return "Jeddah"
                                return v.title() if v else ""
                            df_in["_FacilityNorm"] = df_in[col_fac].fillna("").apply(_norm_fac)
                        else:
                            df_in["_FacilityNorm"] = "Riyadh"

                        return_chart_data = []
                        FACILITIES_R = ["Riyadh", "Dammam", "Jeddah"]
                        month_abbr = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                                      7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
                        ordered_months_ret = sorted(
                            df_in["_month"].dropna().unique().tolist(),
                            key=lambda m: next((k for k, v in month_abbr.items() if v == m), 99),
                        )[:3]
                        if ordered_months_ret:
                            facility_colors_r = {
                                "Riyadh": "#9084ad",
                                "Dammam": "#e8f1fb",
                                "Jeddah": "#538fe7",
                            }
                            for f in FACILITIES_R:
                                by_m = {}
                                sub = df_in[df_in["_FacilityNorm"] == f]
                                for m in ordered_months_ret:
                                    sm = sub[sub["_month"] == m]
                                    tot = len(sm)
                                    hit = int(sm["_is_hit"].sum())
                                    by_m[m] = (
                                        round(100.0 * hit / tot, 2) if tot else 0
                                    )
                                return_chart_data.append({
                                    "type": "column",
                                    "name": f"{f} Hit %",
                                    "color": facility_colors_r.get(f, "#74c0fc"),
                                    "valueSuffix": "%",
                                    "dataPoints": [{"label": m, "y": by_m.get(m, 0)} for m in ordered_months_ret],
                                })
                        df_in = df_in.drop(
                            columns=["_is_hit", "_month", "_FacilityNorm"],
                            errors="ignore",
                        )
                        df_in = df_in[return_columns_display]
                        if selected_month or selected_months_norm:
                            month_col = _find_col(df_in, ["month", "create timestamp"])
                            if month_col and month_col in df_in.columns:
                                if selected_months_norm:
                                    active = {
                                        self.normalize_month_label(m)
                                        for m in selected_months_norm
                                    }
                                else:
                                    active = {
                                        self.normalize_month_label(selected_month)
                                    }
                                if "Create Timestamp" in df_in.columns:
                                    try:
                                        ts = pd.to_datetime(
                                            df_in["Create Timestamp"],
                                            errors="coerce",
                                        )
                                        df_in["_month"] = ts.dt.strftime("%b")
                                        df_in = df_in[
                                            df_in["_month"]
                                            .fillna("")
                                            .str.lower()
                                            .isin([m.lower() for m in active])
                                        ]
                                        df_in = df_in.drop(
                                            columns=["_month"], errors="ignore"
                                        )
                                    except Exception:
                                        pass

                        # حساب Hit/Miss للـ Return (≤24h بين Create Timestamp و Last LPN Rcv TS)
                        return_kpi = None
                        try:
                            ts_create = pd.to_datetime(
                                df_in["Create Timestamp"], errors="coerce"
                            )
                            ts_last = pd.to_datetime(
                                df_in["Last LPN Rcv TS"], errors="coerce"
                            )
                            hours = (ts_last - ts_create).dt.total_seconds() / 3600.0
                            df_in["_is_hit"] = (hours <= 24) & (hours.notna())
                            total_ret = len(df_in)
                            successful_ret = int(df_in["_is_hit"].sum())
                            failed_ret = total_ret - successful_ret
                            hit_pct_ret = (
                                round(100.0 * successful_ret / total_ret, 2)
                                if total_ret else 0
                            )
                            return_kpi = {
                                "total_shipments": total_ret,
                                "successful": successful_ret,
                                "failed": failed_ret,
                                "target": 99,
                                "hit_pct": hit_pct_ret,
                            }
                            df_in = df_in.drop(columns=["_is_hit"], errors="ignore")
                        except Exception:
                            total_ret = len(df_in)
                            return_kpi = {
                                "total_shipments": total_ret,
                                "successful": total_ret,
                                "failed": 0,
                                "target": 99,
                                "hit_pct": 100.0 if total_ret else 0,
                            }

                        sub_tables.append(
                            {
                                "title": "Return",
                                "columns": return_columns_display,
                                "data": df_in.to_dict(orient="records"),
                                "return_kpi": return_kpi,
                            }
                        )
                    else:
                        sub_tables.append(
                            {
                                "title": "Return",
                                "columns": return_columns_display,
                                "data": [],
                                "error": (
                                    "Inbound sheet missing required columns or no RMA rows."
                                    if col_ship_type is not None
                                    else "Column 'Shipment Type' not found in Inbound."
                                ),
                            }
                        )
                except Exception as e_in:
                    import traceback

                    print(traceback.format_exc())
                    sub_tables.append(
                        {
                            "title": "Return",
                            "columns": return_columns_display,
                            "data": [],
                            "error": str(e_in),
                        }
                    )
            else:
                sub_tables.append(
                    {
                        "title": "Return",
                        "columns": return_columns_display,
                        "data": [],
                        "error": "Sheet containing 'Inbound' was not found.",
                    }
                )

            # ✅ التحقق من وجود بيانات بعد الفلترة
            total_count = sum(len(st["data"]) for st in sub_tables)
            if (selected_month or selected_months_norm) and total_count == 0:
                if selected_months_norm:
                    msg = ", ".join(selected_months_norm)
                else:
                    msg = str(selected_month).strip().capitalize()
                return {
                    "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for {msg} in Return & Refusal.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            # 🧩 بناء الـ HTML — نمرّر return_kpi من أول sub_table للكروت فوق الجدول
            return_kpi_for_tab = None
            if sub_tables:
                return_kpi_for_tab = sub_tables[0].get("return_kpi")
            tab_data = {
                "name": "Return & Refusal",
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "chart_title": "Return & Refusal Overview",
                "return_kpi": return_kpi_for_tab,
            }
            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data,
                selected_month if not selected_months_norm else None,
                selected_months_norm or None,
            )
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm_tab},
            )

            # 🧮 Hit %: نفضل نسبة الـ Return KPI (≤24h) إن وُجدت، وإلا من % of Rejection
            hit_pct = 0
            if return_kpi_for_tab and return_kpi_for_tab.get("hit_pct") is not None:
                hit_pct = round(float(return_kpi_for_tab["hit_pct"]), 2)
            else:
                hit_values = []
                for st in sub_tables:
                    if "rejection" in st["title"].lower():
                        for row in st["data"]:
                            val = row.get("% of Rejection", "")
                            try:
                                num = to_percentage_number(val)
                                if num is not None:
                                    hit_values.append(num)
                            except Exception:
                                pass
                hit_pct = round(sum(hit_values) / len(hit_values), 2) if hit_values else 0

            if not chart_data and total_count > 0:
                chart_data = [
                    {
                        "type": "bar",
                        "name": "Hit %",
                        "dataPoints": [{"label": "Return & Refusal", "y": hit_pct}],
                    }
                ]
            if return_chart_data:
                chart_data = return_chart_data

            result = {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "Return & Refusal Overview",
                "count": total_count,
                "hit_pct": hit_pct,
                "tab_data": tab_data,
            }

            return result

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing Return & Refusal data: {e}</p>",
                "chart_data": [],
                "count": 0,
            }

    def filter_expiry(self, request, selected_month=None, selected_months=None):
        """
        تاب Expiry: قراءة من شيت Expiry.
        - فلتر Status: Located, Allocated, Partly Allocated فقط.
        - أعمدة: Facility, Company, LPN Nbr, Status, Item Code, Item Description, Current Qty, batch_nbr, Expiry Date.
        - تحذير: اللي ينتهي خلال 3 شهور = قريب، خلال 6 شهور = warning، يعرض تحت الجدول في Bootstrap 5 alert.
        """
        import pandas as pd
        import os
        from datetime import datetime, timedelta
        from django.template.loader import render_to_string

        try:
            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_names = [s.strip() for s in xls.sheet_names]
            expiry_sheet = next((s for s in sheet_names if "expiry" in s.lower()), None)
            if not expiry_sheet:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Sheet containing 'Expiry' was not found.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            df = pd.read_excel(
                excel_path,
                sheet_name=expiry_sheet,
                engine="openpyxl",
                dtype=str,
                header=0,
            ).fillna("")
            df.columns = df.columns.astype(str).str.strip()

            def _norm(val):
                return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())

            def _find_col(dframe, names):
                nmap = {_norm(c): c for c in dframe.columns}
                for name in names:
                    n = _norm(name)
                    if n in nmap:
                        return nmap[n]
                for col in dframe.columns:
                    if any(_norm(n) in _norm(col) for n in names):
                        return col
                return None

            col_facility = _find_col(df, ["facility", "facility code"])
            col_company = _find_col(df, ["company"])
            col_lpn = _find_col(df, ["lpn nbr", "lpn", "lpn nbr"])
            col_status = _find_col(df, ["status"])
            col_item_code = _find_col(df, ["item code", "itemcode"])
            col_item_desc = _find_col(df, ["item description", "item desc"])
            col_qty = _find_col(df, ["current qty", "currentqty", "qty"])
            col_batch = _find_col(df, ["batch_nbr", "batch nbr", "batch"])
            col_expiry = _find_col(df, ["expiry date", "expirydate", "expiry"])

            if not col_status:
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Column 'Status' not found in Expiry sheet.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            # فلتر Status: Located, Allocated, Partly Allocated
            status_vals = {"located", "allocated", "partly allocated"}
            df = df[
                df[col_status].astype(str).str.strip().str.lower().isin(status_vals)
            ]

            display_columns = [
                "Facility",
                "Company",
                "LPN Nbr",
                "Status",
                "Item Code",
                "Item Description",
                "Current Qty",
                "batch_nbr",
                "Expiry Date",
            ]
            rename_map = {}
            if col_facility:
                rename_map[col_facility] = "Facility"
            if col_company:
                rename_map[col_company] = "Company"
            if col_lpn:
                rename_map[col_lpn] = "LPN Nbr"
            if col_status:
                rename_map[col_status] = "Status"
            if col_item_code:
                rename_map[col_item_code] = "Item Code"
            if col_item_desc:
                rename_map[col_item_desc] = "Item Description"
            if col_qty:
                rename_map[col_qty] = "Current Qty"
            if col_batch:
                rename_map[col_batch] = "batch_nbr"
            if col_expiry:
                rename_map[col_expiry] = "Expiry Date"

            df = df.rename(columns=rename_map)
            for c in display_columns:
                if c not in df.columns:
                    df[c] = ""

            df = df[display_columns]

            # تحويل Expiry Date وتحديد نطاقات: 1–3، 3–6، 6–9 شهور
            today = pd.Timestamp(datetime.now().date())
            three_months = today + pd.DateOffset(months=3)
            six_months = today + pd.DateOffset(months=6)
            nine_months = today + pd.DateOffset(months=9)

            expiry_ser = pd.to_datetime(df["Expiry Date"], errors="coerce")
            df["_expiry_dt"] = expiry_ser
            df["Expiry Date"] = expiry_ser.dt.strftime("%Y-%m-%d").fillna("")

            within_1_3 = (
                (df["_expiry_dt"].notna())
                & (df["_expiry_dt"] >= today)
                & (df["_expiry_dt"] <= three_months)
            )
            within_3_6 = (
                (df["_expiry_dt"].notna())
                & (df["_expiry_dt"] > three_months)
                & (df["_expiry_dt"] <= six_months)
            )
            within_6_9 = (
                (df["_expiry_dt"].notna())
                & (df["_expiry_dt"] > six_months)
                & (df["_expiry_dt"] <= nine_months)
            )
            df = df.drop(columns=["_expiry_dt"], errors="ignore")

            table_data = df[display_columns].to_dict(orient="records")

            # أعداد المنتجات لكل نطاق
            expiry_counts = {
                "within_1_3": int(within_1_3.sum()),
                "within_3_6": int(within_3_6.sum()),
                "within_6_9": int(within_6_9.sum()),
            }

            # خيارات الفلاتر: Facility, Company, Status, Expiry Date
            facility_codes = (
                sorted(
                    df["Facility"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Facility" in df.columns
                else []
            )
            companies = (
                sorted(
                    df["Company"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Company" in df.columns
                else []
            )
            statuses = (
                sorted(
                    df["Status"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Status" in df.columns
                else []
            )
            expiry_dates = (
                sorted(
                    df["Expiry Date"]
                    .astype(str)
                    .str.strip()
                    .replace("", pd.NA)
                    .dropna()
                    .unique()
                    .tolist()
                )
                if "Expiry Date" in df.columns
                else []
            )

            filter_options = {
                "facility_codes": facility_codes,
                "companies": companies,
                "statuses": statuses,
                "expiry_dates": expiry_dates,
            }

            return {
                "detail_html": "<div class='text-center p-4 text-muted fw-semibold'>Loading data...</div>",
                "chart_data": [],
                "count": 0,
                "tab_data": {"name": "Inventory"},
            }
        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error processing Expiry: {e}</p>",
                "chart_data": [],
                "count": 0,
            }

    def filter_total_lead_time_performance(
        self, request, selected_month=None, selected_months=None
    ):
        cache.clear()
        """
        🔹 عرض جدول Miss Breakdown (3PL و Roche كل واحد منفصل)
        🔹 عرض الشارت الخاص بـ 3PL On-Time Delivery
        🔹 عرض خطوات Outbound في الأسفل
        """
        try:
            excel_path = self.get_uploaded_file_path(request)
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger text-center'>⚠️ Excel file not found for display.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sub_tables = []
            chart_data = []
            selected_month_norm = None
            selected_months_norm = []
            actual_target = 0  # يُحدَّث من الشيت الرئيسي إن وُجد

            if selected_month:
                raw_month = str(selected_month).strip()
                parsed = pd.to_datetime(raw_month, errors="coerce")
                if pd.isna(parsed):
                    selected_month_norm = raw_month[:3].capitalize()
                else:
                    selected_month_norm = parsed.strftime("%b")

            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm:
                        selected_months_norm.append(norm)
                # إزالة التكرارات مع الحفاظ على الترتيب
                seen = set()
                selected_months_norm = [
                    m for m in selected_months_norm if not (m in seen or seen.add(m))
                ]

            # ----------------------------
            # 🟦 جدول 3PL SIDE
            # ----------------------------
            sheet_3pl = next(
                (
                    s
                    for s in xls.sheet_names
                    if "total lead time preformance" in s.lower()
                    and "-r" not in s.lower()
                ),
                None,
            )

            final_df_3pl = None

            if sheet_3pl:
                df = pd.read_excel(excel_path, sheet_name=sheet_3pl, engine="openpyxl")
                df.columns = df.columns.str.strip().str.lower()

                required_cols = [
                    "month",
                    "outbound delivery",
                    "kpi",
                    "reason group",
                    "miss reason",
                ]
                if all(col in df.columns for col in required_cols):
                    df["year"] = pd.to_datetime(df["month"], errors="coerce").dt.year
                    df = df[df["year"] == 2025]

                    if "month" in df.columns:
                        # نحاول تحويل القيم في عمود Month إلى تاريخ، ثم استخراج اسم الشهر المختصر
                        df["month"] = pd.to_datetime(
                            df["month"], errors="coerce"
                        ).dt.strftime("%b")
                    else:
                        # fallback لو مفيش عمود Month
                        df["month"] = pd.to_datetime(
                            df["ob distribution date"], errors="coerce"
                        ).dt.strftime("%b")

                    # ترتيب الشهور
                    month_order = [
                        "Jan",
                        "Feb",
                        "Mar",
                        "Apr",
                        "May",
                        "Jun",
                        "Jul",
                        "Aug",
                        "Sep",
                        "Oct",
                        "Nov",
                        "Dec",
                    ]

                    df["month"] = pd.Categorical(
                        df["month"], categories=month_order, ordered=True
                    )
                    missing_months = []
                    if selected_month_norm:
                        df = df[df["month"] == selected_month_norm]
                        if df.empty:
                            return {
                                "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for month {selected_month_norm} in Total Lead Time Performance.</p>",
                                "chart_data": [],
                                "count": 0,
                                "hit_pct": 0,
                            }
                        existing_months = [selected_month_norm]
                    elif selected_months_norm:
                        df = df[df["month"].isin(selected_months_norm)]
                        available_months = [
                            m
                            for m in selected_months_norm
                            if m in df["month"].dropna().unique()
                        ]
                        missing_months = [
                            m for m in selected_months_norm if m not in available_months
                        ]
                        if df.empty:
                            return {
                                "detail_html": "<p class='text-warning text-center p-4'>⚠️ No data available for the selected quarter months in Total Lead Time Performance.</p>",
                                "chart_data": [],
                                "count": 0,
                                "hit_pct": 0,
                            }
                        existing_months = selected_months_norm
                    else:
                        existing_months = [
                            m for m in month_order if m in df["month"].dropna().unique()
                        ]

                    df["reason group"] = (
                        df["reason group"].astype(str).str.strip().str.lower()
                    )
                    df["kpi"] = df["kpi"].astype(str).str.strip().str.lower()
                    df["miss reason"] = (
                        df["miss reason"].astype(str).str.strip().str.lower()
                    )

                    df_hit = df[df["kpi"].str.lower() == "hit"].copy()
                    hit_counts = (
                        df_hit.groupby("month")["outbound delivery"]
                        .nunique()
                        .reindex(existing_months, fill_value=0)
                    )

                    df_3pl_miss = df[
                        (df["kpi"].str.lower() == "miss")
                        & (df["reason group"] == "3pl")
                    ].copy()

                    miss_grouped = (
                        df_3pl_miss.groupby(["miss reason", "month"])[
                            "outbound delivery"
                        ]
                        .nunique()
                        .reset_index(name="count")
                        .pivot_table(
                            index="miss reason",
                            columns="month",
                            values="count",
                            fill_value=0,
                        )
                    )

                    for m in existing_months:
                        if m not in miss_grouped.columns:
                            miss_grouped[m] = 0
                    miss_grouped = miss_grouped[existing_months]

                    final_df_3pl = miss_grouped.copy()
                    final_df_3pl.loc["on time delivery"] = hit_counts
                    final_df_3pl = final_df_3pl.fillna(0).astype(int)
                    final_df_3pl["2025"] = final_df_3pl.sum(axis=1)

                    total_row = final_df_3pl.sum(numeric_only=True)
                    total_row.name = "total"
                    final_df_3pl = pd.concat([final_df_3pl, pd.DataFrame([total_row])])

                    final_df_3pl.reset_index(inplace=True)
                    final_df_3pl.rename(columns={"index": "KPI"}, inplace=True)
                    final_df_3pl["KPI"] = final_df_3pl["KPI"].str.title()

                    desired_order = [
                        "On Time Delivery",
                        "Late Arrive To The Customer",
                        "Customer Close On Arrive",
                        "Remote Area",
                    ]
                    final_df_3pl["order_key"] = final_df_3pl["KPI"].apply(
                        lambda x: (
                            desired_order.index(x)
                            if x in desired_order
                            else len(desired_order) + 1
                        )
                    )
                    final_df_3pl = final_df_3pl.sort_values(
                        by=["order_key", "KPI"]
                    ).drop(columns=["order_key"])
                    # final_df_3pl.insert(1, "Reason Group", "3PL")
                    #
                    # # ✅ حذف عمود Reason Group قبل الإرسال
                    # if "Reason Group" in final_df_3pl.columns:
                    #     final_df_3pl = final_df_3pl.drop(columns=["Reason Group"])

                    # ✅ حساب التارجت الفعلي لكل شهر (On Time ÷ Total × 100)
                    percent_hit = []
                    existing_months = [
                        m
                        for m in final_df_3pl.columns
                        if m not in ["KPI", "Reason Group", "2025", "Total"]
                    ]

                    on_time_row = final_df_3pl.loc[
                        final_df_3pl["KPI"].str.lower() == "on time delivery"
                    ].iloc[0]
                    total_row = final_df_3pl.loc[
                        final_df_3pl["KPI"].str.lower() == "total"
                    ].iloc[0]

                    for m in existing_months:
                        on_time_val = float(on_time_row.get(m, 0))
                        total_val = float(total_row.get(m, 0))

                        # ✅ لو الشهر فيه صفر فعلاً، خليه 0 في الشارت كمان
                        if total_val == 0 or on_time_val == 0:
                            percent = 0
                        else:
                            percent = int(round((on_time_val / total_val) * 100))

                        percent_hit.append(percent)

                    try:
                        total_year_val = total_row["2025"]
                        on_time_year_val = on_time_row["2025"]
                        actual_target = (
                            int(round((on_time_year_val / total_year_val) * 100))
                            if total_year_val > 0
                            else 0
                        )
                    except Exception:
                        actual_target = 100

                    # ✅ إنشاء قائمة بالشهور اللي فيها قيم غير صفرية (فقط للشارت)
                    nonzero_months = [
                        m for i, m in enumerate(existing_months) if percent_hit[i] > 0
                    ]
                    nonzero_percents = [
                        percent_hit[i]
                        for i, m in enumerate(existing_months)
                        if percent_hit[i] > 0
                    ]
                    if not nonzero_months:
                        nonzero_months = existing_months
                        nonzero_percents = [
                            percent_hit[i] for i in range(len(existing_months))
                        ]

                    chart_data.append(
                        {
                            "type": "column",
                            "name": "On-Time Delivery (%)",
                            "color": "#9fc0e4",
                            "showInLegend": True,
                            "related_table": "Miss Breakdown – 3PL Side",  # ✅ ربط الشارت بالجدول
                            "dataPoints": [
                                {"label": m, "y": nonzero_percents[i]}
                                for i, m in enumerate(nonzero_months)
                            ],
                        }
                    )
                    chart_data.append(
                        {
                            "type": "line",
                            "name": f"Target ({actual_target}%)",
                            "color": "red",
                            "showInLegend": True,
                            "related_table": "Miss Breakdown – 3PL Side",  # ✅ ربط الشارت بالجدول
                            "dataPoints": [
                                {"label": m, "y": actual_target} for m in nonzero_months
                            ],
                        }
                    )

                    sub_tables.append(
                        {
                            "title": "Miss Breakdown – 3PL Side",
                            "columns": list(final_df_3pl.columns),
                            "data": final_df_3pl.to_dict(orient="records"),
                        }
                    )
                    # لم نعد نضيف جدول Missing Months هنا، يتم التعامل معه لاحقًا عبر apply_month_filter_to_tab

            # ----------------------------
            # 🟥 جدول ROCHE SIDE
            # ----------------------------
            sheet_roche = next(
                (s for s in xls.sheet_names if "preformance -r" in s.lower()), None
            )
            if sheet_roche:
                df = pd.read_excel(
                    excel_path, sheet_name=sheet_roche, engine="openpyxl"
                )
                df.columns = df.columns.str.strip()
                if "Month" in df.columns:
                    month_order = [
                        "Jan",
                        "Feb",
                        "Mar",
                        "Apr",
                        "May",
                        "Jun",
                        "Jul",
                        "Aug",
                        "Sep",
                        "Oct",
                        "Nov",
                        "Dec",
                    ]
                    df["Month"] = pd.Categorical(
                        df["Month"], categories=month_order, ordered=True
                    )
                    df = df.sort_values("Month")

                    if selected_month_norm:
                        df_filtered = df[
                            df["Month"].astype(str).str.lower()
                            == selected_month_norm.lower()
                        ]
                        if df_filtered.empty:
                            sub_tables.append(
                                {
                                    "title": "Miss Breakdown – Roche Side",
                                    "columns": [],
                                    "data": [],
                                    "message": f"⚠️ لا توجد بيانات متاحة للشهر {selected_month_norm}.",
                                }
                            )
                        else:
                            df_melted = df_filtered.melt(
                                id_vars=["Month"], var_name="KPI", value_name="Count"
                            )
                            pivot_df = (
                                df_melted.groupby(["KPI", "Month"])["Count"]
                                .sum()
                                .unstack(fill_value=0)
                            )
                            pivot_df["2025"] = pivot_df.sum(axis=1)
                            total_row = pivot_df.sum(numeric_only=True)
                            total_row.name = "TOTAL"
                            pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])])
                            pivot_df.reset_index(inplace=True)
                            pivot_df.rename(columns={"index": "KPI"}, inplace=True)
                            keep_cols = [
                                col
                                for col in ["KPI", selected_month_norm]
                                if col in pivot_df.columns
                            ]
                            pivot_df = pivot_df[keep_cols]
                            sub_tables.append(
                                {
                                    "title": "Miss Breakdown – Roche Side",
                                    "columns": list(pivot_df.columns),
                                    "data": pivot_df.to_dict(orient="records"),
                                }
                            )
                    elif selected_months_norm:
                        df_filtered = df[
                            df["Month"]
                            .astype(str)
                            .str.lower()
                            .isin([m.lower() for m in selected_months_norm])
                        ]
                        if df_filtered.empty:
                            sub_tables.append(
                                {
                                    "title": "Miss Breakdown – Roche Side",
                                    "columns": [],
                                    "data": [],
                                    "message": "⚠️ No data available for the selected quarter months.",
                                }
                            )
                        else:
                            df_melted = df_filtered.melt(
                                id_vars=["Month"], var_name="KPI", value_name="Count"
                            )
                            pivot_df = (
                                df_melted.groupby(["KPI", "Month"])["Count"]
                                .sum()
                                .unstack(fill_value=0)
                            )
                            ordered_months = [
                                m for m in selected_months_norm if m in pivot_df.columns
                            ]
                            for m in selected_months_norm:
                                if m not in pivot_df.columns:
                                    pivot_df[m] = 0
                            pivot_df = pivot_df[selected_months_norm]
                            pivot_df["2025"] = pivot_df.sum(axis=1)
                            total_row = pivot_df.sum(numeric_only=True)
                            total_row.name = "TOTAL"
                            pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])])
                            pivot_df.reset_index(inplace=True)
                            pivot_df.rename(columns={"index": "KPI"}, inplace=True)
                            sub_tables.append(
                                {
                                    "title": "Miss Breakdown – Roche Side",
                                    "columns": list(pivot_df.columns),
                                    "data": pivot_df.to_dict(orient="records"),
                                }
                            )
                    else:
                        df_melted = df.melt(
                            id_vars=["Month"], var_name="KPI", value_name="Count"
                        )
                        pivot_df = (
                            df_melted.groupby(["KPI", "Month"])["Count"]
                            .sum()
                            .unstack(fill_value=0)
                            .reindex(columns=month_order, fill_value=0)
                        )
                        pivot_df["2025"] = pivot_df.sum(axis=1)
                        total_row = pivot_df.sum(numeric_only=True)
                        total_row.name = "TOTAL"
                        pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])])
                        pivot_df.reset_index(inplace=True)
                        pivot_df.rename(columns={"index": "KPI"}, inplace=True)
                        pivot_df = pivot_df.loc[:, (pivot_df != 0).any(axis=0)]

                        sub_tables.append(
                            {
                                "title": "Miss Breakdown – Roche Side",
                                "columns": list(pivot_df.columns),
                                "data": pivot_df.to_dict(orient="records"),
                            }
                        )

            # Outbound Shipments (Outbound1 + Outbound2, Hit/Miss) — نفس فكرة Inbound
            outbound_facility = (request.GET.get("outbound_facility") or "").strip()
            outbound_order_nbr = (request.GET.get("outbound_order_nbr") or "").strip()
            outbound_result = self.filter_outbound_shipments(
                request,
                selected_month if not selected_months_norm else None,
                selected_months_norm if selected_months_norm else None,
                selected_facility=outbound_facility or None,
                selected_order_nbr=outbound_order_nbr or None,
            )
            # ✅ جلب نسبة الـ Hit من Outbound (هي اللي هنستخدمها كـ KPI للتاب ده)
            outbound_stats = outbound_result.get("stats", {}) or {}
            outbound_hit_pct = outbound_stats.get("hit_pct", 0) or 0
            # ✅ إذا لم تكن موجودة في stats، نحاول جلبها مباشرة من outbound_result
            if not outbound_hit_pct:
                outbound_hit_pct = outbound_result.get("hit_pct", 0) or 0
            print(
                f"🔍 Total Lead Time Performance - Outbound hit_pct: {outbound_hit_pct}% (from stats: {outbound_stats.get('hit_pct', 'N/A')})"
            )

            if outbound_result.get("sub_tables"):
                outbound_tab = {
                    "name": "Outbound Shipments",
                    "stats": outbound_result.get("stats", {}),
                    "sub_tables": outbound_result["sub_tables"],
                    "chart_data": outbound_result.get("chart_data", []),
                }
                outbound_html = render_to_string(
                    "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                    {
                        "tab": outbound_tab,
                        "selected_month": selected_month
                        or (selected_months_norm[0] if selected_months_norm else None),
                    },
                )
            else:
                outbound_html = outbound_result.get("detail_html", "")

            # لا نرجع "لا توجد بيانات" إلا لو مفيش جداول رئيسية ومفيش محتوى Outbound
            has_outbound = bool(outbound_html and str(outbound_html).strip())
            if not sub_tables and not has_outbound:
                return {
                    "detail_html": "<p class='text-muted'>⚠️ No valid data was found in any sheets.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            # ✅ لا نحتاج لتعيين related_table هنا لأنه تم تعيينه بالفعل لكل dataset
            # if chart_data:
            #     for dataset in chart_data:
            #         dataset.setdefault("related_table", "Total Lead Time Performance")

            tab_data = {
                "name": "Outbound",
                "sub_tables": sub_tables,
                "outbound_html": outbound_html,
                "chart_data": chart_data,
            }
            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data,
                selected_month_norm if not selected_months_norm else None,
                selected_months_norm or None,
            )

            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {
                    "tab": tab_data,
                    "selected_month": month_norm_tab,
                    "selected_months": selected_months_norm,
                },
            )

            total_count = sum(len(st["data"]) for st in sub_tables)

            # ✅ نسبة الـ Hit الخاصة بالـ Outbound (هي اللي هنستخدمها كـ KPI للتاب ده)
            try:
                hit_pct_calculated = (
                    float(outbound_hit_pct) if outbound_hit_pct else 0.0
                )
                hit_pct_calculated = round(hit_pct_calculated, 2)  # تقريب لرقمين عشريين
            except (ValueError, TypeError):
                hit_pct_calculated = 0.0
            print(
                f"✅ Total Lead Time Performance - Using Outbound hit_pct: {hit_pct_calculated}%"
            )

            # ✅ إذا لم يكن هناك chart_data من 3PL، نستخدم chart_data من Outbound
            if not chart_data:
                outbound_chart_data = outbound_result.get("chart_data", []) or []
                if outbound_chart_data:
                    chart_data = outbound_chart_data
                    print(
                        f"✅ Total Lead Time Performance - Using Outbound chart_data: {len(chart_data)} datasets"
                    )

            print(
                f"✅ Total Lead Time Performance - Final chart_data: {len(chart_data)} datasets"
            )

            return {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "Total Lead Time Performance – On-Time Delivery",
                "count": total_count,
                "hit_pct": hit_pct_calculated,  # ✅ نسبة الـ Hit من Outbound
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing data: {e}</p>",
                "chart_data": [],
                "count": 0,
                "hit_pct": 0,  # ✅ إضافة hit_pct في حالة الخطأ
            }

    def filter_dock_to_stock_combined(
        self, request, selected_month=None, selected_months=None
    ):
        """
        🔹 يعرض تاب Dock to stock بالاعتماد على تحليل Inbound (KPI ≤24h).
        """
        cache.clear()
        print("🚀 معالجة Dock to stock — Inbound KPI")

        try:
            from django.template.loader import render_to_string

            inbound_result = self.filter_inbound(
                request, selected_month, selected_months
            )
            sub_tables = inbound_result.get("sub_tables", [])
            chart_data = inbound_result.get("chart_data", [])

            if not sub_tables:
                fallback_html = inbound_result.get("detail_html") or (
                    "<p class='text-warning'>⚠️ No inbound data available.</p>"
                )
                return {
                    "chart_data": chart_data,
                    "detail_html": fallback_html,
                    "count": 0,
                }

            tab_data = {
                "name": "Inbound",
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "canvas_id": "chart-inbound-kpi",
            }

            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in seen:
                        seen.add(norm)
                        selected_months_norm.append(norm)

            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data,
                None if selected_months_norm else selected_month,
                selected_months_norm or None,
            )

            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm_tab},
            )

            stats = inbound_result.get("stats", {})
            total_count = stats.get(
                "total", sum(len(st.get("data", [])) for st in sub_tables)
            )
            hit_pct = stats.get("hit_pct", 0)

            result = {
                "chart_data": chart_data,
                "detail_html": html,
                "count": total_count,
                "canvas_id": tab_data["canvas_id"],
                "hit_pct": hit_pct,
                "target_pct": 100,
                "tab_data": tab_data,
            }
            return _sanitize_for_json(result)
        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "chart_data": [],
                "detail_html": f"<p class='text-danger'>⚠️ Error: {e}</p>",
                "count": 0,
            }
        cache.clear()
        print("🚀 دخلنا الدالة filter_dock_to_stock_combined")

        """
        ✅ فصل Dock to stock إلى جدولين (3PL + Roche)
        ✅ ترتيب الشهور Jan → Dec
        ✅ حساب التارجت الصحيح (on time / total * 100)
        ✅ الشارت موحد (On Time % + Target)
        ✅ عرض الجداول منفصلة
        """
        try:
            import pandas as pd
            import numpy as np
            import os
            from django.template.loader import render_to_string
            from django.utils.text import slugify

            if request and hasattr(request, "session"):
                excel_path = (
                    request.session.get("uploaded_excel_path") or self.get_excel_path()
                )
            else:
                excel_path = self.get_excel_path()

            if not excel_path or not os.path.exists(excel_path):
                return {
                    "chart_data": [],
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "count": 0,
                }

            # ترتيب الشهور
            def order_months(months):
                month_map = {
                    "jan": 1,
                    "feb": 2,
                    "mar": 3,
                    "apr": 4,
                    "may": 5,
                    "jun": 6,
                    "jul": 7,
                    "aug": 8,
                    "sep": 9,
                    "oct": 10,
                    "nov": 11,
                    "dec": 12,
                }
                months_unique = list(dict.fromkeys(months))

                def month_key(m):
                    if m is None:
                        return 999
                    m_str = str(m).strip()
                    m_lower = m_str.lower()[:3]
                    if m_lower in month_map:
                        return month_map[m_lower]
                    if m_str.isdigit():
                        return 1000 + int(m_str)
                    return 2000 + months_unique.index(m)

                return sorted(months_unique, key=month_key)

            # =======================================
            # 🟢 معالجة Dock to Stock (3PL)
            # =======================================
            selected_months_norm = []
            if selected_months:
                if isinstance(selected_months, str):
                    selected_months = [selected_months]
                seen = set()
                for m in selected_months:
                    norm = self.normalize_month_label(m)
                    if norm and norm not in seen:
                        seen.add(norm)
                        selected_months_norm.append(norm)

            result_3pl = self.filter_dock_to_stock_3pl(
                request, selected_month, selected_months
            )
            df_3pl_table = pd.DataFrame()
            df_chart_combined = {}
            selected_month_norm = None
            if selected_month and not selected_months_norm:
                raw_month = str(selected_month).strip()
                parsed = pd.to_datetime(raw_month, errors="coerce")
                if pd.isna(parsed):
                    selected_month_norm = raw_month[:3].capitalize()
                else:
                    selected_month_norm = parsed.strftime("%b")

            if "chart_data" in result_3pl and result_3pl["chart_data"]:
                df_kpi_full = pd.DataFrame(result_3pl["chart_data"])

                # تحويل الأرقام إلى int
                for col in df_kpi_full.columns:
                    if col != "KPI":
                        df_kpi_full[col] = df_kpi_full[col].apply(
                            lambda x: int(round(float(x))) if pd.notna(x) else 0
                        )

                # حساب النسب الشهرية
                on_time_rows = df_kpi_full[
                    df_kpi_full["KPI"].str.lower().str.contains("on time", na=False)
                ]
                total_rows = df_kpi_full[
                    df_kpi_full["KPI"].str.lower().str.contains("total", na=False)
                ]

                target_correct, on_time_percentage = {}, {}
                month_cols = [
                    c
                    for c in df_kpi_full.columns
                    if c not in ["KPI", "2025", "Total", "TOTAL"]
                ]

                for col in month_cols:
                    try:
                        on_time_val = float(on_time_rows[col].sum())
                        total_val = float(total_rows[col].sum())
                        percentage = (
                            int(round((on_time_val / total_val) * 100))
                            if total_val
                            else 0
                        )
                        target_correct[col] = percentage
                        on_time_percentage[col] = percentage
                    except Exception as e:
                        print(f"⚠️ Error in {col}: {e}")
                        target_correct[col] = on_time_percentage[col] = 0

                df_chart_combined["3PL On Time %"] = on_time_percentage
                df_chart_combined["Target"] = target_correct

                # تجهيز الجدول النهائي
                df_kpi = df_kpi_full[
                    ~df_kpi_full["KPI"].str.lower().str.contains("target", na=False)
                ].copy()
                ordered_cols = ["KPI"] + [
                    c for c in order_months(df_kpi.columns.tolist()) if c != "KPI"
                ]
                df_3pl_table = df_kpi[ordered_cols]
                if selected_months_norm:
                    keep_cols = ["KPI"] + [
                        m for m in selected_months_norm if m in df_3pl_table.columns
                    ]
                    if "2025" in df_3pl_table.columns:
                        keep_cols.append("2025")
                    df_3pl_table = df_3pl_table[
                        [col for col in keep_cols if col in df_3pl_table.columns]
                    ]
                elif selected_month_norm:
                    keep_cols = ["KPI", selected_month_norm]
                    if "2025" in df_3pl_table.columns:
                        keep_cols.append("2025")
                    df_3pl_table = df_3pl_table[
                        [col for col in keep_cols if col in df_3pl_table.columns]
                    ]

                # ✅ إضافة صف "3PL Delay" بعد "On Time Receiving"
                on_time_receiving_idx = None
                for idx in df_3pl_table.index:
                    kpi_value = str(df_3pl_table.loc[idx, "KPI"]).strip()
                    if "on time receiving" in kpi_value.lower():
                        on_time_receiving_idx = idx
                        break

                if on_time_receiving_idx is not None:
                    # إنشاء صف جديد بقيم صفرية
                    delay_row = {"KPI": "3PL Delay"}
                    for col in df_3pl_table.columns:
                        if col != "KPI":
                            delay_row[col] = 0

                    # تحويل DataFrame إلى قائمة من القواميس
                    rows_list = df_3pl_table.to_dict(orient="records")

                    # العثور على موضع الصف في القائمة
                    insert_position = None
                    for i, row_dict in enumerate(rows_list):
                        kpi_value = str(row_dict.get("KPI", "")).strip()
                        if "on time receiving" in kpi_value.lower():
                            insert_position = i + 1
                            break

                    # إدراج الصف الجديد
                    if insert_position is not None:
                        rows_list.insert(insert_position, delay_row)
                        df_3pl_table = pd.DataFrame(rows_list)

            reasons_3pl = result_3pl.get("reason", [])

            # =======================================
            # 🔵 معالجة Dock to Stock (Roche)
            # =======================================
            reasons_roche = []
            try:

                # df_roche = pd.read_excel(excel_path, sheet_name="Dock to stock - Roche", engine="openpyxl")
                # قراءة كل الشيتات أولاً
                xls = pd.ExcelFile(excel_path, engine="openpyxl")

                # محاولة إيجاد الشيت الصحيح تلقائيًا (حتى لو الاسم فيه مسافات أو اختلاف حروف)
                sheet_name = None
                for name in xls.sheet_names:
                    if (
                        "dock" in name.lower()
                        and "stock" in name.lower()
                        and "roche" in name.lower()
                    ):
                        sheet_name = name
                        break

                if not sheet_name:
                    raise ValueError(
                        f"❌ لم يتم العثور على شيت Roche في الملف. الشيتات المتاحة: {xls.sheet_names}"
                    )

                print(f"✅ تم استخدام الشيت: {sheet_name}")

                # قراءة الشيت الصحيح
                df_roche = pd.read_excel(xls, sheet_name=sheet_name)
                df_roche.columns = df_roche.columns.astype(str).str.strip()

                print("🔍 Roche columns:", df_roche.columns.tolist())

                month_col = df_roche.columns[0]

                melted_df = df_roche.melt(
                    id_vars=[month_col], var_name="KPI", value_name="Value"
                )
                pivot_df = (
                    melted_df.pivot_table(
                        index="KPI", columns=month_col, values="Value", aggfunc="sum"
                    )
                    .reset_index()
                    .rename_axis(None, axis=1)
                )

                # تحويل القيم إلى int
                for col in pivot_df.columns:
                    if col != "KPI":
                        pivot_df[col] = pivot_df[col].apply(
                            lambda x: int(round(float(x))) if pd.notna(x) else 0
                        )

                ordered_cols = ["KPI"] + [
                    c for c in order_months(pivot_df.columns.tolist()) if c != "KPI"
                ]
                pivot_df = pivot_df[ordered_cols]

                # حذف الأعمدة "Total" بعد الشهور
                pivot_df = pivot_df.loc[
                    :, ~pivot_df.columns.str.lower().str.contains("total")
                ]

                # حساب عمود 2025 (إجمالي كل الشهور)
                # حساب عمود 2025 (إجمالي كل الشهور)
                month_cols = [
                    c
                    for c in pivot_df.columns
                    if c not in ["KPI", "Reason Group", "2025"]
                ]
                pivot_df["2025"] = pivot_df[month_cols].sum(axis=1).astype(int)

                # إضافة صف Total في نهاية الجدول
                total_row = {"KPI": "Total (Roche)"}
                for col in pivot_df.columns:
                    if col != "KPI":
                        total_row[col] = int(pivot_df[col].sum())
                pivot_df = pd.concat(
                    [pivot_df, pd.DataFrame([total_row])], ignore_index=True
                )

                # حذف عمود Reason Group نهائيًا قبل الإرجاع
                if "Reason Group" in pivot_df.columns:
                    pivot_df = pivot_df.drop(columns=["Reason Group"])

                df_roche_table = pivot_df
                if selected_months_norm:
                    roche_cols = ["KPI"] + [
                        m for m in selected_months_norm if m in df_roche_table.columns
                    ]
                    if "2025" in df_roche_table.columns:
                        roche_cols.append("2025")
                    df_roche_table = df_roche_table[
                        [col for col in roche_cols if col in df_roche_table.columns]
                    ]
                elif selected_month_norm:
                    roche_cols = ["KPI", selected_month_norm]
                    if "2025" in df_roche_table.columns:
                        roche_cols.append("2025")
                    df_roche_table = df_roche_table[
                        [col for col in roche_cols if col in df_roche_table.columns]
                    ]
                # reasons_roche = self.filter_dock_to_stock_roche_reasons(request)
                reasons_roche = []

            except Exception as e:
                print(f"⚠️ Roche read error: {e}")
                df_roche_table = pd.DataFrame()

            # =======================================
            # 🟣 تجهيز الشارت
            # =======================================
            all_months = order_months(
                sorted(
                    set().union(*[list(v.keys()) for v in df_chart_combined.values()])
                )
            )
            if selected_months_norm:
                all_months = [m for m in selected_months_norm if m in all_months]
            on_time_values = df_chart_combined.get("3PL On Time %", {})
            target_values = df_chart_combined.get("Target", {})

            hit_pct = (
                min(round(float(np.mean(list(on_time_values.values()))), 2), 100)
                if on_time_values
                else 0
            )
            target_pct = (
                min(round(float(np.mean(list(target_values.values()))), 2), 100)
                if target_values
                else 100
            )

            chart_data = []
            if selected_month_norm or any(v != 0 for v in on_time_values.values()):
                chart_data.append(
                    {
                        "type": "column",
                        "name": "On time receiving (%)",
                        "color": "#d0e7ff",
                        "showInLegend": False,  # ✅ إخفاء الـ legend لتجنب التكرار
                        "dataPoints": [
                            {"label": m, "y": min(float(on_time_values.get(m, 0)), 100)}
                            for m in all_months
                        ],
                    }
                )

            # ✅ إزالة dataset الـ target لأننا نستخدم خط مخصص فقط
            # if selected_month_norm or any(v != 0 for v in target_values.values()):
            #     chart_data.append(...)

            inbound_result = self.filter_inbound(
                request, selected_month, selected_months
            )
            inbound_html = inbound_result.get("detail_html", "")
            inbound_sub_table = inbound_result.get("sub_table")
            combined_reasons = list(reasons_3pl) + list(reasons_roche)

            # =======================================
            # 🧱 بناء العرض النهائي
            # =======================================
            if chart_data:
                for dataset in chart_data:
                    dataset.setdefault("related_table", "Inbound")

            # ✅ إضافة chart_data لكل sub_table بشكل منفصل
            chart_data_3pl = []
            chart_data_roche = []
            if chart_data:
                for dataset in chart_data:
                    dataset_3pl = dataset.copy()
                    dataset_3pl["related_table"] = "Inbound — 3PL"
                    chart_data_3pl.append(dataset_3pl)

                    dataset_roche = dataset.copy()
                    dataset_roche["related_table"] = "Inbound — Roche"
                    chart_data_roche.append(dataset_roche)

            tab_data = {
                "name": "Inbound",
                "sub_tables": [
                    {
                        "id": "sub-table-inbound-3pl",
                        "title": "Inbound — 3PL",
                        "columns": df_3pl_table.columns.tolist(),
                        "data": df_3pl_table.to_dict(orient="records"),
                        "chart_data": chart_data_3pl,
                    },
                    {
                        "id": "sub-table-inbound-roche",
                        "title": "Inbound — Roche",
                        "columns": df_roche_table.columns.tolist(),
                        "data": df_roche_table.to_dict(orient="records"),
                        "chart_data": chart_data_roche,
                    },
                ],
                "combined_reasons": combined_reasons,
                "canvas_id": f"chart-{slugify('inbound')}",
                "inbound_html": inbound_html,
                "chart_data": chart_data,  # ✅ الاحتفاظ بـ chart_data العام أيضاً
            }
            if inbound_sub_table:
                tab_data["sub_tables"].append(inbound_sub_table)
            month_norm_tab = self.apply_month_filter_to_tab(
                tab_data,
                (
                    (selected_month_norm or selected_month)
                    if not selected_months_norm
                    else None
                ),
                selected_months_norm or None,
            )

            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {"tab": tab_data, "selected_month": month_norm_tab},
            )

            total_count = len(df_3pl_table) + len(df_roche_table)

            print(f"📊 [RESULT] Inbound — Hit={hit_pct}%, Target={target_pct}")

            return {
                "chart_data": chart_data,
                "detail_html": html,
                "count": total_count,
                "canvas_id": tab_data["canvas_id"],
                "hit_pct": hit_pct,
                "target_pct": target_pct,
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "chart_data": [],
                "detail_html": f"<p class='text-danger'>⚠️ Error: {e}</p>",
                "count": 0,
            }

    def overview_tab(
        self,
        request=None,
        selected_month=None,
        selected_months=None,
        from_all_in_one=False,
    ):
        from concurrent.futures import ThreadPoolExecutor, as_completed

        cache.clear()
        tab_cards = []

        target_manual = {
            "inbound": 99,
            "outbound": 98,
            "total lead time performance": 98,
            "pods update": 98,
            "return & refusal": 100,
            "rejections": 100,
            "inventory": 99,
        }

        def process_tab(tab_name):
            detail_html, count, hit_pct_val = "", 0, 0
            chart_data = []
            chart_type = "bar"
            try:
                res = {}
                tab_lower = tab_name.lower()
                month_for_filters = selected_month if not selected_months else None

                if tab_lower == "rejections":
                    res = self.filter_rejections_combined(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif tab_lower == "return & refusal":
                    # Return tab: استخدم نفس منطق Return الجديد (Shipment_nbr تبدأ بـ 250 من ARAMCO Inbound Report)
                    res = self.filter_inbound(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                        tab_name="Return & Refusal",
                    )
                elif tab_lower == "inbound":
                    res = self.filter_dock_to_stock_combined(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif "pods update" in tab_lower:
                    res = self.filter_pods_update(request, month_for_filters)
                elif tab_lower == "outbound" or "total lead time performance" in tab_lower:
                    res = self.filter_total_lead_time_performance(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif tab_lower == "inventory":
                    res = self.filter_inventory(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )

                # النسبة الفعلية من التاب (من res أو من stats)
                hit_pct = res.get("hit_pct")
                if hit_pct is None and isinstance(res.get("stats"), dict):
                    hit_pct = res["stats"].get("hit_pct")
                if hit_pct is None:
                    hit_pct = 0
                if isinstance(hit_pct, dict):
                    if selected_month and selected_month.capitalize() in hit_pct:
                        hit_pct_val = hit_pct[selected_month.capitalize()]
                    else:
                        hit_pct_val = int(round(sum(hit_pct.values()) / len(hit_pct))) if hit_pct else 0
                else:
                    try:
                        hit_pct_val = int(round(float(hit_pct)))
                    except (TypeError, ValueError):
                        hit_pct_val = 0

                hit_pct_val = max(0, min(hit_pct_val, 100))

                target_pct = target_manual.get(tab_lower, 100)
                color_class = "bg-success" if hit_pct_val >= target_pct else "bg-danger"

                progress_html = f"""
                    <div class='mb-3'>
                        <div class='d-flex justify-content-between align-items-center mb-1'>
                            <strong class='text-capitalize'>{tab_name}</strong>
                            <small>{hit_pct_val}% / Target: {target_pct}%</small>
                        </div>
                        <div class='progress' style='height: 20px;'>
                            <div class='progress-bar {color_class}' role='progressbar'
                                 style='width: {hit_pct_val}%;' aria-valuenow='{hit_pct_val}'
                                 aria-valuemin='0' aria-valuemax='100'>
                                 {hit_pct_val}%
                            </div>
                        </div>
                    </div>
                """

                detail_html = progress_html + (res.get("detail_html", "") or "")
                count = res.get("count", 0)
                chart_data = res.get("chart_data", []) or []
                if chart_data and isinstance(chart_data, list) and len(chart_data) > 0:
                    chart_type = (chart_data[0].get("type") or "bar").lower()

            except Exception:
                detail_html = "<p class='text-muted'>No data available.</p>"
                hit_pct_val = 0
                target_pct = target_manual.get(tab_name.lower(), 100)

            return {
                "name": tab_name,
                "hit_pct": hit_pct_val,
                "target_pct": target_pct,
                "detail_html": detail_html,
                "count": count,
                "chart_data": chart_data,
                "chart_type": chart_type,
            }

        tabs_order = [
            "Inbound",
            "Outbound",
            "Return & Refusal",
            "PODs update",
            "Inventory",
        ]

        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(process_tab, t): t for t in tabs_order}
            for future in as_completed(futures):
                tab_cards.append(future.result())

        tab_cards.sort(key=lambda x: tabs_order.index(x["name"]))

        if not from_all_in_one:
            tab_cards = [
                t
                for t in tab_cards
                if t.get("name", "").strip().lower()
                not in ["rejections", "return & refusal"]
            ]

        all_progress_html = "<div class='card p-4 shadow-sm rounded-4 mb-4'>"
        all_progress_html += (
            "<h5 class='fw-bold text-primary mb-3'>📈 نسب الأداء لكل التابات</h5>"
        )
        for tab in tab_cards:
            color_class = (
                "bg-success" if tab["hit_pct"] >= tab["target_pct"] else "bg-danger"
            )
            all_progress_html += f"""
                <div class='mb-3'>
                    <div class='d-flex justify-content-between align-items-center mb-1'>
                        <strong>{tab['name']}</strong>
                        <small>{tab['hit_pct']}% / Target: {tab['target_pct']}%</small>
                    </div>
                    <div class='progress' style='height: 20px;'>
                        <div class='progress-bar {color_class}' role='progressbar'
                             style='width: {tab['hit_pct']}%;' aria-valuenow='{tab['hit_pct']}'
                             aria-valuemin='0' aria-valuemax='100'>
                             {tab['hit_pct']}%
                        </div>
                    </div>
                </div>
            """
        all_progress_html += "</div>"

        return {"tab_cards": tab_cards, "detail_html": all_progress_html}

    def _get_dashboard_include_context(self, request):
        """
        يُرجع سياق الداشبورد (نفس منطق dashboard_tab) لاستخدامه عند include
        container-fluid-dashboard في التمبلت حتى تورث base واللينكات والشارتس.
        """
        context = get_dashboard_tab_context(request)
        context["title"] = self.DASHBOARD_TAB_NAME
        # نفس ملف التابات بالضبط (الملف المرفوع أو latest/all_sheet من المجلد) لقراءة كروت الداشبورد
        excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()

        # كل الداتا من الشيت فقط — لا قيم يدوية. لو مفيش ملف أو الشيت فاضي نستخدم قيم فارغة/صفر.
        # محاولة قراءة من كاش الداتابيز أولاً لتسريع فتح الداشبورد
        cache_hit = False
        if excel_path:
            excel_path_norm = os.path.normpath(os.path.abspath(excel_path))
            try:
                cached = DashboardDataCache.objects.filter(source_file_path=excel_path_norm).first()
                if not cached:
                    for c in DashboardDataCache.objects.all():
                        cp = (c.source_file_path or "").strip()
                        if cp == excel_path_norm:
                            cached = c
                            break
                        try:
                            if cp and os.path.exists(cp) and os.path.realpath(cp) == os.path.realpath(excel_path_norm):
                                cached = c
                                break
                        except OSError:
                            pass
                if cached and cached.data:
                    cache_data = cached.data
                    if _dashboard_cache_valid(cache_data, excel_path):
                        if cache_data.get("inbound_kpi") is not None:
                            context["inbound_kpi"] = cache_data["inbound_kpi"]
                        if cache_data.get("pending_shipments") is not None:
                            context["pending_shipments"] = cache_data["pending_shipments"]
                        for key in ("outbound_chart_data", "outbound_kpi", "outbound_kpi_keys_from_sheet",
                                    "pod_compliance_chart_data", "pod_status_breakdown",
                                    "returns_kpi", "returns_chart_data", "inventory_kpi",
                                    "inventory_warehouse_table", "inventory_capacity_data", "returns_region_table"):
                            if cache_data.get(key) is not None:
                                context[key] = cache_data[key]
                        cache_hit = True
            except Exception:
                pass
            if not cache_hit:
                try:
                    json_cache = _load_dashboard_cache_json()
                    cache_data = json_cache.get(excel_path_norm)
                    if not cache_data and json_cache:
                        for k, v in json_cache.items():
                            try:
                                if k == excel_path_norm or (os.path.exists(k) and os.path.realpath(k) == os.path.realpath(excel_path_norm)):
                                    cache_data = v
                                    break
                            except OSError:
                                pass
                    if cache_data and _dashboard_cache_valid(cache_data, excel_path):
                        if cache_data.get("inbound_kpi") is not None:
                            context["inbound_kpi"] = cache_data["inbound_kpi"]
                        if cache_data.get("pending_shipments") is not None:
                            context["pending_shipments"] = cache_data["pending_shipments"]
                        for key in ("outbound_chart_data", "outbound_kpi", "outbound_kpi_keys_from_sheet",
                                    "pod_compliance_chart_data", "pod_status_breakdown",
                                    "returns_kpi", "returns_chart_data", "inventory_kpi",
                                    "inventory_warehouse_table", "inventory_capacity_data", "returns_region_table"):
                            if cache_data.get(key) is not None:
                                context[key] = cache_data[key]
                        cache_hit = True
                except Exception:
                    pass

        if excel_path and not cache_hit:
            built = _build_dashboard_cache_data(excel_path)
            if built:
                built["_file_mtime"] = _get_file_mtime(excel_path)
                if built.get("inbound_kpi") is not None:
                    context["inbound_kpi"] = built["inbound_kpi"]
                if built.get("pending_shipments") is not None:
                    context["pending_shipments"] = built["pending_shipments"]
                for key in ("outbound_chart_data", "outbound_kpi", "outbound_kpi_keys_from_sheet",
                            "pod_compliance_chart_data", "pod_status_breakdown",
                            "returns_kpi", "returns_chart_data", "inventory_kpi",
                            "inventory_warehouse_table", "inventory_capacity_data", "returns_region_table"):
                    if built.get(key) is not None:
                        context[key] = built[key]
                if not context.get("inventory_kpi"):
                    context["inventory_kpi"] = {
                        "no_of_location": 0, "total_qty": 0,
                        "no_of_location_riyadh": 0, "no_of_location_dammam": 0, "no_of_location_jeddah": 0,
                        "total_qty_riyadh": 0, "total_qty_dammam": 0, "total_qty_jeddah": 0,
                        "hit_pct": 0, "total_skus": 0, "total_lpns": 0, "utilization_pct": "",
                    }
                excel_path_norm = os.path.normpath(os.path.abspath(excel_path))
                try:
                    DashboardDataCache.objects.update_or_create(
                        source_file_path=excel_path_norm,
                        defaults={"data": _dedupe_cache_data(built)},
                    )
                except Exception:
                    pass
                try:
                    _save_dashboard_cache_to_json(excel_path_norm, built)
                except Exception:
                    pass
            else:
                inbound_data = _read_inbound_data_from_excel(excel_path)
                if inbound_data:
                    context["inbound_kpi"] = inbound_data["inbound_kpi"]
                    context["pending_shipments"] = inbound_data["pending_shipments"]

                charts_from_excel = _read_dashboard_charts_from_excel(excel_path)
                for key, value in charts_from_excel.items():
                    if value is not None:
                        context[key] = value

                outbound_data = _read_outbound_data_from_excel(excel_path)
                if outbound_data and "outbound_kpi" in outbound_data:
                    context["outbound_kpi"] = outbound_data["outbound_kpi"]
                    context["outbound_kpi_keys_from_sheet"] = outbound_data.get("outbound_kpi_keys_from_sheet", [])

                try:
                    pods_result = self.filter_pods_update(request, selected_month=None, selected_months=None)
                    if pods_result and "error" not in pods_result and pods_result.get("sub_tables"):
                        st = pods_result["sub_tables"][0]
                        pod_chart = st.get("chart_data") or []
                        if pod_chart and isinstance(pod_chart, list):
                            categories_pod = []
                            series_pod = []
                            for s in pod_chart:
                                pts = s.get("dataPoints") or []
                                if pts and not categories_pod:
                                    categories_pod = [p.get("label", "") for p in pts]
                                data_vals = [float(p.get("y", 0)) for p in pts]
                                series_pod.append({
                                    "name": s.get("name") or "Hit %",
                                    "data": data_vals,
                                })
                            context["pod_compliance_chart_data"] = {
                                "categories": categories_pod,
                                "series": series_pod,
                            }
                except Exception:
                    pass
                if "pod_compliance_chart_data" not in context:
                    pods_data = _read_pods_data_from_excel(excel_path)
                    if pods_data:
                        context["pod_compliance_chart_data"] = {
                            "categories": pods_data.get("categories", []),
                            "series": pods_data.get("series", []),
                        }
                        if "pod_status_breakdown" in pods_data:
                            context["pod_status_breakdown"] = pods_data["pod_status_breakdown"]

                returns_data = _read_returns_data_from_excel(excel_path)
                if returns_data:
                    context["returns_kpi"] = returns_data.get("returns_kpi", {})
                    if "returns_chart_data" in returns_data:
                        context["returns_chart_data"] = returns_data["returns_chart_data"]

                inventory_data = _read_inventory_data_from_excel(excel_path)
                _empty_inv_kpi = {
                    "no_of_location": 0, "total_qty": 0,
                    "no_of_location_riyadh": 0, "no_of_location_dammam": 0, "no_of_location_jeddah": 0,
                    "total_qty_riyadh": 0, "total_qty_dammam": 0, "total_qty_jeddah": 0,
                    "hit_pct": 0, "total_skus": 0, "total_lpns": 0, "utilization_pct": "",
                }
                if inventory_data and inventory_data.get("inventory_kpi"):
                    inv = inventory_data["inventory_kpi"]
                    context["inventory_kpi"] = {
                        "no_of_location": inv.get("no_of_location", 0),
                        "no_of_location_riyadh": inv.get("no_of_location_riyadh", 0),
                        "no_of_location_dammam": inv.get("no_of_location_dammam", 0),
                        "no_of_location_jeddah": inv.get("no_of_location_jeddah", 0),
                        "total_qty": inv.get("total_qty", 0),
                        "total_qty_riyadh": inv.get("total_qty_riyadh", 0),
                        "total_qty_dammam": inv.get("total_qty_dammam", 0),
                        "total_qty_jeddah": inv.get("total_qty_jeddah", 0),
                        "hit_pct": inv.get("hit_pct", 0),
                        "total_skus": inv.get("total_skus", 0),
                        "total_lpns": inv.get("total_lpns", 0),
                        "utilization_pct": inv.get("utilization_pct", ""),
                    }
                else:
                    context["inventory_kpi"] = _empty_inv_kpi

                dashboard_wh = _read_dashboard_warehouse_from_excel(excel_path)
                if dashboard_wh:
                    context["inventory_warehouse_table"] = dashboard_wh.get("inventory_warehouse_table", [])
                    context["inventory_capacity_data"] = dashboard_wh.get("inventory_capacity_data", {})
                else:
                    capacity_data = _read_inventory_snapshot_capacity_from_excel(excel_path)
                    if capacity_data:
                        context["inventory_capacity_data"] = capacity_data.get("inventory_capacity_data", {})
                    warehouse_table = _read_inventory_warehouse_table_from_excel(excel_path)
                    if warehouse_table:
                        context["inventory_warehouse_table"] = warehouse_table.get("inventory_warehouse_table", [])

                returns_region = _read_returns_region_table_from_excel(excel_path)
                if returns_region:
                    context["returns_region_table"] = returns_region.get("returns_region_table", [])

        # Order Process Time (يُحسب من filter_outbound_shipments — يعتمد على request/كاش الشيتات)
        if excel_path:
            try:
                outbound_result = self.filter_outbound_shipments(
                    request,
                    selected_facility=(request.GET.get("outbound_facility") or "").strip() or None,
                    selected_order_nbr=(request.GET.get("outbound_order_nbr") or "").strip() or None,
                )
                ob_stats = outbound_result.get("stats") or {}
                context["order_process_kpi"] = {
                    "hit_pct": ob_stats.get("hit_pct", 0),
                    "miss_pct": ob_stats.get("miss_pct", 0),
                    "hit": ob_stats.get("hit", 0),
                    "miss": ob_stats.get("miss", 0),
                    "total": ob_stats.get("total", 0),
                }
                ob_chart_data = outbound_result.get("chart_data") or []
                if ob_chart_data and isinstance(ob_chart_data, list):
                    categories_ob = []
                    series_ob = []
                    for s in ob_chart_data:
                        pts = s.get("dataPoints") or []
                        if pts and not categories_ob:
                            categories_ob = [p.get("label", "") for p in pts]
                        data_vals = [float(p.get("y", 0)) for p in pts]
                        series_ob.append({
                            "name": s.get("name") or "",
                            "data": data_vals,
                        })
                    context["order_process_chart_data"] = {
                        "categories": categories_ob,
                        "series": series_ob,
                    }
            except Exception:
                pass

        # قيم فارغة/صفر فقط عند غياب الملف أو فشل القراءة (حتى لا يكسر القالب)
        _empty_inbound_kpi = {
            "number_of_vehicles": 0,
            "number_of_shipments": 0,
            "number_of_pallets": 0,
            "total_quantity": 0,
            "total_quantity_display": "0",
        }
        _empty_outbound_kpi = {
            "released_orders": 0,
            "picked_orders": 0,
            "number_of_pallets": 0,
        }
        _empty_pod_chart = {"categories": [], "series": []}
        _empty_pod_breakdown = [
            {"label": "On Time", "pct": 0, "color": "#7FB7A6"},
            {"label": "Pending", "pct": 0, "color": "#A8C8EB"},
            {"label": "Late", "pct": 0, "color": "#E8A8A2"},
        ]

        # تتبع أي أقسام تعرض قيماً افتراضية (صفر) لعرض تنبيه "ارفع الملف مرة أخرى"
        missing_by_section = {}
        if "inbound_kpi" not in context:
            missing_by_section["Dashboard – Inbound"] = ["Number of Shipments", "Number of Pallets (LPNs)", "Total Quantity", "Pending Shipments"]
        context.setdefault("inbound_kpi", _empty_inbound_kpi)
        context.setdefault("pending_shipments", [])

        _outbound_card_names = {"released_orders": "Released Orders", "picked_orders": "Picked Orders", "number_of_pallets": "Number of Pallets (LPNs)"}
        if "outbound_kpi" not in context:
            missing_by_section.setdefault("Dashboard – Outbound", []).extend(["Released Orders", "Picked Orders", "Number of Pallets (LPNs)"])
        else:
            keys_from_sheet = context.get("outbound_kpi_keys_from_sheet") or []
            for key, card_name in _outbound_card_names.items():
                if key not in keys_from_sheet:
                    missing_by_section.setdefault("Dashboard – Outbound", []).append(card_name)
        context.setdefault("outbound_kpi", _empty_outbound_kpi)
        # لا نضيف PODs Compliance (chart) لقائمة الناقص — الشارت يملأ من filter_pods_update وقد يظهر لاحقاً
        context.setdefault("outbound_chart_data", _empty_pod_chart)
        context.setdefault("pod_compliance_chart_data", _empty_pod_chart)
        context.setdefault("pod_status_breakdown", _empty_pod_breakdown)

        if "returns_kpi" not in context:
            missing_by_section["Dashboard – Returns"] = ["Total SKUs", "Total LPNs", "Returns chart"]
        context.setdefault("returns_kpi", {"total_skus": 0, "total_lpns": 0})
        # Total LPNs في قسم Returns = نفس رقم Inbound (Number of Pallets) ليتطابق مع الجزء الأعلى
        context["returns_kpi"] = dict(context["returns_kpi"])
        context["returns_kpi"]["total_lpns"] = context.get("inbound_kpi", {}).get("number_of_pallets", 0)
        context.setdefault("returns_chart_data", _empty_pod_chart)
        context.setdefault("order_process_chart_data", _empty_pod_chart)
        context.setdefault("order_process_kpi", {"hit_pct": 0, "miss_pct": 0, "hit": 0, "miss": 0, "total": 0})
        context.setdefault("returns_region_table", [])

        if "inventory_kpi" not in context:
            missing_by_section["Dashboard – Inventory"] = ["No of Location", "Total Qty", "Capacity chart", "Warehouse table"]
        context.setdefault("inventory_kpi", {
            "no_of_location": 0, "total_qty": 0,
            "no_of_location_riyadh": 0, "no_of_location_dammam": 0, "no_of_location_jeddah": 0,
            "total_qty_riyadh": 0, "total_qty_dammam": 0, "total_qty_jeddah": 0,
            "hit_pct": 0, "total_skus": 0, "total_lpns": 0, "utilization_pct": "",
        })
        context.setdefault("inventory_capacity_data", {"used": 0, "available": 0})
        context.setdefault("inventory_warehouse_table", [])

        context["dashboard_missing_data"] = [{"section": k, "cards": v} for k, v in missing_by_section.items()]
        return context

    def dashboard_tab(self, request):
        """
        🔹 تاب Dashboard: يعرض تصميم الداشبورد (container-fluid-dashboard).
        التمبلت منفصل عن excel-sheet-table ويُحمّل داخل منطقة المحتوى عند اختيار تاب Dashboard.
        نفس فكرة rejection: نرجع detail_html + chart_data + chart_title عشان الشارتات تبقى دينامك.
        """
        try:
            context = self._get_dashboard_include_context(request)
            html = render_to_string(
                "container-fluid-dashboard.html",
                context,
                request=request,
            )
            # نفس شكل الـ rejection: chart_data و chart_title للشارتات الدينامك
            outbound_chart = context.get("outbound_chart_data")
            chart_data = []
            if outbound_chart and isinstance(outbound_chart, dict):
                categories = outbound_chart.get("categories", [])
                series = outbound_chart.get("series", [])
                if categories and series is not None:
                    chart_data.append({
                        "type": "line",
                        "name": "POD Compliance",
                        "dataPoints": [{"label": c, "y": float(s)} for c, s in zip(categories, series)],
                    })
            return {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "Dashboard – POD Compliance",
                "dashboard_charts": {
                    "outbound": context.get("outbound_chart_data"),
                    "returns": context.get("returns_chart_data"),
                    "inventory": context.get("inventory_capacity_data"),
                },
            }
        except Exception as e:
            import traceback

            traceback.print_exc()
            return {"error": f"An error occurred while loading Dashboard: {e}"}

    def meeting_points_tab(self, request):
        """
        🔹 عرض تاب Meeting Points & Action مع إمكانية الفلترة حسب الحالة (منتهية / غير منتهية)
        """
        try:
            # ✅ جلب الحالة من الـ GET parameter
            status_filter = request.GET.get(
                "status"
            )  # القيم الممكنة: done / pending / all

            # ✅ استرجاع كل النقاط بالترتيب
            meeting_points = MeetingPoint.objects.all().order_by(
                "is_done", "-created_at"
            )

            # ✅ تطبيق الفلترة بناءً على الحالة
            if status_filter == "done":
                meeting_points = meeting_points.filter(is_done=True)
            elif status_filter == "pending":
                meeting_points = meeting_points.filter(is_done=False)
            # 'all' يعرض كل النقاط (done + pending)
            # لا حاجة لفلترة إضافية لأنه استرجعنا كل النقاط في البداية

            # ✅ إحصائيات
            done_count = meeting_points.filter(is_done=True).count()
            total_count = meeting_points.count()

            # ✅ تجهيز البيانات للتمبلت مع assigned_to
            meeting_data = [
                {
                    "id": p.id,
                    "description": p.description,
                    "assigned_to": getattr(
                        p, "assigned_to", ""
                    ),  # ✅ الاسم ممكن يكون فاضي
                    "status": "Done" if p.is_done else "Pending",
                    "created_at": p.created_at,
                    "target_date": p.target_date,
                }
                for p in meeting_points
            ]

            context = {
                "meeting_points": meeting_points,
                "meeting_data": meeting_data,  # لو حابة تستخدمي البيانات مباشرة في JS
                "done_count": done_count,
                "total_count": total_count,
                "status_filter": status_filter,
            }

            # ✅ بناء HTML من التمبلت
            html = render_to_string("meeting_points.html", context, request=request)

            # ✅ إرجاع النتيجة
            return JsonResponse(
                {
                    "detail_html": html,
                    "count": meeting_points.count(),
                    "done_count": done_count,
                    "total_count": total_count,
                },
                safe=False,
            )

        except Exception as e:
            import traceback

            traceback.print_exc()
            return JsonResponse(
                {"error": f"An error occurred while loading data: {e}"}, status=500
            )


class MeetingPointListCreateView(View):
    template_name = "meeting_points.html"

    def get(self, request, *args, **kwargs):
        status_filter = request.GET.get("status")  # "done" أو "pending" أو None

        today = date.today()
        current_month, current_year = today.month, today.year

        # حساب الشهر السابق
        if current_month == 1:
            prev_month = 12
            prev_year = current_year - 1
        else:
            prev_month = current_month - 1
            prev_year = current_year

        # ✅ جلب كل النقاط (الشهر الحالي كله + pending من الشهر السابق)
        meeting_points = MeetingPoint.objects.filter(
            Q(created_at__year=current_year, created_at__month=current_month)
            | Q(created_at__year=prev_year, created_at__month=prev_month, is_done=False)
        ).order_by("is_done", "-created_at")

        # ✅ تطبيق الفلتر لو المستخدم اختار حاجة
        if status_filter == "done":
            meeting_points = meeting_points.filter(is_done=True)
        elif status_filter == "pending":
            meeting_points = meeting_points.filter(is_done=False)

        done_count = meeting_points.filter(is_done=True).count()
        total_count = meeting_points.count()

        return render(
            request,
            self.template_name,
            {
                "meeting_points": meeting_points,
                "done_count": done_count,
                "total_count": total_count,
                "status_filter": status_filter,
            },
        )

    def post(self, request, *args, **kwargs):
        description = request.POST.get("description", "").strip()
        target_date = request.POST.get("target_date", "").strip() or None
        assigned_to = request.POST.get("assigned_to", "").strip() or None

        if description:
            point = MeetingPoint.objects.create(
                description=description,
                target_date=target_date,
                assigned_to=assigned_to if assigned_to else None,
            )

            return JsonResponse(
                {
                    "id": point.id,
                    "description": point.description,
                    "assigned_to": point.assigned_to,
                    "created_at": str(point.created_at),
                    "target_date": str(point.target_date),
                    "is_done": point.is_done,
                }
            )

        return JsonResponse({"error": "Empty description"}, status=400)


class ToggleMeetingPointView(View):
    def post(self, request, pk, *args, **kwargs):
        point = get_object_or_404(MeetingPoint, pk=pk)
        point.is_done = not point.is_done
        point.save()
        return JsonResponse({"is_done": point.is_done})


class DoneMeetingPointView(View):
    def post(self, request, pk, *args, **kwargs):
        point = get_object_or_404(MeetingPoint, pk=pk)
        point.is_done = not point.is_done
        point.save()
        return JsonResponse({"is_done": point.is_done})
