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
from datetime import date
from django.db.models import Q
from django.template.loader import render_to_string
from calendar import month_abbr, month_name

from django.views.decorators.csrf import csrf_exempt
from django.utils.decorators import method_decorator
from django.utils.text import slugify

from .models import MeetingPoint


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
        priority_files = ["latest.xlsm", "latest.xlsx"] + self.DEFAULT_EXCEL_FILENAMES
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

        # --------------------------
        # Resolve Excel path
        # --------------------------
        excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
        data_is_uploaded = os.path.exists(excel_path)

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
                "overview": lambda: self.overview_tab(
                    request=request,
                    selected_month=effective_month,
                    selected_months=quarter_months or None,
                ),
                "dashboard": lambda: self.dashboard_tab(request),
                "all": lambda: self.filter_all_tabs(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "order general information": lambda: self.filter_order_general_information(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "return & refusal": lambda: self.filter_rejections_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "rejections": lambda: self.filter_rejections_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "airport clearance": lambda: self.filter_airport_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "airport combined": lambda: self.filter_airport_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "seaport clearance": lambda: self.filter_seaport_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "total lead time performance": lambda: self.filter_total_lead_time_performance(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "data logger": lambda: self.filter_data_logger_measurement(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "dock to stock": lambda: self.filter_dock_to_stock_combined(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "pods update": lambda: self.filter_pods_update(
                    request,
                    effective_month,
                    selected_months=quarter_months or None,
                ),
                "meeting points": lambda: self.meeting_points_tab(request),
                "cross docking": lambda: self.cross_docking_tab(request),
            }

            # Iterate available tab filters
            for key, func in tab_filter_map.items():
                if key in selected_tab:
                    print(f"📂 Executing tab filter: {key}")
                    try:
                        result = func()

                        # Direct HttpResponse/JsonResponse
                        if isinstance(result, HttpResponse):
                            print(
                                "ℹ️ Filter returned HttpResponse/JsonResponse; returning as-is."
                            )
                            return result

                        # Dict/list response → JSON
                        if isinstance(result, (dict, list)):
                            return JsonResponse(result, safe=False)

                        # String response (likely HTML)
                        if isinstance(result, str):
                            return JsonResponse({"detail_html": result}, safe=False)

                        # Fallback conversion
                        return JsonResponse({"detail_html": str(result)}, safe=False)

                    except Exception as e:
                        import traceback

                        print("❌ Error while executing tab filter:", key)
                        traceback.print_exc()
                        return JsonResponse(
                            {"error": f"Error in '{key}': {str(e)}"},
                            status=200,
                        )

            # Overview (never cached)
            if selected_tab == "overview":
                print("🔹 Loading Overview tab")
                overview_result = self.overview_tab(
                    request=request,
                    selected_month=effective_month,
                    selected_months=quarter_months or None,
                )
                return JsonResponse(overview_result, safe=True)

            # All-in-One (never cached)
            elif selected_tab == "all":
                print("🔹 Loading All-in-One tab")
                all_result = self.filter_all_tabs(
                    request=request,
                    selected_month=effective_month,
                    selected_months=quarter_months or None,
                )
                return JsonResponse(all_result, safe=False)

            # Remaining tabs
            elif selected_tab == "order general information":
                return JsonResponse(
                    self.filter_order_general_information(request, selected_month),
                    safe=False,
                )
            elif selected_tab in ["rejections", "return & refusal"]:
                return JsonResponse(
                    self.filter_rejections_combined(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            elif (
                "airport clearance" in selected_tab
                or "airport combined" in selected_tab
            ):
                return JsonResponse(
                    self.filter_airport_combined(request, effective_month), safe=False
                )
            elif selected_tab == "seaport clearance":
                return JsonResponse(
                    self.filter_seaport_combined(request, effective_month), safe=False
                )
            elif selected_tab in [
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
            elif "data logger" in selected_tab:
                return JsonResponse(
                    self.filter_data_logger_measurement(
                        request,
                        effective_month,
                        selected_months=quarter_months or None,
                    ),
                    safe=False,
                )
            elif "dock to stock - roche" in selected_tab:
                return JsonResponse(
                    self.filter_dock_to_stock_roche(request, effective_month),
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
            elif "cross docking" in selected_tab:
                return self.cross_docking_tab(request)
            elif selected_tab:
                raw_data = self.render_raw_sheet(request, selected_tab)
                return JsonResponse(raw_data, safe=False)
            else:
                return JsonResponse({"error": "⚠️ Please select a tab first."})

        # ====================== الطلب العادي ======================
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
            # التابات اللي تحب تاخدها من الداشبورد: عدّل هنا (أسماء كما في الإكسل)
            EXCLUDE_SHEETS_BASE = ["Sheet2"]
            # لو حابب تحذف تابات إضافية: زود أسمائهم هنا (بالضبط كما في الإكسل)
            EXCLUDE_SHEETS_EXTRA = getattr(
                self.__class__, "EXCLUDE_TABS", []
            )  # أو عدّل EXCLUDE_TABS في أول الكلاس
            EXCLUDE_SHEETS = list(EXCLUDE_SHEETS_BASE) + list(EXCLUDE_SHEETS_EXTRA)

            include_only = getattr(self.__class__, "INCLUDE_ONLY_TABS", None)
            if include_only:
                # عرض التابات المذكورة فقط (الاسم كما في الإكسل)
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
                "Return & Refusal",
                "Dock to stock",
                "Total Lead Time Performance",
                "Meeting Points & Action",
            ]
            if include_only:
                include_set_v = {s.strip() for s in include_only}
                filtered_tabs += [v for v in virtual_tabs if v in include_set_v]
            else:
                filtered_tabs += virtual_tabs

            ordered_tabs = [
                self.DASHBOARD_TAB_NAME,
                "Dock to stock",
                "Total Lead Time Performance",
                "PODs update",
                "Return & Refusal",
                "Meeting Points & Action",
            ]

            filtered_tabs = [tab for tab in ordered_tabs if tab in filtered_tabs]
            excel_tabs = [{"original": name, "display": name} for name in filtered_tabs]

        except Exception as e:
            print(f"⚠️ [ERROR] تعذر قراءة الشيتات من الملف: {e}")
            excel_tabs = []

        # ======================================================
        # 🗓️ استخراج كل الشهور من جميع الشيتات الممكنة
        # ======================================================
        all_months = set()
        try:
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
                    all_months.update(df["MonthName"].dropna().unique().tolist())
                except Exception as inner_e:
                    continue

            all_months = sorted(
                all_months, key=lambda m: pd.to_datetime(m, format="%b")
            )
            print("📅 [INFO] الشهور المستخرجة من كل الشيتات:", all_months)
        except Exception as e:
            print("⚠️ [ERROR] أثناء استخراج الشهور:", e)
            all_months = []

        meeting_points = MeetingPoint.objects.all().order_by("is_done", "-created_at")
        done_count = meeting_points.filter(is_done=True).count()
        total_count = meeting_points.count()

        overview_data = self.overview_tab(
            request=request, selected_month=selected_month or None
        )
        all_tab_data = self.filter_all_tabs(
            request=request, selected_month=selected_month or None
        )

        return render(
            request,
            self.template_name,
            {
                "data_is_uploaded": True,
                "months": all_months,
                "excel_tabs": excel_tabs,
                "active_tab": "all",
                "tab_summaries": [],
                "form": ExcelUploadForm(),
                "meeting_points": meeting_points,
                "done_count": done_count,
                "total_count": total_count,
                "all_tab_data": all_tab_data,
                "raw_tab_data": None,
            },
        )

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
        ext = (
            os.path.splitext(getattr(excel_file, "name", "") or "latest.xlsx")[1]
            or ".xlsx"
        )
        if ext.lower() not in (".xlsx", ".xlsm"):
            ext = ".xlsx"
        file_path = os.path.join(folder_path, "latest" + ext)

        try:
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
                    # ✅ محاولة تغيير الصلاحيات أولاً (على PythonAnywhere قد يكون الملف محمي)
                    os.chmod(file_path, 0o644)
                    os.remove(file_path)
                    print(f"🗑️ [DEBUG] تم حذف الملف القديم: {file_path}")
                except PermissionError as pe:
                    print(
                        f"⚠️ [DEBUG] تحذير: لا يمكن حذف الملف القديم (PermissionError): {pe}"
                    )
                    # ✅ محاولة حفظ الملف باسم مؤقت ثم استبداله
                    temp_path = os.path.join(folder_path, "latest_temp.xlsx")
                    with open(temp_path, "wb+") as destination:
                        for chunk in excel_file.chunks():
                            destination.write(chunk)
                    # ✅ محاولة استبدال الملف القديم بالجديد
                    try:
                        os.replace(temp_path, file_path)
                        print(f"✅ [DEBUG] تم استبدال الملف باستخدام os.replace")
                    except Exception as replace_error:
                        print(
                            f"⚠️ [DEBUG] تحذير: لا يمكن استبدال الملف: {replace_error}"
                        )
                        # ✅ إذا فشل الاستبدال، استخدم الملف المؤقت
                        file_path = temp_path
                except Exception as delete_error:
                    print(f"⚠️ [DEBUG] تحذير: خطأ في حذف الملف القديم: {delete_error}")
                    # ✅ المتابعة مع حفظ الملف الجديد (سيستبدل الملف القديم)

            # ✅ حفظ الملف الجديد
            with open(file_path, "wb+") as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            # ✅ التأكد من الصلاحيات الصحيحة للملف الجديد
            try:
                os.chmod(file_path, 0o644)
            except Exception as chmod_error:
                print(f"⚠️ [DEBUG] تحذير: لا يمكن تغيير صلاحيات الملف: {chmod_error}")

            print(f"✅ [DEBUG] تم حفظ الملف بنجاح في: {file_path}")

            # ✅ حفظ مسار الملف في الجلسة لاستخدامه في الجلسات التالية بدون إعادة الرفع
            request.session["uploaded_excel_path"] = file_path
            request.session.save()
            print(f"💾 [DEBUG] تم حفظ مسار الملف في الجلسة: {file_path}")

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
        تصدير الملف الأصلي المرفوع (Roche KPI new.xlsx) فقط مع الحفاظ على الألوان والتنسيق
        """
        from openpyxl import load_workbook

        # 📂 البحث عن الملف الأصلي "Roche KPI new.xlsx" في مجلد media
        folder_path = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
        original_excel_path = os.path.join(folder_path, "Roche KPI new.xlsx")

        # إذا لم يوجد، جرب البحث عن latest.xlsx كبديل
        if not os.path.exists(original_excel_path):
            latest_path = os.path.join(folder_path, "latest.xlsx")
            if os.path.exists(latest_path):
                original_excel_path = latest_path
                print(
                    f"📄 [EXPORT] تم العثور على latest.xlsx بدلاً من Roche KPI new.xlsx"
                )
            else:
                # جرب من الجلسة
                saved_path = request.session.get("uploaded_excel_path")
                if saved_path and os.path.exists(saved_path):
                    original_excel_path = saved_path
                    print(f"📄 [EXPORT] تم استخدام الملف من الجلسة: {saved_path}")
                else:
                    print(f"⚠️ [EXPORT] لم يتم العثور على الملف الأصلي")
                    return HttpResponse(
                        "❌ لم يتم العثور على الملف الأصلي (Roche KPI new.xlsx)",
                        status=404,
                    )

        try:
            print(f"📄 [EXPORT] جاري قراءة الملف الأصلي: {original_excel_path}")

            # قراءة الملف باستخدام openpyxl للحفاظ على التنسيق والألوان
            workbook = load_workbook(original_excel_path)

            # حفظ الملف في BytesIO مع الحفاظ على كل التنسيق
            output = BytesIO()
            workbook.save(output)
            output.seek(0)

            print(f"✅ [EXPORT] تم نسخ الملف الأصلي بنجاح مع الحفاظ على التنسيق")

            # إنشاء اسم الملف للتنزيل
            filename_parts = ["Roche KPI Dashboard Data"]
            if selected_months:
                filename_parts.append("-".join(selected_months))
            elif selected_month:
                filename_parts.append(selected_month)
            safe_filename = " ".join(filename_parts)

            response = HttpResponse(
                output.getvalue(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            response["Content-Disposition"] = (
                f'attachment; filename="{safe_filename}.xlsx"'
            )
            return response

        except Exception as e:
            print(f"⚠️ [EXPORT] حدث خطأ عند قراءة الملف الأصلي: {e}")
            import traceback

            traceback.print_exc()
            return HttpResponse(f"❌ حدث خطأ عند تصدير الملف: {str(e)}", status=500)

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

            clean_tabs = []
            for tab in overview_data.get("tab_cards", []):
                name = tab.get("name", "غير معروف")

                try:
                    hit = float(tab.get("hit_pct", 0))
                except:
                    hit = 0
                hit = int(round(max(0, min(hit, 100))))

                try:
                    target = float(tab.get("target_pct", 100))
                except:
                    target = 100

                clean_tabs.append(
                    {
                        "name": name,
                        "hit_pct": hit,
                        "target_pct": int(target),
                        "count": tab.get("count", 0),
                        "chart_type": tab.get("chart_type", "bar"),
                        "chart_data": tab.get("chart_data", []),
                    }
                )

            # ✅ Airport Clearance
            airport_data = self.filter_airport_combined(request, month_for_filters)
            if airport_data:
                hit_airport = airport_data.get("hit_pct", 0)
                try:
                    hit_airport = float(hit_airport)
                except:
                    hit_airport = 0
                hit_airport = int(round(max(0, min(hit_airport, 100))))

                existing_names = [t["name"].strip().lower() for t in clean_tabs]
                if "airport clearance" not in existing_names:
                    clean_tabs.append(
                        {
                            "name": "Airport Clearance",
                            "hit_pct": hit_airport,
                            "target_pct": 98,
                            "count": airport_data.get("count", 0),
                            "chart_type": "column",
                            "chart_data": [],
                        }
                    )

            # ✅ Seaport Clearance
            # ✅ Seaport Clearance
            seaport_data = self.filter_seaport_combined(
                request, selected_month=month_for_filters
            )
            if seaport_data and seaport_data.get("chart_data"):
                # خدي النسبة الحقيقية زي ما هي راجعة من التاب
                hit_seaport = seaport_data.get("hit_pct", 0)

                # تأكدي إنها عدد صحيح
                try:
                    hit_seaport = int(round(float(hit_seaport)))
                except:
                    hit_seaport = 0

                # لو أقل من صفر أو أكتر من 100 نصححها
                hit_seaport = max(0, min(hit_seaport, 100))

                existing_names = [t["name"].strip().lower() for t in clean_tabs]
                if "seaport clearance" not in existing_names:
                    clean_tabs.append(
                        {
                            "name": "Seaport Clearance",
                            "hit_pct": hit_seaport,
                            "target_pct": 98,
                            "count": seaport_data.get("count", 0),
                            "chart_type": "column",
                            "chart_data": seaport_data.get("chart_data", []),
                        }
                    )

            # ✅ PODs Update
            pods_data = self.filter_pods_update(request, month_for_filters)
            if pods_data and pods_data.get("hit_pct") is not None:
                try:
                    hit_pods = float(pods_data.get("hit_pct", 0))
                except:
                    hit_pods = 0
                hit_pods = int(round(max(0, min(hit_pods, 100))))

                existing_names = [t["name"].strip().lower() for t in clean_tabs]
                if "pods update" not in existing_names:
                    clean_tabs.append(
                        {
                            "name": "PODs update",
                            "hit_pct": hit_pods,
                            "target_pct": 98,
                            "count": pods_data.get("count", 0),
                            "chart_type": "column",
                            "chart_data": [],
                        }
                    )

            # ✅ ترتيب التابات حسب الأولوية
            desired_order = [
                "Total Lead Time Performance",
                "Return & Refusal",
                "Data Logger Measurement",
                "Dock to stock",
                "Airport Clearance",
                "Seaport Clearance",
                "PODs update",
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

    def filter_inbound(self, request, selected_month=None, selected_months=None):
        """
        🔹 يحسب KPI لشحنات الـ Inbound (≤24 ساعة بين Create Timestamp و Last LPN Rcv TS).
        🔹 يعيد جدول ملخص شهري + جدول تفصيلي للشحنات.
        """
        try:
            import os

            excel_path = self.get_uploaded_file_path(request) or self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next(
                (s for s in xls.sheet_names if "inbound" in s.lower()), None
            )
            if not sheet_name:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Sheet containing 'Inbound' was not found.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            if df.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Inbound sheet is empty.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

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
                    if any(normalize_name(name) in col_norm for name in possible_names):
                        return col
                return None

            column_aliases = {
                "facility": [
                    "facility code",
                    "facility",
                    "facilitycode",
                ],
                "shipment": [
                    "shipment nbr",
                    "shipment number",
                    "shipment no",
                    "shipment#",
                    "shipment id",
                ],
                "status": ["status", "shipment status"],
                "create_ts": ["create timestamp", "created timestamp", "creation time"],
                "arrival": ["arrival date", "arrival timestamp"],
                "offload": ["offloading date", "offload date", "offload timestamp"],
                "last_lpn": [
                    "last lpn rcv ts",
                    "last lpn receive ts",
                    "last lpn rcv timestamp",
                    "last lpn rcv",
                ],
                "reason": [
                    "reason",
                    "miss reason",
                    "delay reason",
                    "remarks",
                    "comments",
                    "cause",
                    "reason code",
                ],
            }

            column_map = {
                key: find_column(names) for key, names in column_aliases.items()
            }

            required_keys = ["shipment", "status", "create_ts", "last_lpn"]
            missing_required = [key for key in required_keys if not column_map.get(key)]
            if missing_required:
                missing_labels = ", ".join(missing_required)
                return {
                    "detail_html": f"<p class='text-danger'>⚠️ Missing required columns for inbound analysis: {missing_labels}</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            rename_map = {}
            if column_map.get("facility"):
                rename_map[column_map["facility"]] = "Facility Code"
            if column_map["shipment"]:
                rename_map[column_map["shipment"]] = "Shipment Nbr"
            if column_map["status"]:
                rename_map[column_map["status"]] = "Status"
            if column_map["create_ts"]:
                rename_map[column_map["create_ts"]] = "Create Timestamp"
            if column_map["arrival"]:
                rename_map[column_map["arrival"]] = "Arrival Date"
            if column_map["offload"]:
                rename_map[column_map["offload"]] = "Offloading Date"
            if column_map["last_lpn"]:
                rename_map[column_map["last_lpn"]] = "Last LPN Rcv TS"
            if column_map.get("reason"):
                rename_map[column_map["reason"]] = "Reason"

            df = df.rename(columns=rename_map)
            if "Reason" not in df.columns:
                df["Reason"] = ""
            if "Facility Code" not in df.columns:
                df["Facility Code"] = ""

            for col in [
                "Create Timestamp",
                "Arrival Date",
                "Offloading Date",
                "Last LPN Rcv TS",
            ]:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                else:
                    df[col] = pd.NaT

            df["Shipment Nbr"] = df["Shipment Nbr"].astype(str).str.strip()
            df["Status"] = df["Status"].astype(str).str.strip()
            if "Facility Code" in df.columns:
                df["Facility Code"] = df["Facility Code"].astype(str).str.strip()

            df["Cycle Hours"] = (
                (df["Last LPN Rcv TS"] - df["Create Timestamp"])
                .dt.total_seconds()
                .div(3600)
            )
            df["Cycle Hours"] = df["Cycle Hours"].round(2)
            df["Cycle Days"] = (df["Cycle Hours"] / 24).round(2)

            df["is_hit"] = df["Cycle Hours"].le(24)
            df["is_hit"] = df["is_hit"] & df["Cycle Hours"].notna()

            df["HIT or MISS"] = np.where(df["is_hit"], "Hit", "Miss")
            df.loc[df["Cycle Hours"].isna(), "HIT or MISS"] = "Pending"

            month_source = df["Create Timestamp"].copy()
            month_source = month_source.fillna(df["Arrival Date"])
            month_source = month_source.fillna(df["Offloading Date"])
            df["Month"] = month_source.dt.strftime("%b")

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
                    .fillna("")
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
            elif selected_month_norm:
                df = df[
                    df["Month"].fillna("").str.lower() == selected_month_norm.lower()
                ]

            if df.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ No inbound records for the selected period.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            df_summary = df.dropna(subset=["Month"]).copy()
            if df_summary.empty:
                return {
                    "detail_html": "<p class='text-warning'>⚠️ Inbound data has no valid month values.</p>",
                    "sub_tables": [],
                    "chart_data": [],
                }

            def month_order_value(label):
                if not label:
                    return 999
                label = label.strip()[:3].title()
                for idx in range(1, 13):
                    if month_abbr[idx] == label:
                        return idx
                return 999

            total_per_month = (
                df_summary.groupby("Month")["Shipment Nbr"]
                .nunique()
                .reset_index(name="Total_Shipments")
            )
            hits_df = (
                df_summary[df_summary["is_hit"]]
                .groupby("Month")["Shipment Nbr"]
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
            summary_df = summary_df.sort_values(
                by="Month", key=lambda col: col.map(month_order_value)
            )

            months_with_miss = summary_df[summary_df["Misses"] > 0]["Month"].tolist()
            months_with_hit_only = summary_df[summary_df["Misses"] == 0][
                "Month"
            ].tolist()
            ordered_months = months_with_miss + months_with_hit_only

            df_miss = df_summary[~df_summary["is_hit"]].copy()
            df_miss["Reason"] = (
                df_miss.get("Reason", pd.Series([""] * len(df_miss)))
                .astype(str)
                .str.strip()
            )
            df_miss.loc[df_miss["Reason"].isin(["", "nan", "NaN"]), "Reason"] = (
                "(No reason)"
            )

            reason_pivot = None
            if not df_miss.empty and "Reason" in df_miss.columns:
                reason_counts = (
                    df_miss.groupby(["Reason", "Month"])["Shipment Nbr"]
                    .nunique()
                    .reset_index(name="cnt")
                )
                if not reason_counts.empty:
                    reason_pivot = reason_counts.pivot_table(
                        index="Reason", columns="Month", values="cnt", fill_value=0
                    ).reset_index()
                    for m in ordered_months:
                        if m not in reason_pivot.columns:
                            reason_pivot[m] = 0
                    reason_pivot = reason_pivot.reindex(
                        columns=["Reason"]
                        + [c for c in ordered_months if c in reason_pivot.columns]
                    )

            kpi_rows = []
            for _, row in summary_df.iterrows():
                m = row["Month"]
                kpi_rows.append(
                    {
                        "Month": m,
                        "Total Shipments": int(row["Total_Shipments"]),
                        "Hit (≤24h)": int(row["Hits"]),
                        "Miss (>24h)": int(row["Misses"]),
                        "Hit %": float(row["Hit %"]),
                        "_has_miss": row["Misses"] > 0,
                    }
                )

            pivot_cols = ["KPI"] + ordered_months
            if len(ordered_months) >= 2:
                pivot_cols.append("2025")

            hit_pct_row = {"KPI": "Hit %"}
            total_row = {"KPI": "Total Shipments"}
            hit_row = {"KPI": "Hit (≤24h)"}
            miss_row = {"KPI": "Miss (>24h)"}
            for m in ordered_months:
                r = next((x for x in kpi_rows if x["Month"] == m), None)
                if r:
                    total_val = int(r["Total Shipments"])
                    hit_val = int(r["Hit (≤24h)"])
                    miss_val = int(r["Miss (>24h)"])
                    total_row[m] = total_val
                    hit_row[m] = hit_val
                    miss_row[m] = miss_val
                    hit_pct_row[m] = (
                        int(round(hit_val / total_val * 100)) if total_val > 0 else 0
                    )
            if "2025" in pivot_cols:
                total_2025 = sum(r["Total Shipments"] for r in kpi_rows)
                hit_2025 = sum(r["Hit (≤24h)"] for r in kpi_rows)
                hit_pct_row["2025"] = (
                    int(round(hit_2025 / total_2025 * 100)) if total_2025 > 0 else 0
                )
                total_row["2025"] = int(sum(r["Total Shipments"] for r in kpi_rows))
                hit_row["2025"] = int(sum(r["Hit (≤24h)"] for r in kpi_rows))
                miss_row["2025"] = int(sum(r["Miss (>24h)"] for r in kpi_rows))

            summary_data_pivot = [hit_pct_row, total_row, hit_row, miss_row]

            if reason_pivot is not None and not reason_pivot.empty:
                for _, r in reason_pivot.iterrows():
                    reason_row = {"KPI": str(r["Reason"])}
                    for c in ordered_months:
                        if c in r.index:
                            reason_row[c] = int(r[c]) if pd.notna(r[c]) else 0
                    if "2025" in pivot_cols:
                        reason_row["2025"] = int(
                            sum(
                                int(r[c]) if c in r.index and pd.notna(r[c]) else 0
                                for c in ordered_months
                            )
                        )
                    summary_data_pivot.append(reason_row)

            def _to_display_int(val):
                if val is None or (
                    isinstance(val, float) and (np.isnan(val) or np.isinf(val))
                ):
                    return 0
                if isinstance(val, (int, np.integer)):
                    return int(val)
                try:
                    return int(round(float(val)))
                except (ValueError, TypeError):
                    return 0

            for row in summary_data_pivot:
                for k in list(row.keys()):
                    if k != "KPI" and isinstance(
                        row[k], (int, float, np.integer, np.floating)
                    ):
                        row[k] = _to_display_int(row[k])

            summary_columns = pivot_cols
            summary_data = summary_data_pivot

            overall_total = int(df.shape[0])
            overall_hits = int(df["is_hit"].sum())
            overall_miss = overall_total - overall_hits
            overall_hit_pct = (
                round((overall_hits / overall_total) * 100, 2) if overall_total else 0
            )

            chart_data = []
            hit_pct_for_chart = next(
                (r for r in summary_data if r.get("KPI") == "Hit %"), None
            )
            if hit_pct_for_chart:
                data_points = []
                for m in ordered_months:
                    v = hit_pct_for_chart.get(m)
                    if v is not None and isinstance(v, (int, float)):
                        data_points.append({"label": m, "y": float(v)})
                if "2025" in hit_pct_for_chart:
                    data_points.append(
                        {"label": "2025", "y": float(hit_pct_for_chart["2025"])}
                    )
                chart_data.append(
                    {
                        "type": "column",
                        "name": "Inbound Hit %",
                        "color": "#74c0fc",
                        "related_table": "sub-table-inbound-hit-summary",
                        "dataPoints": data_points,
                    }
                )

            detail_columns = [
                "Facility Code",
                "Shipment Nbr",
                "Status",
                "Create Timestamp",
                "Arrival Date",
                "Offloading Date",
                "Last LPN Rcv TS",
                "Days",
                "HIT or MISS",
            ]

            detail_df = df.copy()
            detail_df["_sort_ts"] = detail_df["Create Timestamp"]

            def _fmt_date(x):
                if pd.isna(x) or x is pd.NaT:
                    return ""
                try:
                    return pd.Timestamp(x).strftime("%Y-%m-%d %H:%M")
                except Exception:
                    return ""

            for col in [
                "Create Timestamp",
                "Arrival Date",
                "Offloading Date",
                "Last LPN Rcv TS",
            ]:
                if col in detail_df.columns:
                    detail_df[col] = detail_df[col].apply(_fmt_date)
                else:
                    detail_df[col] = ""

            detail_df["Days"] = detail_df["Cycle Days"].apply(
                lambda x: "" if pd.isna(x) else str(int(np.ceil(float(x))))
            )

            if "Facility Code" not in detail_df.columns:
                detail_df["Facility Code"] = ""
            if "HIT or MISS" not in detail_df.columns:
                detail_df["HIT or MISS"] = ""

            for col in detail_columns:
                if col not in detail_df.columns:
                    detail_df[col] = ""

            drop_cols = [
                c
                for c in ["_sort_ts", "Cycle Hours", "Cycle Days", "is_hit"]
                if c in detail_df.columns
            ]
            detail_rows_raw = (
                detail_df.sort_values("_sort_ts", ascending=False)
                .drop(columns=drop_cols)
                .head(500)[detail_columns]
                .to_dict(orient="records")
            )

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

            months_with_miss_label = (
                " — Months with Miss: " + ", ".join(months_with_miss)
                if months_with_miss
                else " — All months Hit"
            )
            summary_table = {
                "id": "sub-table-inbound-hit-summary",
                "title": "Inbound KPI ≤ 24h" + months_with_miss_label,
                "columns": summary_columns,
                "data": summary_data,
                "chart_data": chart_data,
                "months_with_miss": months_with_miss,
                "months_with_hit_only": months_with_hit_only,
            }

            # طباعة جدول KPI (Inbound KPI ≤ 24h) في الترمينال
            print("\n" + "=" * 80)
            print("Inbound KPI ≤ 24h — summary_table data")
            print("=" * 80)
            print("Columns:", summary_columns)
            print("-" * 80)
            for row in summary_data:
                kpi_name = row.get("KPI", "")
                rest = {k: v for k, v in row.items() if k != "KPI"}
                print(f"  KPI: {kpi_name}")
                for col, val in rest.items():
                    print(f"    {col}: {val}")
                print("-" * 80)
            print("=" * 80 + "\n")

            detail_table = {
                "id": "sub-table-inbound-detail",
                "title": "Inbound Shipments Detail",
                "columns": detail_columns,
                "data": detail_rows,
                "chart_data": [],
                "full_width": True,
            }

            # طباعة جدول Inbound Shipments Detail في الترمينال
            print("\n" + "=" * 100)
            print(
                "Inbound Shipments Detail — Create Timestamp | Arrival Date | Offloading Date | Status | HIT or MISS"
            )
            print("=" * 100)
            for i, row in enumerate(detail_rows[:50], 1):  # أول 50 صف
                create_ts = str(row.get("Create Timestamp", ""))[:16]
                arrival = str(row.get("Arrival Date", ""))[:16]
                offload = str(row.get("Offloading Date", ""))[:16]
                status = str(row.get("Status", ""))[:18]
                hit_miss = str(row.get("HIT or MISS", ""))
                print(
                    f"  {i:3d} | Create: {create_ts:16s} | Arrival: {arrival:16s} | Offload: {offload:16s} | Status: {status:18s} | {hit_miss}"
                )
            if len(detail_rows) > 50:
                print(f"  ... و {len(detail_rows) - 50} صف إضافي")
            print("=" * 100 + "\n")

            return {
                "detail_html": "",
                "sub_tables": [summary_table, detail_table],
                "chart_data": chart_data,
                "stats": {
                    "total": overall_total,
                    "hit": overall_hits,
                    "miss": overall_miss,
                    "hit_pct": overall_hit_pct,
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
    def filter_data_logger_measurement(
        self, request, selected_month=None, selected_months=None
    ):
        print("🟢 [DEBUG] ✅ دخل على filter_data_logger_measurement()")

        try:
            import pandas as pd, numpy as np, os, traceback
            from django.template.loader import render_to_string

            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {"error": "⚠️ Excel file not found."}

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next(
                (s for s in xls.sheet_names if "data logger measurement" in s.lower()),
                None,
            )
            if not sheet_name:
                return {
                    "error": "⚠️ No sheet containing 'Data logger Measurement' was found."
                }

            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.str.strip()

            if "Month" not in df.columns:
                return {"error": "⚠️ Column 'Month' was not found in the sheet."}

            month_order = [
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "2025",
            ]

            # تنظيف البيانات
            df = df.dropna(subset=["Month"])
            df = df[df["Month"].isin(month_order[:-1])]

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
                self.normalize_month_label(selected_month) if selected_month else None
            )
            if selected_months_norm:
                df = df[
                    df["Month"]
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning text-center p-4'>⚠️ No data available for the selected quarter months in Data Logger Measurement.</p>",
                        "chart_data": [],
                        "count": 0,
                        "hit_pct": 0,
                    }
            elif selected_month_norm:
                df = df[df["Month"].str.lower() == selected_month_norm.lower()]
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for {selected_month_norm}.</p>",
                        "chart_data": [],
                        "count": 0,
                    }

            # Pivot
            df_melted = df.melt(id_vars=["Month"], var_name="KPI", value_name="Value")
            df_pivot = df_melted.pivot_table(
                index="KPI",
                columns="Month",
                values="Value",
                aggfunc="sum",
                fill_value=0,
            )
            month_scope = (
                selected_months_norm if selected_months_norm else month_order[:-1]
            )
            df_pivot = df_pivot[[m for m in month_scope if m in df_pivot.columns]]
            df_pivot["2025"] = df_pivot.sum(axis=1)

            # صف TOTAL
            total_row = df_pivot.sum(numeric_only=True)
            total_row.name = "TOTAL"
            if "Late send to ROCHE" in df_pivot.index:
                new_df = pd.DataFrame(columns=df_pivot.columns)
                for idx in df_pivot.index:
                    new_df.loc[idx] = df_pivot.loc[idx]
                    if idx == "Late send to ROCHE":
                        new_df.loc["TOTAL"] = total_row
                df_pivot = new_df
            else:
                df_pivot.loc["TOTAL"] = total_row

            # حساب النسب
            ontime_percentages = []
            months_iterable = [m for m in df_pivot.columns if m != "2025"]

            for month in months_iterable:
                if month not in df_pivot.columns:
                    continue
                on_time = (
                    pd.to_numeric(df_pivot.at["On time sent", month], errors="coerce")
                    if "On time sent" in df_pivot.index
                    else 0
                )
                total = (
                    pd.to_numeric(df_pivot.at["TOTAL", month], errors="coerce")
                    if "TOTAL" in df_pivot.index
                    else 0
                )
                pct = round((on_time / total * 100), 2) if total else 0
                ontime_percentages.append(pct)

            # ✅ فلترة الشهور اللي كلها أصفار (قبل تجهيز الشارت)
            valid_months = []
            valid_percentages = []

            for i, month in enumerate(months_iterable):
                # نحول كل القيم في الشهر إلى أرقام للتأكد إن مفيش قيم نصية
                col_values = pd.to_numeric(df_pivot[month], errors="coerce").fillna(0)
                if col_values.sum() != 0:  # الشهر فيه بيانات حقيقية
                    valid_months.append(month)
                    valid_percentages.append(ontime_percentages[i])

            print(f"📊 [DEBUG] الشهور المعروضة في الشارت فقط: {valid_months}")

            # 🎯 الشارت فقط للـ valid_months
            chart_data = [
                {
                    "type": "bar",
                    "name": "On Time Sent %",
                    "showInLegend": True,
                    "color": "#d0e7ff",
                    "dataPoints": [
                        {"label": m, "y": valid_percentages[i]}
                        for i, m in enumerate(valid_months)
                    ],
                },
                {
                    "type": "line",
                    "name": "Target (100%)",
                    "color": "red",
                    "showInLegend": True,
                    "dataPoints": [{"label": m, "y": 100} for m in valid_months],
                },
            ]

            # الجدول كما هو (ما نحذفش منه حاجة)
            table_data = df_pivot.reset_index().to_dict(orient="records")
            columns = ["KPI"] + [m for m in month_order if m in df_pivot.columns]

            tab = {
                "name": "Data Logger Measurement",
                "columns": columns,
                "data": table_data,
                "chart_data": chart_data,  # ✅ إضافة chart_data للتاب
            }

            month_norm = self.apply_month_filter_to_tab(tab, selected_month)
            html = render_to_string(
                "forms-table/table/bootstrap-table/basic-table/components/excel-sheet-table.html",
                {
                    "tab": tab,
                    "table_title": "Data Logger Measurement",
                    "selected_month": month_norm,
                },
            )

            # ✅ حساب hit الإجمالية
            try:
                total_on_time = (
                    pd.to_numeric(df_pivot.loc["On time sent"], errors="coerce")
                    .fillna(0)
                    .sum()
                )
                total_all = (
                    pd.to_numeric(df_pivot.loc["TOTAL"], errors="coerce")
                    .fillna(0)
                    .sum()
                )
                hit_pct = (
                    round((total_on_time / total_all * 100), 2) if total_all else 0
                )
            except Exception:
                hit_pct = 0

            return {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "On Time Sent % Performance",
                "hit_pct": hit_pct,  # ✅ أضفناها هنا
                "tab_data": tab,
            }

        except Exception as e:
            print("❌ [ERROR] Exception in filter_data_logger_measurement():")
            traceback.print_exc()
            return {"error": f"⚠️ Error while reading data: {e}"}

    def filter_pods_update(self, request, selected_month=None, selected_months=None):
        """
        🔹 قراءة شيت PODs Update من Excel
        🔹 عرض الأشهر كما هي في الشيت (بدون ترتيب ثابت)
        🔹 عرض جدول Closed / Pending / Total فقط
        🔹 رسم شارت النسبة المئوية لـ Closed فقط (%)
        🔹 حساب نسبة الأداء النهائية (YTD %)
        """
        print("🟢 [DEBUG] ✅ دخل على filter_pods_update()")

        try:
            import pandas as pd
            from django.template.loader import render_to_string
            import os

            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {"error": "⚠️ Excel file not found."}

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next((s for s in xls.sheet_names if "pod" in s.lower()), None)
            if not sheet_name:
                return {"error": "⚠️ No sheet containing the word 'POD' was found."}

            # 🔹 قراءة البيانات من الشيت
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = df.columns.str.strip()

            print(f"🔍 [DEBUG PODs] الأعمدة الموجودة: {df.columns.tolist()}")
            print(f"🔍 [DEBUG PODs] عدد الصفوف: {len(df)}")
            print(f"🔍 [DEBUG PODs] أول 5 صفوف:\n{df.head()}")

            # 🔹 التأكد من الأعمدة الأساسية
            columns_map = {}
            for col in df.columns:
                name = col.strip().lower()
                if "month" in name:
                    columns_map["Month"] = col
                elif "closed" in name:
                    columns_map["Closed"] = col
                elif "pending" in name:
                    columns_map["Pending"] = col

            if len(columns_map) < 3:
                return {"error": "⚠️ Required columns are missing."}

            df = df.rename(columns=columns_map)
            df = df.dropna(subset=["Month"])
            df = df[~df["Month"].astype(str).str.lower().eq("total")]

            # ✅ استخدام الأشهر الموجودة في Excel كما هي
            df["MonthAbbr"] = df["Month"].astype(str).str.strip().str.title()

            # ✅ فلترة الشهر المختار (لو موجود)
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

            if selected_months_norm:
                df = df[
                    df["MonthAbbr"]
                    .str.lower()
                    .isin([m.lower() for m in selected_months_norm])
                ]
                print(
                    f"🔍 [DEBUG PODs] فلترة الكوارتر: {', '.join(selected_months_norm)}, الصفوف: {len(df)}"
                )

                if df.empty:
                    return {
                        "detail_html": "<p class='text-warning text-center p-4'>⚠️ No data available for the selected quarter months in PODs Update.</p>",
                        "count": 0,
                        "hit_pct": 0,
                    }
                selected_month = None

            if selected_month:
                selected_month_norm = self.normalize_month_label(selected_month)
                if selected_month_norm:
                    df = df[df["MonthAbbr"].str.lower() == selected_month_norm.lower()]
                    print(
                        f"🔍 [DEBUG PODs] فلترة الشهر: {selected_month_norm}, الصفوف: {len(df)}"
                    )

                    if df.empty:
                        return {
                            "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for {selected_month_norm} in PODs Update.</p>",
                            "count": 0,
                            "hit_pct": 0,
                        }

            # 🔹 استخراج القيم
            closed_values = df["Closed"].tolist()
            pending_values = df["Pending"].tolist()
            total_values = [c + p for c, p in zip(closed_values, pending_values)]

            # ✅ حساب المجاميع السنوية
            closed_sum = sum(closed_values)
            pending_sum = sum(pending_values)
            total_sum = sum(total_values)

            # ✅ ترتيب الأشهر حسب ما جاءت في Excel (بعد الفلترة إن وُجدت)
            month_order_display = df["MonthAbbr"].tolist()

            selected_month_norm = (
                self.normalize_month_label(selected_month) if selected_month else None
            )
            if selected_month_norm:
                df = df[df["MonthAbbr"].str.lower() == selected_month_norm.lower()]
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for {selected_month_norm}.</p>",
                        "chart_data": [],
                        "count": 0,
                    }
                month_order_display = df["MonthAbbr"].tolist()

            # ✅ إضافة صف المجموع (YTD) لو مفيش فلترة
            append_ytd = not bool(selected_month_norm or selected_months_norm)
            if append_ytd:
                month_order_display = month_order_display + ["YTD"]

            if append_ytd:
                closed_values.append(closed_sum)
                pending_values.append(pending_sum)
                total_values.append(total_sum)

            # ✅ حساب النسب المئوية (للشارت فقط)
            closed_percent_values = [
                round((c / t) * 100, 2) if t != 0 else 0
                for c, t in zip(closed_values, total_values)
            ]

            print("📊 Closed % per month:", closed_percent_values)

            # ✅ بناء الجدول النهائي بدون صف KPI %
            table_data = pd.DataFrame({"KPI": ["Closed", "Pending", "Total"]})
            for i, month in enumerate(month_order_display):
                col_values = [
                    int(closed_values[i]) if i < len(closed_values) else 0,
                    int(pending_values[i]) if i < len(pending_values) else 0,
                    int(total_values[i]) if i < len(total_values) else 0,
                ]
                table_data[month] = col_values

            # ✅ نسبة الأداء النهائية (YTD)
            hit_pct = closed_percent_values[-1] if closed_percent_values else 0

            # ✅ تجهيز بيانات الشارت
            chart_data = [
                {
                    "type": "column",
                    "name": "Closed %",
                    "color": "#9fc0e4",
                    "showInLegend": True,
                    "indexLabel": "{y}%",
                    "related_table": "PODs YTD",  # ✅ إضافة related_table
                    "dataPoints": [
                        {"label": m, "y": closed_percent_values[i]}
                        for i, m in enumerate(month_order_display)
                    ],
                },
                {
                    "type": "line",
                    "name": "Target (%)",
                    "color": "red",
                    "showInLegend": True,
                    "related_table": "PODs YTD",  # ✅ إضافة related_table
                    "dataPoints": [{"label": m, "y": 100} for m in month_order_display],
                },
            ]

            # ✅ بناء الجدول للعرض
            sub_tables = [
                {
                    "id": "sub-table-pods-ytd",  # ✅ إضافة ID فريد
                    "title": "PODs YTD",
                    "columns": table_data.columns.tolist(),
                    "data": table_data.to_dict(orient="records"),
                    "chart_data": chart_data,  # ✅ إضافة chart_data لكل sub_table
                }
            ]

            tab_data = {
                "name": "PODs Update",
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "chart_title": "PODs Closed % Performance",
                "hit_pct": hit_pct,
                "target_pct": 100,
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

            tab_name = f"PODs Update ({hit_pct}%)"

            return {
                "name": tab_name,
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "PODs Closed % Performance",
                "hit_pct": hit_pct,
                "target_pct": 100,
                "count": len(table_data),
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {"error": f"⚠️ Error while processing data: {e}"}

    def filter_order_general_information(
        self, request, selected_month=None, selected_months=None
    ):
        """
        🔹 دمج عرض 4 شيتات:
            1️⃣ Urgent Orders Details ← جدول + شارت (% From Full Order Number)
            2️⃣ Outbound Details ← جدول + شارت (% Of Normal Orders + % Of Booking Orders)
            3️⃣ Picking Accuracy ← جدول + شارت (% Of Actual)
            4️⃣ Quality Exceptions (Temperature deviation سابقاً)
        🔹 عرض جميع البيانات بدون فلترة بالشهر
        """
        print("🟢 [DEBUG] ✅ دخل على filter_order_general_information()")

        try:
            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {"error": "⚠️ Excel file not found."}

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sub_tables = []

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

            sheets_info = {
                "Urgent orders details": "Urgent Orders Details",
                "Outbound details": "Outbound Details",
                "Picking Accuracy": "Picking Accuracy",
                "Temperature deviation": "Quality Exceptions",
            }

            for sheet_name, table_title in sheets_info.items():
                if sheet_name not in xls.sheet_names:
                    continue

                try:
                    df = pd.read_excel(
                        excel_path, sheet_name=sheet_name, engine="openpyxl"
                    )
                    df.columns = df.columns.str.strip().str.title()
                    df = df.dropna(how="all")

                    if df.empty:
                        sub_tables.append(
                            {
                                "id": f"sub-table-{slugify(table_title)}",
                                "title": table_title,
                                "columns": [],
                                "data": [],
                                "chart_data": [],
                                "message": "لا توجد بيانات متاحة.",
                            }
                        )
                        continue

                    # ✅ حفظ نسخة من البيانات الأصلية قبل التحويل (لإنشاء chart_data)
                    df_original = df.copy()

                    # ✅ تحويل النسب المئوية إلى أعداد صحيحة بدون علامة عشرية (مثل 27% وليس 27.00%)
                    for col in df.select_dtypes(include=["float", "float64"]).columns:
                        df[col] = df[col].apply(
                            lambda x: f"{int(round(x * 100))}%" if pd.notna(x) else ""
                        )

                    # ✅ إنشاء chart_data بناءً على نوع الشيت
                    chart_data = []

                    # البحث عن عمود Month أو الشهر
                    month_col = None
                    for col in df_original.columns:
                        if col.lower() in ["month", "monthabbr", "month_abbr"]:
                            month_col = col
                            break

                    # 1️⃣ Urgent Orders Details - الشارت يعرض "% From Full Order Number"
                    if table_title == "Urgent Orders Details":
                        pct_col = None
                        for col in df_original.columns:
                            col_lower = str(col).lower()
                            # البحث عن عمود يحتوي على "%" و "full" أو "order" و "number"
                            if "%" in str(col) and (
                                "full" in col_lower
                                or ("order" in col_lower and "number" in col_lower)
                            ):
                                pct_col = col
                                break
                        # إذا لم نجد، جرب البحث عن أي عمود يحتوي على "%"
                        if not pct_col:
                            for col in df_original.columns:
                                if "%" in str(col) and col != month_col:
                                    pct_col = col
                                    break

                        if month_col and pct_col:
                            data_points = []
                            for _, row in df_original.iterrows():
                                month_val = (
                                    str(row[month_col]).strip()
                                    if pd.notna(row[month_col])
                                    else ""
                                )
                                pct_val = row[pct_col]
                                if pd.notna(pct_val):
                                    # تحويل من نسبة (0.27) إلى نسبة مئوية (27)
                                    if isinstance(pct_val, (int, float)):
                                        pct_val = (
                                            pct_val * 100 if pct_val <= 1 else pct_val
                                        )
                                    else:
                                        try:
                                            pct_val = float(
                                                str(pct_val).replace("%", "")
                                            )
                                        except:
                                            continue
                                    data_points.append(
                                        {"label": month_val, "y": round(pct_val, 2)}
                                    )

                            if data_points:
                                chart_data.append(
                                    {
                                        "type": "column",
                                        "name": "% From Full Order Number",
                                        "color": "#9fc0e4",
                                        "showInLegend": True,
                                        "dataPoints": data_points,
                                        "related_table": table_title,
                                    }
                                )

                    # 2️⃣ Outbound Details - الشارت يعرض "% Of Normal Orders" و "% Of Booking Orders"
                    elif table_title == "Outbound Details":
                        normal_col = None
                        booking_col = None

                        # البحث الأول: البحث عن الأعمدة المحددة
                        for col in df_original.columns:
                            col_lower = str(col).lower()
                            col_str = str(col)
                            if "%" in col_str:
                                if "normal" in col_lower:
                                    normal_col = col
                                elif "booking" in col_lower:
                                    booking_col = col

                        # Fallback: إذا لم نجد الأعمدة المحددة، نبحث عن أي عمود يحتوي على "%"
                        if not normal_col and not booking_col:
                            for col in df_original.columns:
                                col_str = str(col)
                                if "%" in col_str and col != month_col:
                                    # نأخذ أول عمود نسبة مئوية كـ fallback
                                    if not normal_col:
                                        normal_col = col
                                    elif not booking_col:
                                        booking_col = col
                                    if normal_col and booking_col:
                                        break

                        if month_col:
                            # Normal Orders
                            if normal_col:
                                data_points_normal = []
                                for _, row in df_original.iterrows():
                                    month_val = (
                                        str(row[month_col]).strip()
                                        if pd.notna(row[month_col])
                                        else ""
                                    )
                                    val = row[normal_col]
                                    if pd.notna(val):
                                        if isinstance(val, (int, float)):
                                            val = val * 100 if val <= 1 else val
                                        else:
                                            try:
                                                val = float(str(val).replace("%", ""))
                                            except:
                                                continue
                                        data_points_normal.append(
                                            {"label": month_val, "y": round(val, 2)}
                                        )

                                if data_points_normal:
                                    chart_data.append(
                                        {
                                            "type": "stackedColumn100",  # ✅ تغيير إلى stackedColumn100 للـ stacked chart
                                            "name": (
                                                normal_col
                                                if normal_col
                                                else "% Of Normal Orders"
                                            ),
                                            "color": "#9084ad",
                                            "showInLegend": True,
                                            "dataPoints": data_points_normal,
                                            "related_table": table_title,
                                            "stack": "stack1",  # ✅ إضافة stack name للـ stacked chart
                                        }
                                    )

                            # Booking Orders
                            if booking_col:
                                data_points_booking = []
                                for _, row in df_original.iterrows():
                                    month_val = (
                                        str(row[month_col]).strip()
                                        if pd.notna(row[month_col])
                                        else ""
                                    )
                                    val = row[booking_col]
                                    if pd.notna(val):
                                        if isinstance(val, (int, float)):
                                            val = val * 100 if val <= 1 else val
                                        else:
                                            try:
                                                val = float(str(val).replace("%", ""))
                                            except:
                                                continue
                                        data_points_booking.append(
                                            {"label": month_val, "y": round(val, 2)}
                                        )

                                if data_points_booking:
                                    chart_data.append(
                                        {
                                            "type": "stackedColumn100",  # ✅ تغيير إلى stackedColumn100 للـ stacked chart
                                            "name": (
                                                booking_col
                                                if booking_col
                                                else "% Of Booking Orders"
                                            ),
                                            "color": "#9fc0e4",
                                            "showInLegend": True,
                                            "dataPoints": data_points_booking,
                                            "related_table": table_title,
                                            "stack": "stack1",  # ✅ إضافة stack name للـ stacked chart (نفس الاسم)
                                        }
                                    )

                            # ✅ إذا لم نجد أي عمود، نستخدم أول عمود نسبة مئوية كـ fallback
                            if not chart_data and normal_col:
                                data_points = []
                                for _, row in df_original.iterrows():
                                    month_val = (
                                        str(row[month_col]).strip()
                                        if pd.notna(row[month_col])
                                        else ""
                                    )
                                    val = row[normal_col]
                                    if pd.notna(val):
                                        if isinstance(val, (int, float)):
                                            val = val * 100 if val <= 1 else val
                                        else:
                                            try:
                                                val = float(str(val).replace("%", ""))
                                            except:
                                                continue
                                        data_points.append(
                                            {"label": month_val, "y": round(val, 2)}
                                        )

                                if data_points:
                                    chart_data.append(
                                        {
                                            "type": "column",
                                            "name": normal_col,
                                            "color": "#4caf50",
                                            "showInLegend": True,
                                            "dataPoints": data_points,
                                            "related_table": table_title,
                                        }
                                    )

                    # 3️⃣ Picking Accuracy - الشارت يعرض "% Of Actual"
                    elif table_title == "Picking Accuracy":
                        actual_col = None
                        for col in df_original.columns:
                            col_lower = str(col).lower()
                            col_str = str(col)
                            if "%" in col_str and "actual" in col_lower:
                                actual_col = col
                                break
                        # إذا لم نجد، جرب البحث عن أي عمود يحتوي على "%" و "actual"
                        if not actual_col:
                            for col in df_original.columns:
                                col_str = str(col)
                                if "%" in col_str and col != month_col:
                                    actual_col = col
                                    break

                        if month_col and actual_col:
                            data_points = []
                            for _, row in df_original.iterrows():
                                month_val = (
                                    str(row[month_col]).strip()
                                    if pd.notna(row[month_col])
                                    else ""
                                )
                                val = row[actual_col]
                                if pd.notna(val):
                                    if isinstance(val, (int, float)):
                                        val = val * 100 if val <= 1 else val
                                    else:
                                        try:
                                            val = float(str(val).replace("%", ""))
                                        except:
                                            continue
                                    data_points.append(
                                        {"label": month_val, "y": round(val, 2)}
                                    )

                            if data_points:
                                chart_data.append(
                                    {
                                        "type": "column",
                                        "name": "% Of Actual",
                                        "color": "#9fc0e4",
                                        "showInLegend": True,
                                        "dataPoints": data_points,
                                        "related_table": table_title,
                                    }
                                )

                    # 4️⃣ Quality Exceptions - بدون شارت (جدول فقط في المنتصف)
                    elif table_title == "Quality Exceptions":
                        # ✅ لا نضيف chart_data لهذا الجدول
                        chart_data = []

                    # ✅ إنشاء sub_table مع ID فريد و chart_data
                    sub_table_id = f"sub-table-{slugify(table_title)}"
                    sub_table = {
                        "id": sub_table_id,
                        "title": table_title,
                        "columns": list(df.columns),
                        "data": df.to_dict(orient="records"),
                        "chart_data": chart_data,  # ✅ إضافة chart_data
                    }
                    sub_tables.append(sub_table)

                    print(
                        f"✅ تمت معالجة الشيت: {sheet_name} → title='{table_title}' → chart_data: {len(chart_data)} datasets"
                    )
                    if chart_data:
                        print(
                            f"   📊 Chart datasets: {[ds.get('name', 'N/A') for ds in chart_data]}"
                        )
                        print(f"   📊 Month column found: {month_col}")
                        print(
                            f"   📊 Data points count: {sum(len(ds.get('dataPoints', [])) for ds in chart_data)}"
                        )
                    else:
                        print(f"   ⚠️ No chart_data created for {table_title}")
                        print(f"   📊 Available columns: {list(df_original.columns)}")
                        print(f"   📊 Month column: {month_col}")

                except Exception as e:
                    print(f"⚠️ خطأ أثناء قراءة الشيت {sheet_name}: {e}")
                    sub_tables.append(
                        {
                            "id": f"sub-table-{slugify(table_title)}",
                            "title": table_title,
                            "columns": [],
                            "data": [],
                            "chart_data": [],  # ✅ إضافة chart_data فارغ
                            "message": f"⚠️ خطأ أثناء قراءة البيانات: {e}",
                        }
                    )

            if not sub_tables:
                return {"error": "⚠️ No valid data was found in any sheets."}

            tab_data = {
                "name": "Order General Information",
                "sub_tables": sub_tables,
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

            total_count = sum(len(st["data"]) for st in sub_tables)

            return {
                "name": "Order General Information",
                "detail_html": html,
                "count": total_count,
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {"error": f"⚠️ Error while processing data: {e}"}

    def filter_rejections_combined(
        self, request, selected_month=None, selected_months=None
    ):
        """
        🔹 عرض 3 شيتات:
            1️⃣ Rejection
            2️⃣ Rejection breakdown
            3️⃣ Return
        🔹 الجداول تُعرض كما في الشيت (القيم، النسب، التنسيق كما هو)
        🔹 شارت دائري يعرض نسب Booking orders من شيت Rejection
        """
        import pandas as pd
        import os
        from django.template.loader import render_to_string

        print("🟣 [DEBUG] ✅ دخل على filter_rejections_combined()")

        try:
            excel_path = self.get_excel_path()
            print("📁 [DEBUG] excel_path:", excel_path)

            if not excel_path or not os.path.exists(excel_path):
                print("❌ [DEBUG] لم يتم العثور على ملف Excel.")
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "chart_data": [],
                    "count": 0,
                }

            # ✅ قراءة الملف
            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_names = [s.strip() for s in xls.sheet_names]
            print("🧾 [DEBUG] الشيتات الموجودة:", sheet_names)

            sub_tables = []
            chart_data = []
            color_palette = [
                "#007fa3",
                "#ff4d4d",
                "#ffa500",
                "#28a745",
                "#6f42c1",
                "#17a2b8",
                "#ffc107",
                "#fd7e14",
                "#20c997",
                "#6610f2",
                "#e83e8c",
                "#343a40",
            ]

            sheets_to_show = ["Rejection", "Rejection breakdown", "Return"]

            # دالة لتحويل القيم الرقمية إلى نسب مئوية نصية (للعرض فقط)
            # دالة لتحويل القيم الرقمية إلى نسب مئوية نصية (بدون كسور)
            def to_percentage_display(val):
                """تحويل رقم مثل 0.06 إلى '6%' أو 0.085 إلى '9%' بدون كسور"""
                if val is None or str(val).strip() == "":
                    return ""
                try:
                    num = float(val)
                    if num <= 1:
                        num = num * 100
                    # 🔹 تحويل إلى عدد صحيح بدون فواصل
                    num_int = int(round(num))
                    return f"{num_int}%"
                except:
                    return str(val)

            # دالة للحصول على القيمة الرقمية للشارت فقط (كعدد صحيح)
            def to_percentage_number(val):
                """تحويل القيمة إلى int بين 0 و100 بشكل آمن"""
                try:
                    if val is None:
                        return None
                    val_str = str(val).replace("%", "").strip()
                    if val_str == "":
                        return None
                    num = float(val_str)
                    if num <= 1:
                        num *= 100
                    return int(round(num))
                except Exception:
                    return None

            # 🔁 قراءة كل شيت متوقع
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

            for expected_name in sheets_to_show:
                matched_sheet = next(
                    (s for s in sheet_names if expected_name.lower() in s.lower()), None
                )

                if not matched_sheet:
                    print(f"⚠️ [DEBUG] الشيت {expected_name} غير موجود.")
                    sub_tables.append(
                        {
                            "title": expected_name,
                            "columns": [],
                            "data": [],
                            "error": f"Sheet '{expected_name}' not found",
                        }
                    )
                    continue

                print(f"📄 [DEBUG] قراءة الشيت: {matched_sheet}")
                try:
                    df = pd.read_excel(
                        excel_path,
                        sheet_name=matched_sheet,
                        engine="openpyxl",
                        dtype=str,
                        header=0,
                    ).fillna("")

                    df.columns = df.columns.astype(str).str.strip()

                    # ✅ إعادة تسمية الأعمدة في شيت Rejection
                    if "rejection" in matched_sheet.lower():
                        column_rename_map = {}
                        for col in df.columns:
                            col_lower = col.lower().strip()
                            if "normal orders" in col_lower:
                                column_rename_map[col] = "Number of Rejection"
                            elif "booking orders" in col_lower:
                                column_rename_map[col] = "% of Rejection"
                        if column_rename_map:
                            df = df.rename(columns=column_rename_map)

                    # ✅ فلترة الشهر المختار (لو موجود)
                    if selected_month or selected_months_norm:
                        selected_month_norm = self.normalize_month_label(selected_month)
                        # البحث عن عمود Month في الشيت
                        month_col_candidates = [
                            c
                            for c in df.columns
                            if c.strip().lower()
                            in ["month", "monthabbr", "month abbreviation"]
                        ]
                        active_filters = set()
                        if selected_months_norm:
                            active_filters = {m.lower() for m in selected_months_norm}
                        elif selected_month_norm:
                            active_filters = {selected_month_norm.lower()}

                        if month_col_candidates and active_filters:
                            month_col = month_col_candidates[0]

                            def _normalize_cell(val):
                                result = self.normalize_month_label(val)
                                return result

                            df["_normalized_month"] = df[month_col].apply(
                                _normalize_cell
                            )
                            df = df[
                                df["_normalized_month"].notna()
                                & (
                                    df["_normalized_month"]
                                    .str.lower()
                                    .isin(active_filters)
                                )
                            ]
                            df = df.drop(columns=["_normalized_month"])
                            print(
                                f"🔍 [DEBUG Rejections] فلترة {matched_sheet} للشهور: {active_filters}, عدد الصفوف: {len(df)}"
                            )

                    # ✅ تعديل عرض عمود % of Rejection في شيت Rejection فقط
                    if (
                        "rejection" in matched_sheet.lower()
                        and "% of Rejection" in df.columns
                    ):
                        df["% of Rejection"] = df["% of Rejection"].apply(
                            to_percentage_display
                        )

                    # ✅ تجهيز صف التوتال وبيانات شارت Return (إن وجد)
                    if "return" in matched_sheet.lower():
                        month_names = [
                            "jan",
                            "feb",
                            "mar",
                            "apr",
                            "may",
                            "jun",
                            "jul",
                            "aug",
                            "sep",
                            "oct",
                            "nov",
                            "dec",
                        ]
                        month_cols = [
                            c
                            for c in df.columns
                            if str(c).strip().lower() in month_names
                        ]

                        if month_cols:
                            total_row = {col: "" for col in df.columns}
                            key_col = df.columns[0] if len(df.columns) else None
                            if key_col:
                                total_row[key_col] = "Total"

                            chart_points = []
                            allowed_months = (
                                set(m.lower() for m in selected_months_norm)
                                if selected_months_norm
                                else None
                            )
                            for idx, col in enumerate(month_cols):
                                if (
                                    allowed_months
                                    and str(col).lower() not in allowed_months
                                ):
                                    continue
                                numeric_series = (
                                    pd.to_numeric(df[col], errors="coerce")
                                    .fillna(0)
                                    .astype(float)
                                )
                                df[col] = numeric_series.round().astype(int)
                                total_value = int(numeric_series.sum().round())
                                total_row[col] = total_value
                                chart_points.append(
                                    {
                                        "label": col,
                                        "y": total_value,
                                        "color": color_palette[
                                            idx % len(color_palette)
                                        ],
                                    }
                                )

                            df = pd.concat(
                                [df, pd.DataFrame([total_row])], ignore_index=True
                            )

                            if chart_points:
                                chart_data.append(
                                    {
                                        "type": "column",
                                        "name": "Return Totals",
                                        "dataPoints": chart_points,
                                        "color": "#007fa3",
                                        "related_table": matched_sheet,
                                    }
                                )

                    # ✅ حفظ الجدول كما هو (نفس الأعمدة الأصلية)
                    sub_tables.append(
                        {
                            "title": matched_sheet,
                            "columns": df.columns.tolist(),
                            "data": df.to_dict(orient="records"),
                        }
                    )

                    # ✅ تجهيز بيانات الشارت (لكن بدون تعديل الجدول)
                    if (
                        "rejection" in matched_sheet.lower()
                        and "% of Rejection" in df.columns
                    ):
                        chart_points = []
                        allowed_labels = (
                            set(m.lower() for m in selected_months_norm)
                            if selected_months_norm
                            else None
                        )
                        for i, row in df.iterrows():
                            label = str(row.get("Month", f"Row {i + 1}")).strip()
                            raw_val = row.get("% of Rejection", "")
                            value = to_percentage_number(raw_val)
                            if (
                                allowed_labels
                                and label
                                and label.lower() not in allowed_labels
                            ):
                                continue
                            if value is not None:  # ✅ تأكد أن القيمة ليست None
                                chart_points.append(
                                    {
                                        "label": label or f"Row {i + 1}",
                                        "y": value,
                                        "color": color_palette[i % len(color_palette)],
                                    }
                                )

                        if chart_points:
                            chart_data.append(
                                {
                                    "type": "doughnut",
                                    "name": "Booking Orders %",
                                    "showInLegend": True,
                                    "dataPoints": chart_points,
                                }
                            )

                except Exception as e:
                    import traceback

                    print(traceback.format_exc())
                    sub_tables.append(
                        {
                            "title": matched_sheet,
                            "columns": [],
                            "data": [],
                            "error": str(e),
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

            # 🧩 بناء الـ HTML
            tab_data = {
                "name": "Return & Refusal",
                "sub_tables": sub_tables,
                "chart_data": chart_data,
                "chart_title": "Return & Refusal Overview",
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

            # 🧮 حساب hit% من متوسط القيم في % of Rejection (بالنسبة المئوية)
            hit_values = []
            for st in sub_tables:
                if "rejection" in st["title"].lower():
                    for row in st["data"]:
                        val = row.get("% of Rejection", "")
                        try:
                            num = to_percentage_number(val)
                            if num is not None:
                                hit_values.append(num)
                        except:
                            pass

            hit_pct = round(sum(hit_values) / len(hit_values), 2) if hit_values else 0

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

    def filter_airport_combined(
        self, request, selected_month=None, selected_months=None
    ):
        """
        Clearance Pivot Table + Chart لصفحة Airport
        - Hit Summary
        - Roch Delay Reasons
        - Transit KPI Summary
        """
        cache.clear()
        try:
            import pandas as pd
            from django.template.loader import render_to_string
            import os

            # 🧾 قراءة ملف الإكسل
            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "count": 0,
                }

            # ✅ قراءة شيت Airport Clearance (مع سقوط إذا غير موجود)
            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            sheet_name = next(
                (s for s in xls.sheet_names if "airport clearance" in s.lower()), None
            )
            if not sheet_name:
                return {
                    "detail_html": "<p class='text-warning text-center'>⚠️ Sheet 'Airport Clearance' was not found in the uploaded workbook.</p>",
                    "chart_data": [],
                    "count": 0,
                    "hit_pct": 0,
                }
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # 🧹 تنظيف الأعمدة
            df.columns = (
                df.columns.str.strip()
                .str.replace(r"[\n\r\t]+", "", regex=True)
                .str.replace(r"\s+", " ", regex=True)
            )
            df.rename(columns={"Clearanace Remarks": "Clearance Remarks"}, inplace=True)

            # ✅ الأعمدة المطلوبة الأساسية
            required_cols = ["Month", "Clearance Handling KPI", "MBL/AWB"]
            for col in required_cols:
                if col not in df.columns:
                    return {
                        "detail_html": f"<p class='text-danger'>⚠️ Column '{col}' does not exist in the sheet.</p>",
                        "count": 0,
                    }

            # 🧹 تنظيف القيم وتحويل الشهر
            df["Month"] = pd.to_datetime(df["Month"], errors="coerce")
            df["MonthAbbr"] = df["Month"].dt.strftime(
                "%b"
            )  # 👈 الشهر في شكل مختصر مثل Jan, Feb, Oct

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

            if selected_months_norm:
                month_lower_list = [m.lower() for m in selected_months_norm]
                df = df[df["MonthAbbr"].str.lower().isin(month_lower_list)]
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning text-center'>⚠️ No data available for months {', '.join(selected_months_norm)} in Airport Clearance.</p>",
                        "count": 0,
                        "hit_pct": 0,
                    }
                selected_month = None

            # ✅ فلترة الشهر المختار (لو موجود)
            if selected_month:
                selected_month_norm = str(selected_month).strip().capitalize()
                df = df[df["MonthAbbr"].str.lower() == selected_month_norm.lower()]
                print(f"🔍 [DEBUG] فلترة البيانات للشهر: {selected_month_norm}")
                print(f"🔍 [DEBUG] عدد الصفوف بعد الفلترة: {len(df)}")

                # ✅ التحقق من وجود بيانات بعد الفلترة
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning'>⚠️ No data available for {selected_month_norm}.</p>",
                        "count": 0,
                        "hit_pct": 0,
                    }

            # ✅ طباعة كل قيم الشهور الموجودة في الشيت (الأصلية والمختصرة)
            print("🚀 [DEBUG] Original Month column values:")
            print(df["Month"].dropna().unique())

            print("🚀 [DEBUG] Extracted Month Abbreviations:")
            print(df["MonthAbbr"].dropna().unique())

            df["Clearance Handling KPI"] = (
                df["Clearance Handling KPI"].astype(str).str.strip()
            )
            df["MBL/AWB"] = df["MBL/AWB"].astype(str).str.strip()

            # ===========================================================
            # ✅ 1️⃣ Hit Summary
            # ===========================================================
            df_hit = df[
                df["Clearance Handling KPI"].str.lower() == "hit"
            ].drop_duplicates(subset=["MBL/AWB", "Month"])
            df_hit["KPI"] = "On-time Delivery"

            month_labels = (
                df_hit.dropna(subset=["Month"])
                .sort_values("Month")["MonthAbbr"]
                .unique()
                .tolist()
            )

            pivot_hit = df_hit.pivot_table(
                index="KPI",
                columns="MonthAbbr",
                values="MBL/AWB",
                aggfunc="count",
                fill_value=0,
            ).reset_index()

            # ✅ صف Total
            total_row = {"KPI": "Total"}
            for m in month_labels:
                total_row[m] = int(pivot_hit[m].sum() if m in pivot_hit.columns else 0)
            pivot_hit = pd.concat(
                [pivot_hit, pd.DataFrame([total_row])], ignore_index=True
            )

            # ✅ حساب Hit %
            hit_percent = {}
            for m in month_labels:
                hit = pivot_hit.at[0, m] if m in pivot_hit.columns else 0
                total = pivot_hit.at[1, m] if m in pivot_hit.columns else 1
                hit_percent[m] = int((hit / total * 100) if total > 0 else 0)

            # ✅ تجهيز الجدول والشارت
            hit_chart = [
                {
                    "title": "Airport Clearance (Hit Summary)",
                    "type": "column",
                    "name": "Hit %",
                    "color": "#9fc0e4",
                    "dataPoints": [
                        {"label": m, "y": hit_percent[m]} for m in month_labels
                    ],
                    "related_table": "Airport Clearance (Hit Summary)",
                },
                {
                    "title": "Airport Clearance (Hit Summary)",
                    "type": "line",
                    "name": "Target (98%)",
                    "color": "#a3d977",
                    "markerSize": 5,
                    "dataPoints": [{"label": m, "y": 98} for m in month_labels],
                    "related_table": "Airport Clearance (Hit Summary)",
                },
            ]

            hit_table = {
                "id": "airport-hit-summary",
                "title": "Airport Clearance (Hit Summary)",
                "columns": pivot_hit.columns.tolist(),
                "data": pivot_hit.to_dict(orient="records"),
                "chart_data": hit_chart,
            }

            # ===========================================================
            # ✅ 2️⃣ Roch Delay Reasons
            # ===========================================================
            required_roch_cols = [
                "Month",
                "Clearance Handling KPI",
                "MBL/AWB",
                "Clearance Group",
                "Clearance Remarks",
            ]
            for col in required_roch_cols:
                if col not in df.columns:
                    return {
                        "detail_html": f"<p class='text-danger'>⚠️ Column '{col}' does not exist in the sheet.</p>",
                        "count": 0,
                    }

            df_roch = df[
                (df["Clearance Handling KPI"].str.lower() == "miss")
                & (df["Clearance Group"].str.lower().str.contains("roch", na=False))
            ].drop_duplicates(subset=["MBL/AWB"])

            if selected_month:
                df_roch = df_roch[
                    df_roch["MonthAbbr"].str.lower() == selected_month.lower()
                ]

            if df_roch.empty:
                roch_table = {
                    "id": "airport-roch-delay-reasons",
                    "title": "Roch Delay Reasons",
                    "columns": ["Reason", "Count"],
                    "data": [{"Reason": "لا توجد أسباب تأخير", "Count": "-"}],
                    "chart_data": [],  # لا يوجد شارت لهذا الجدول
                }
            else:
                month_label = (
                    selected_month.capitalize()
                    if selected_month
                    else df_roch["MonthAbbr"].iloc[0]
                )
                reasons_count = (
                    df_roch.groupby("Clearance Remarks")["MBL/AWB"]
                    .nunique()
                    .reset_index()
                    .rename(
                        columns={"Clearance Remarks": "Reason", "MBL/AWB": month_label}
                    )
                    .sort_values(by=month_label, ascending=False)
                )
                roch_table = {
                    "id": "airport-roch-delay-reasons",
                    "title": "Roch Delay Reasons",
                    "columns": reasons_count.columns.tolist(),
                    "data": reasons_count.to_dict(orient="records"),
                    "chart_data": [],  # لا يوجد شارت لهذا الجدول
                }

            # ===========================================================
            # ✅ 3️⃣ Transit KPI Summary
            # ===========================================================
            required_transit_cols = [
                "Month",
                "Transit KPI",
                "MBL/AWB",
                "Transit Group",
                "Transit Remarks",
            ]
            for col in required_transit_cols:
                if col not in df.columns:
                    return {
                        "detail_html": f"<p class='text-danger'>⚠️ Column '{col}' does not exist in the sheet.</p>",
                        "count": 0,
                    }

            df_transit = df.copy()
            df_transit["Month"] = pd.to_datetime(df_transit["Month"], errors="coerce")
            df_transit["MonthAbbr"] = df_transit["Month"].dt.strftime(
                "%b"
            )  # 👈 تحويل الشهر
            df_transit = df_transit.dropna(subset=["Month"])

            # فلترة الشهر لو المستخدم اختار
            if selected_month:
                df_transit = df_transit[
                    df_transit["MonthAbbr"].str.lower() == selected_month.lower()
                ]

            months_to_show = (
                df_transit.sort_values("Month")["MonthAbbr"].unique().tolist()
            )

            df_hit_t = df_transit[
                df_transit["Transit KPI"].str.lower() == "hit"
            ].drop_duplicates(subset=["MBL/AWB"])
            hit_counts = {
                m: int(df_hit_t[df_hit_t["MonthAbbr"] == m]["MBL/AWB"].nunique())
                for m in months_to_show
            }

            df_miss_t = df_transit[
                df_transit["Transit KPI"].str.lower() == "miss"
            ].drop_duplicates(subset=["MBL/AWB"])
            if not df_miss_t.empty:
                pivot_miss = (
                    df_miss_t.groupby(
                        ["Transit Remarks", "Transit Group", "MonthAbbr"]
                    )["MBL/AWB"]
                    .nunique()
                    .unstack(fill_value=0)
                    .reset_index()
                )
            else:
                pivot_miss = pd.DataFrame(
                    columns=["Transit Remarks", "Transit Group"] + months_to_show
                )

            rows = [
                {
                    "KPI (Delay Reason)": "On-time Delivery",
                    "Responsible": "Tamer",
                    **{m: hit_counts.get(m, 0) for m in months_to_show},
                }
            ]
            for _, row in pivot_miss.iterrows():
                rows.append(
                    {
                        "KPI (Delay Reason)": row["Transit Remarks"],
                        "Responsible": row["Transit Group"],
                        **{m: int(row.get(m, 0)) for m in months_to_show},
                    }
                )

            total_row = {"KPI (Delay Reason)": "Total", "Responsible": "-"}
            for m in months_to_show:
                total_row[m] = sum(r[m] for r in rows if m in r)
            rows.append(total_row)

            pivot_transit_display = pd.DataFrame(
                rows, columns=["KPI (Delay Reason)", "Responsible"] + months_to_show
            )

            transit_hit_percent = {
                m: int((rows[0][m] / total_row[m]) * 100) if total_row[m] > 0 else 0
                for m in months_to_show
            }

            transit_chart = [
                {
                    "title": f"Transit KPI (Hit %) - {sheet_name}",
                    "type": "column",
                    "name": "Transit Hit %",
                    "color": "#b699d3",
                    "dataPoints": [
                        {"label": m, "y": transit_hit_percent[m]}
                        for m in months_to_show
                    ],
                    "related_table": "Transit KPI Summary",
                },
                {
                    "title": f"Transit KPI (Target) - {sheet_name}",
                    "type": "line",
                    "name": "Target (98%)",
                    "color": "#f44336",
                    "markerSize": 5,
                    "dataPoints": [{"label": m, "y": 98} for m in months_to_show],
                    "related_table": "Transit KPI Summary",
                },
            ]

            transit_table = {
                "id": "airport-transit-kpi-summary",
                "title": "Transit KPI Summary",  # ✅ استخدام title ثابت بدون f-string
                "columns": pivot_transit_display.columns.tolist(),
                "data": pivot_transit_display.to_dict(orient="records"),
                "chart_data": transit_chart,  # ✅ إضافة chart_data للجدول
            }

            # ✅ Debug: طباعة معلومات transit_table
            print(f"🔍 [DEBUG Airport] transit_table title: {transit_table['title']}")
            print(f"🔍 [DEBUG Airport] transit_table id: {transit_table['id']}")
            print(f"🔍 [DEBUG Airport] transit_chart count: {len(transit_chart)}")
            print(
                f"🔍 [DEBUG Airport] transit_table chart_data count: {len(transit_table.get('chart_data', []))}"
            )
            if transit_table.get("chart_data"):
                print(
                    f"🔍 [DEBUG Airport] transit_table chart_data names: {[ds.get('name', 'N/A') for ds in transit_table['chart_data']]}"
                )
                print(
                    f"🔍 [DEBUG Airport] transit_table chart_data types: {[ds.get('type', 'N/A') for ds in transit_table['chart_data']]}"
                )

            # ===========================================================
            # ✅ تجميع كل الجداول والشارتات
            # ===========================================================
            tab_data = {
                "name": sheet_name,
                "sub_tables": [hit_table, roch_table, transit_table],
                "chart_data": hit_chart + transit_chart,
                "chart_title": f"{sheet_name} Overview",
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

            # ✅ حساب hit_pct مع تجنب القسمة على صفر
            total_count = df.shape[0] if df.shape[0] > 0 else 1
            hit_pct = (
                int(round((df_hit.shape[0] / total_count) * 100, 0))
                if total_count > 0
                else 0
            )

            return {
                "detail_html": html,
                "count": df_hit.shape[0],
                "hit_pct": hit_pct,
                "chart_data": hit_chart + transit_chart,
                "chart_title": f"{sheet_name} Overview",
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing data: {e}</p>",
                "count": 0,
            }

    def filter_seaport_combined(
        self, request, selected_month=None, selected_months=None
    ):
        """
        Seaport Clearance - Hit Summary + Roch Delay Reasons
        """
        cache.clear()
        try:
            import pandas as pd
            from django.template.loader import render_to_string
            import os

            excel_path = self.get_excel_path()
            if not excel_path or not os.path.exists(excel_path):
                return {
                    "detail_html": "<p class='text-danger'>⚠️ Excel file not found.</p>",
                    "count": 0,
                }

            # تنظيف أسماء الشيتات
            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            cleaned_sheets = {s.strip().lower(): s for s in xls.sheet_names}
            target_sheet_key = "seaport clearance"
            if target_sheet_key not in cleaned_sheets:
                return {
                    "detail_html": f"<p class='text-danger'>⚠️ Sheet '{target_sheet_key}' does not exist.</p>",
                    "count": 0,
                }
            sheet_name = cleaned_sheets[target_sheet_key]

            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")

            # تنظيف الأعمدة
            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.replace(r"[\n\r\t]+", "", regex=True)
                .str.replace(r"\s+", " ", regex=True)
            )

            # العثور على الأعمدة الأساسية
            container_col_candidates = [
                c for c in df.columns if c.strip().lower() == "container no."
            ]
            if not container_col_candidates:
                return {
                    "detail_html": "⚠️ Column 'Container No.' does not exist.",
                    "count": 0,
                }
            container_col = container_col_candidates[0]
            container_index = df.columns.get_loc(container_col)
            if container_index + 1 >= len(df.columns):
                return {
                    "detail_html": "⚠️ Column 'Month' was not found after 'Container No.'",
                    "count": 0,
                }
            month_col = df.columns[container_index + 1]

            kpi_candidates = [
                c for c in df.columns if c.strip().lower() == "clearance handling kpi"
            ]
            if not kpi_candidates:
                return {
                    "detail_html": "⚠️ Column 'Clearance Handling KPI' does not exist.",
                    "count": 0,
                }
            kpi_col = kpi_candidates[0]

            # عمود Group و Remarks
            group_candidates = [
                c for c in df.columns if c.strip().lower() == "clearance group"
            ]
            if not group_candidates:
                return {
                    "detail_html": "⚠️ Column 'Clearance Group' does not exist.",
                    "count": 0,
                }
            group_col = group_candidates[0]

            remarks_candidates = [
                c for c in df.columns if c.strip().lower() == "clearance remarks"
            ]
            if not remarks_candidates:
                return {
                    "detail_html": "⚠️ Column 'Clearance Remarks' does not exist.",
                    "count": 0,
                }
            remarks_col = remarks_candidates[0]

            # تنظيف القيم
            df[container_col] = df[container_col].astype(str).str.strip()
            df[month_col] = pd.to_datetime(df[month_col], errors="coerce")
            df["MonthAbbr"] = df[month_col].dt.strftime("%b")
            df[kpi_col] = df[kpi_col].astype(str).str.strip()
            df[group_col] = df[group_col].astype(str).str.strip()
            df[remarks_col] = df[remarks_col].astype(str).str.strip()

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

            if selected_months_norm:
                month_lower_list = [m.lower() for m in selected_months_norm]
                df = df[df["MonthAbbr"].str.lower().isin(month_lower_list)]
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for months {', '.join(selected_months_norm)} in Seaport Clearance.</p>",
                        "count": 0,
                        "hit_pct": 0,
                        "chart_data": [],
                    }
                selected_month = None

            # ✅ فلترة الشهر المختار (لو موجود)
            if selected_month:
                selected_month_norm = str(selected_month).strip().capitalize()
                df = df[df["MonthAbbr"].str.lower() == selected_month_norm.lower()]
                print(f"🔍 [DEBUG Seaport] فلترة البيانات للشهر: {selected_month_norm}")
                print(f"🔍 [DEBUG Seaport] عدد الصفوف بعد الفلترة: {len(df)}")

                # ✅ التحقق من وجود بيانات بعد الفلترة
                if df.empty:
                    return {
                        "detail_html": f"<p class='text-warning text-center p-4'>⚠️ No data available for {selected_month_norm} in Seaport Clearance.</p>",
                        "count": 0,
                        "hit_pct": 0,
                        "chart_data": [],
                    }

            # ===========================================================
            # 1️⃣ Hit Summary
            # ===========================================================
            df_hit = df[df[kpi_col].str.lower() == "hit"].drop_duplicates(
                subset=[container_col, month_col]
            )
            df_hit["KPI"] = "On-time Delivery"
            month_labels = (
                df_hit.dropna(subset=[month_col])
                .sort_values(month_col)["MonthAbbr"]
                .unique()
                .tolist()
            )

            pivot_hit = df_hit.pivot_table(
                index="KPI",
                columns="MonthAbbr",
                values=container_col,
                aggfunc="count",
                fill_value=0,
            ).reset_index()

            total_row = {"KPI": "Total"}
            for m in month_labels:
                total_row[m] = int(pivot_hit[m].sum() if m in pivot_hit.columns else 0)
            pivot_hit = pd.concat(
                [pivot_hit, pd.DataFrame([total_row])], ignore_index=True
            )

            # ===========================================================
            # Hit % لكل شهر فقط
            # ===========================================================
            hit_percent = {}
            print("⚡ On-time Delivery % لكل شهر:")
            for m in month_labels:
                hit = (
                    pivot_hit.at[0, m] if m in pivot_hit.columns else 0
                )  # Hit = On-time Delivery
                total = (
                    pivot_hit.at[1, m] if m in pivot_hit.columns else 1
                )  # Total في الشهر
                percent = int((hit / total) * 100) if total > 0 else 0
                hit_percent[m] = percent
                print(f"  {m}: {percent}%")  # يطبع النسبة في التيرمينال

            hit_chart = [
                {
                    "title": "Seaport Clearance (Hit Summary)",
                    "type": "column",
                    "name": "Hit %",
                    "color": "#9fc0e4",
                    "dataPoints": [
                        {"label": m, "y": hit_percent[m]} for m in month_labels
                    ],
                    "related_table": "Seaport Clearance (Hit Summary)",
                },
                {
                    "title": "Seaport Clearance (Hit Summary)",
                    "type": "line",
                    "name": "Target (98%)",
                    "color": "#a3d977",
                    "markerSize": 5,
                    "dataPoints": [{"label": m, "y": 98} for m in month_labels],
                    "related_table": "Seaport Clearance (Hit Summary)",
                },
            ]

            hit_table = {
                "id": "seaport-hit-summary",
                "title": "Seaport Clearance (Hit Summary)",
                "columns": pivot_hit.columns.tolist(),
                "data": pivot_hit.to_dict(orient="records"),
                "chart_data": hit_chart,
            }

            # ✅ Debug: طباعة معلومات hit_table
            print(f"🔍 [DEBUG Seaport] hit_table title: {hit_table['title']}")
            print(f"🔍 [DEBUG Seaport] hit_table id: {hit_table['id']}")
            print(f"🔍 [DEBUG Seaport] hit_chart count: {len(hit_chart)}")
            print(
                f"🔍 [DEBUG Seaport] hit_table chart_data count: {len(hit_table.get('chart_data', []))}"
            )
            if hit_table.get("chart_data"):
                print(
                    f"🔍 [DEBUG Seaport] hit_table chart_data names: {[ds.get('name', 'N/A') for ds in hit_table['chart_data']]}"
                )
                print(
                    f"🔍 [DEBUG Seaport] hit_table chart_data types: {[ds.get('type', 'N/A') for ds in hit_table['chart_data']]}"
                )

            # ===========================================================
            # 2️⃣ Roch Delay Reasons
            # ===========================================================
            df_miss = df[df[kpi_col].str.lower() == "miss"].drop_duplicates(
                subset=[container_col, month_col]
            )
            if selected_month:
                df_miss = df_miss[
                    df_miss["MonthAbbr"].str.lower() == selected_month.lower()
                ]

            if df_miss.empty:
                roch_table = {
                    "id": "seaport-roch-delay-reasons",
                    "title": "Roch Delay Reasons",
                    "columns": ["Reason", "Count"],
                    "data": [{"Reason": "لا توجد أسباب تأخير", "Count": "-"}],
                    "chart_data": [],  # لا يوجد شارت لهذا الجدول
                }
            else:
                month_label = (
                    selected_month.capitalize()
                    if selected_month
                    else df_miss["MonthAbbr"].iloc[0]
                )
                reasons_count = (
                    df_miss.groupby(remarks_col)[container_col]
                    .nunique()
                    .reset_index()
                    .rename(columns={remarks_col: "Reason", container_col: month_label})
                    .sort_values(by=month_label, ascending=False)
                )
                roch_table = {
                    "id": "seaport-roch-delay-reasons",
                    "title": "Roch Delay Reasons",
                    "columns": reasons_count.columns.tolist(),
                    "data": reasons_count.to_dict(orient="records"),
                    "chart_data": [],  # لا يوجد شارت لهذا الجدول
                }

            # ===========================================================
            # 3️⃣ Transit KPI Summary
            # ===========================================================
            transit_kpi_col_candidates = [
                c for c in df.columns if c.strip().lower() == "transit kpi"
            ]
            transit_group_col_candidates = [
                c for c in df.columns if c.strip().lower() == "transit group"
            ]
            transit_remarks_col_candidates = [
                c for c in df.columns if c.strip().lower() == "transit remarks"
            ]
            transit_container_candidates = [
                c for c in df.columns if c.strip().lower() == "container no."
            ]

            if not transit_kpi_col_candidates:
                return {
                    "detail_html": "⚠️ Column 'Transit KPI' does not exist.",
                    "count": 0,
                }
            if not transit_group_col_candidates:
                return {
                    "detail_html": "⚠️ Column 'Transit Group' does not exist.",
                    "count": 0,
                }
            if not transit_remarks_col_candidates:
                return {
                    "detail_html": "⚠️ Column 'Transit Remarks' does not exist.",
                    "count": 0,
                }

            transit_kpi_col = transit_kpi_col_candidates[0]
            transit_group_col = transit_group_col_candidates[0]
            transit_remarks_col = transit_remarks_col_candidates[0]
            transit_container_col = transit_container_candidates[0]

            # تنظيف القيم
            df[transit_kpi_col] = df[transit_kpi_col].astype(str).str.strip()
            df[transit_group_col] = df[transit_group_col].astype(str).str.strip()
            df[transit_remarks_col] = df[transit_remarks_col].astype(str).str.strip()
            df[transit_container_col] = (
                df[transit_container_col].astype(str).str.strip()
            )

            # فلترة الشهر لو موجود
            df_transit = df.copy()
            if selected_month:
                df_transit = df_transit[
                    df_transit["MonthAbbr"].str.lower() == selected_month.lower()
                ]

            months_to_show = (
                df_transit.sort_values(month_col)["MonthAbbr"].unique().tolist()
            )

            # Hit
            df_hit_transit = df_transit[
                df_transit[transit_kpi_col].str.lower() == "hit"
            ].drop_duplicates(subset=[transit_container_col, month_col])
            hit_counts = {
                m: int(
                    df_hit_transit[df_hit_transit["MonthAbbr"] == m][
                        transit_container_col
                    ].nunique()
                )
                for m in months_to_show
            }

            # Miss
            df_miss_transit = df_transit[
                df_transit[transit_kpi_col].str.lower() == "miss"
            ].drop_duplicates(subset=[transit_container_col, month_col])
            if not df_miss_transit.empty:
                pivot_miss = (
                    df_miss_transit.groupby(
                        [transit_remarks_col, transit_group_col, "MonthAbbr"]
                    )[transit_container_col]
                    .nunique()
                    .unstack(fill_value=0)
                    .reset_index()
                )
            else:
                pivot_miss = pd.DataFrame(
                    columns=[transit_remarks_col, transit_group_col] + months_to_show
                )

            rows = [
                {
                    "KPI (Delay Reason)": "On-time Delivery",
                    "Responsible": "Tamer",
                    **{m: hit_counts.get(m, 0) for m in months_to_show},
                }
            ]
            for _, row in pivot_miss.iterrows():
                rows.append(
                    {
                        "KPI (Delay Reason)": row[transit_remarks_col],
                        "Responsible": row[transit_group_col],
                        **{m: int(row.get(m, 0)) for m in months_to_show},
                    }
                )

            total_row = {"KPI (Delay Reason)": "Total", "Responsible": "-"}
            for m in months_to_show:
                total_row[m] = sum(r[m] for r in rows if m in r)
            rows.append(total_row)

            pivot_transit_display = pd.DataFrame(
                rows, columns=["KPI (Delay Reason)", "Responsible"] + months_to_show
            )

            transit_hit_percent = {
                m: int((rows[0][m] / total_row[m]) * 100) if total_row[m] > 0 else 0
                for m in months_to_show
            }

            transit_chart = [
                {
                    "title": f"Transit KPI (Hit %) - {sheet_name}",
                    "type": "column",
                    "name": "Transit Hit %",
                    "color": "#b699d3",
                    "dataPoints": [
                        {"label": m, "y": transit_hit_percent[m]}
                        for m in months_to_show
                    ],
                    "related_table": "Transit KPI Summary",
                },
                {
                    "title": f"Transit KPI (Target) - {sheet_name}",
                    "type": "line",
                    "name": "Target (98%)",
                    "color": "#f44336",
                    "markerSize": 5,
                    "dataPoints": [{"label": m, "y": 98} for m in months_to_show],
                    "related_table": "Transit KPI Summary",
                },
            ]

            transit_table = {
                "id": "seaport-transit-kpi-summary",
                "title": "Transit KPI Summary",  # ✅ استخدام title ثابت بدون f-string
                "columns": pivot_transit_display.columns.tolist(),
                "data": pivot_transit_display.to_dict(orient="records"),
                "chart_data": transit_chart,  # ✅ إضافة chart_data للجدول
            }

            # ✅ Debug: طباعة معلومات transit_table
            print(f"🔍 [DEBUG Seaport] transit_table title: {transit_table['title']}")
            print(f"🔍 [DEBUG Seaport] transit_table id: {transit_table['id']}")
            print(f"🔍 [DEBUG Seaport] transit_chart count: {len(transit_chart)}")
            print(
                f"🔍 [DEBUG Seaport] transit_table chart_data count: {len(transit_table.get('chart_data', []))}"
            )
            if transit_table.get("chart_data"):
                print(
                    f"🔍 [DEBUG Seaport] transit_table chart_data names: {[ds.get('name', 'N/A') for ds in transit_table['chart_data']]}"
                )
                print(
                    f"🔍 [DEBUG Seaport] transit_table chart_data types: {[ds.get('type', 'N/A') for ds in transit_table['chart_data']]}"
                )

            # ===========================================================
            # إخراج النتائج
            # ===========================================================
            tab_data = {
                "name": sheet_name,
                "sub_tables": [hit_table, roch_table, transit_table],
                "chart_data": hit_chart + transit_chart,
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

            return {
                "detail_html": html,
                "count": df_hit.shape[0],
                "hit_pct": hit_percent,  # Hit % لكل شهر الآن
                "chart_data": hit_chart + transit_chart,
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing data: {e}</p>",
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

            outbound_html_result = self.filter_outbound(
                request, selected_month if not selected_months_norm else None
            )
            outbound_html = outbound_html_result.get("detail_html", "")

            if not sub_tables:
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
                "name": "Total Lead Time Performance",
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

            return {
                "detail_html": html,
                "chart_data": chart_data,
                "chart_title": "Total Lead Time Performance – On-Time Delivery",
                "count": total_count,
                "hit_pct": actual_target,  # ✅ أضفنا النسبة هنا
                "tab_data": tab_data,
            }

        except Exception as e:
            import traceback

            print(traceback.format_exc())
            return {
                "detail_html": f"<p class='text-danger'>⚠️ Error while processing data: {e}</p>",
                "chart_data": [],
                "count": 0,
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
                "name": "Dock to stock — Inbound",
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
                    dataset.setdefault("related_table", "Dock to stock")

            # ✅ إضافة chart_data لكل sub_table بشكل منفصل
            chart_data_3pl = []
            chart_data_roche = []
            if chart_data:
                for dataset in chart_data:
                    # ✅ نسخ dataset لكل sub_table مع related_table الصحيح
                    dataset_3pl = dataset.copy()
                    dataset_3pl["related_table"] = "Dock to stock — 3PL"
                    chart_data_3pl.append(dataset_3pl)

                    dataset_roche = dataset.copy()
                    dataset_roche["related_table"] = "Dock to stock — Roche"
                    chart_data_roche.append(dataset_roche)

            tab_data = {
                "name": "Dock to stock",
                "sub_tables": [
                    {
                        "id": "sub-table-dock-to-stock-3pl",  # ✅ إضافة ID فريد
                        "title": "Dock to stock — 3PL",
                        "columns": df_3pl_table.columns.tolist(),
                        "data": df_3pl_table.to_dict(orient="records"),
                        "chart_data": chart_data_3pl,  # ✅ إضافة chart_data لكل sub_table
                    },
                    {
                        "id": "sub-table-dock-to-stock-roche",  # ✅ إضافة ID فريد
                        "title": "Dock to stock — Roche",
                        "columns": df_roche_table.columns.tolist(),
                        "data": df_roche_table.to_dict(orient="records"),
                        "chart_data": chart_data_roche,  # ✅ إضافة chart_data لكل sub_table
                    },
                ],
                "combined_reasons": combined_reasons,
                "canvas_id": f"chart-{slugify('dock-to-stock')}",
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

            print(f"📊 [RESULT] Dock to stock — Hit={hit_pct}%, Target={target_pct}")

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

    def calculate_tab_progress(self, tab_name, selected_month=None):
        print(
            "🟢 [START] حساب نسبة التقدم للتاب:",
            tab_name,
            "— selected_month=",
            selected_month,
        )
        try:
            import pandas as pd, os

            print("────────────────────────────────────────────────────")
            print(
                f"🟢 [START] حساب نسبة التقدم للتاب: '{tab_name}' — selected_month={selected_month}"
            )
            print(f"🟡 [DEBUG] tab_name.lower() = {tab_name.lower()}")
            print("────────────────────────────────────────────────────")

            excel_path = self.get_excel_path()
            print(f"🟢 [DEBUG] استدعاء self.get_excel_path() => {excel_path}")

            if not excel_path or not os.path.exists(excel_path):
                print("⚠️ [DEBUG] ملف Excel غير موجود أو المسار غير صالح.")
                return 0, 100

            xls = pd.ExcelFile(excel_path, engine="openpyxl")
            print(f"🟢 [DEBUG] أسماء شيتات الملف: {xls.sheet_names}")

            # ============== 1) PODs Update ==============
            if tab_name.lower() == "pods update":
                print("🔷 [TAB] PODs Update — بداية المعالجة")
                sheet_name = next(
                    (s for s in xls.sheet_names if "pod" in s.lower()), None
                )
                if not sheet_name:
                    print("⚠️ [DEBUG] لم يتم العثور على شيت PODs.")
                    return 0, 100

                df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
                df.columns = df.columns.str.strip()

                columns_map = {}
                for col in df.columns:
                    name = col.strip().lower()
                    if "month" in name:
                        columns_map["Month"] = col
                    elif "closed" in name:
                        columns_map["Closed"] = col
                    elif "pending" in name:
                        columns_map["Pending"] = col
                if len(columns_map) < 3:
                    print(
                        "⚠️ [DEBUG] الأعمدة المطلوبة (Month, Closed, Pending) غير كافية."
                    )
                    return 0, 100

                df = df.rename(columns=columns_map)
                df = df.dropna(subset=["Month"])
                df = df[~df["Month"].astype(str).str.lower().eq("total")]

                if selected_month:
                    prefix = selected_month[:3]
                    df = df[df["Month"].astype(str).str.startswith(prefix)]

                closed_sum = df["Closed"].apply(pd.to_numeric, errors="coerce").sum()
                pending_sum = df["Pending"].apply(pd.to_numeric, errors="coerce").sum()
                total_sum = closed_sum + pending_sum
                hit_pct = (
                    round((closed_sum / total_sum) * 100, 2) if total_sum != 0 else 0
                )
                print(f"✅ [RESULT] PODs Update -> hit_pct={hit_pct}%")
                return hit_pct, 100

            # ============== 2) Total Lead Time Performance ==============
            elif tab_name.lower() == "total lead time performance":
                print("🔷 [TAB] Total Lead Time Performance — بداية المعالجة")
                sheet_name = next(
                    (
                        s
                        for s in xls.sheet_names
                        if "total lead time preformance" in s.lower()
                        and "-r" not in s.lower()
                    ),
                    None,
                )
                if not sheet_name:
                    print("⚠️ [DEBUG] لا يوجد شيت Total Lead Time Performance.")
                    return 0, 100

                df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
                df.columns = df.columns.str.strip().str.lower()

                if "month" not in df.columns or "kpi" not in df.columns:
                    print("⚠️ [DEBUG] الأعمدة الأساسية غير موجودة.")
                    return 0, 100

                df["month"] = pd.to_datetime(df["month"], errors="coerce").dt.strftime(
                    "%b"
                )
                if selected_month:
                    sel = selected_month[:3]
                    df = df[df["month"] == sel]

                total = len(df)
                hit = len(df[df["kpi"].astype(str).str.lower() != "miss"])
                hit_pct = round((hit / total) * 100, 2) if total > 0 else 0
                print(f"✅ [RESULT] Total Lead Time Performance -> hit_pct={hit_pct}%")
                return hit_pct, 100

            # ============== 3) Order General Information ==============
            elif tab_name.lower() == "order general information":
                print("🔷 [TAB] Order General Information — بداية المعالجة")
                sheet_urgent = next(
                    (
                        s
                        for s in xls.sheet_names
                        if "urgent orders details" in s.lower()
                    ),
                    None,
                )
                sheet_outbound = next(
                    (s for s in xls.sheet_names if "outbound details" in s.lower()),
                    None,
                )
                if not sheet_urgent:
                    print("⚠️ [DEBUG] لم يتم العثور على شيت Urgent orders details.")
                    return 0, 100

                df_urgent = pd.read_excel(
                    excel_path, sheet_name=sheet_urgent, engine="openpyxl"
                )
                df_urgent.columns = df_urgent.columns.str.strip().str.lower()
                total_orders = None
                if sheet_outbound:
                    df_out = pd.read_excel(
                        excel_path, sheet_name=sheet_outbound, engine="openpyxl"
                    )
                    total_orders = len(df_out)

                count_col_name = next(
                    (c for c in df_urgent.columns if "count" in c and "urg" in c), None
                )
                pct_col_name = next(
                    (c for c in df_urgent.columns if "%" in c or "percent" in c), None
                )

                if count_col_name and total_orders and total_orders > 0:
                    urgent_sum = (
                        pd.to_numeric(df_urgent[count_col_name], errors="coerce")
                        .fillna(0)
                        .sum()
                    )
                    non_urgent = max(total_orders - urgent_sum, 0)
                    hit_pct = round((non_urgent / total_orders) * 100, 2)
                    return hit_pct, 100

                if pct_col_name:
                    pct_vals = pd.to_numeric(
                        df_urgent[pct_col_name].astype(str).str.replace("%", ""),
                        errors="coerce",
                    ).dropna()
                    if len(pct_vals) > 0:
                        avg_urgent_pct = pct_vals.mean()
                        hit_pct = round(100 - avg_urgent_pct, 2)
                        return hit_pct, 100

                return 0, 100

            # ============== 4) Rejections ==============
            elif tab_name.lower() == "rejections":
                print("🔷 [TAB] Rejections — بداية المعالجة")
                # نحاول حساب نسبة بسيطة من Booking orders
                sheet_rejection = next(
                    (
                        s
                        for s in xls.sheet_names
                        if "rejection" in s.lower() and "breakdown" not in s.lower()
                    ),
                    None,
                )
                if not sheet_rejection:
                    print("⚠️ [DEBUG] لم يتم العثور على شيت Rejection.")
                    return 0, 100

                df = pd.read_excel(
                    excel_path, sheet_name=sheet_rejection, engine="openpyxl"
                )
                df.columns = df.columns.str.strip()
                if "Booking orders" in df.columns:
                    vals = pd.to_numeric(
                        df["Booking orders"].astype(str).str.replace("%", ""),
                        errors="coerce",
                    ).dropna()
                    if len(vals) > 0:
                        avg_val = vals.mean()
                        hit_pct = round(100 - avg_val, 2)
                        print(f"✅ [RESULT] Rejections -> hit_pct={hit_pct}%")
                        return hit_pct, 100
                print("⚠️ [DEBUG] لم يتم العثور على عمود Booking orders.")
                return 0, 100

            # ============== 5) Data Logger Measurement ==============
            elif tab_name.lower() == "data logger measurement":
                print("🔷 [TAB] Data Logger Measurement — بداية المعالجة")

                try:
                    # 🟢 استدعاء نفس دالة الفلترة الأصلية
                    res = self.filter_data_logger_measurement(None, selected_month)
                    print("🟢 [STEP] تم تنفيذ filter_data_logger_measurement() بنجاح")

                    # 🟢 استخراج نسبة On Time Sent من البيانات اللي راجعة من الشارت
                    if "chart_data" in res and len(res["chart_data"]) > 0:
                        chart_series = res["chart_data"][
                            0
                        ]  # أول سلسلة (On Time Sent %)
                        y_values = [
                            p["y"]
                            for p in chart_series["dataPoints"]
                            if isinstance(p["y"], (int, float)) and p["y"] > 0
                        ]

                        # ✅ النسبة النهائية = آخر شهر فيه بيانات (زي اللي بتظهر في الشارت)
                        hit_pct = y_values[-1] if y_values else 0
                        print(
                            f"✅ [RESULT] نسبة On Time Sent النهائية (آخر شهر) = {hit_pct}%"
                        )
                    else:
                        print("⚠️ [DEBUG] لا توجد بيانات chart_data صالحة — النسبة = 0")
                        hit_pct = 0

                    target_pct = 100
                    return hit_pct, target_pct

                except Exception as e:
                    print(f"❌ [ERROR] أثناء معالجة Data Logger Measurement: {e}")
                    return 0, 100

            # ============== 6) Dock to Stock ==============
            elif tab_name.lower() == "dock to stock":
                print("🔷 [TAB] Dock to stock — بداية المعالجة")

                try:
                    res = self.filter_dock_to_stock_combined(None, selected_month)
                    print("🟢 [DEBUG] نتيجة filter_dock_to_stock_combined:", bool(res))

                    chart_data = res.get("chart_data", [])
                    if chart_data:
                        on_time_series = next(
                            (s for s in chart_data if "on time" in s["name"].lower()),
                            None,
                        )
                        target_series = next(
                            (s for s in chart_data if "target" in s["name"].lower()),
                            None,
                        )

                        if on_time_series and target_series:
                            # نجيب آخر قيمة صالحة
                            on_time_points = [
                                p["y"]
                                for p in on_time_series["dataPoints"]
                                if isinstance(p["y"], (int, float))
                            ]
                            target_points = [
                                p["y"]
                                for p in target_series["dataPoints"]
                                if isinstance(p["y"], (int, float))
                            ]

                            if on_time_points:
                                last_hit = on_time_points[-1]
                            else:
                                last_hit = 0

                            if target_points:
                                last_target = target_points[-1]
                            else:
                                last_target = 100

                            # 🔹 تأكد أن النسبة لا تتعدى 100%
                            hit_pct = min(round(float(last_hit), 2), 100)
                            target_pct = min(round(float(last_target), 2), 100)

                            print(
                                f"📊 [RESULT] Dock to stock (من الشارت): Hit={hit_pct}%, Target={target_pct}%"
                            )
                            return hit_pct, target_pct

                    # ✅ fallback لو مفيش chart_data
                    detail_html = res.get("detail_html", "")
                    if "On time" in detail_html and "%" in detail_html:
                        import re

                        matches = re.findall(r"(\d+(?:\.\d+)?)\s*%", detail_html)
                        if matches:
                            hit_pct = float(matches[-1])
                            hit_pct = min(hit_pct, 100)
                            print(
                                f"📊 [RESULT] Dock to stock (من HTML): Hit={hit_pct}%"
                            )
                            return hit_pct, 100

                    print("⚠️ [INFO] لا توجد بيانات Dock to stock صالحة.")
                    return 0, 100

                except Exception as e:
                    print(f"⚠️ [ERROR] أثناء حساب progress للتاب Dock to stock: {e}")
                    return 0, 100

            # ============== 7)Airport Clearance ==============
            elif "airport clearance" in tab_name.lower():
                print("🔷 [TAB] Airport Clearance — بداية المعالجة")
                result = self.filter_airport_combined(None, selected_month)
                if result and "chart_data" in result:
                    on_time_series = next(
                        (
                            s
                            for s in result["chart_data"]
                            if s["name"].lower().startswith("on time")
                        ),
                        None,
                    )
                    if on_time_series and on_time_series["dataPoints"]:
                        hit_values = [
                            p["y"]
                            for p in on_time_series["dataPoints"]
                            if isinstance(p["y"], (int, float))
                        ]
                        hit_avg = (
                            round(sum(hit_values) / len(hit_values), 2)
                            if hit_values
                            else 0
                        )
                        print(f"✅ [RESULT] Airport Clearance -> hit_pct={hit_avg}%")
                        return hit_avg, 100
                print("⚠️ [DEBUG] لا توجد بيانات صالحة لـ Airport Clearance.")
                return 0, 100

            # ============== 8) Seaport Clearance ==============
            elif "seaport clearance" in tab_name.lower():
                print("🔷 [TAB] Seaport Clearance — بداية المعالجة")
                try:
                    result = self.filter_seaport_combined(None, selected_month)
                    print("🟢 [DEBUG] نتيجة filter_seaport_combined:", bool(result))

                    # 🔹 التأكد من وجود بيانات الشارت
                    if result and "chart_data" in result:
                        on_time_series = next(
                            (
                                s
                                for s in result["chart_data"]
                                if "on time receiving" in s["name"].lower()
                            ),
                            None,
                        )
                        if on_time_series and on_time_series.get("dataPoints"):
                            # استخراج القيم من الأعمدة
                            values = [
                                p["y"]
                                for p in on_time_series["dataPoints"]
                                if isinstance(p["y"], (int, float))
                            ]
                            if values:
                                avg_hit = round(sum(values) / len(values), 2)
                                print(
                                    f"✅ [RESULT] Seaport Clearance -> hit_pct={avg_hit}%"
                                )
                                return avg_hit, 100
                            else:
                                print(
                                    "⚠️ [DEBUG] لم يتم العثور على قيم عددية في On time receiving"
                                )
                                return 0, 100
                    print("⚠️ [DEBUG] لا توجد بيانات شارت صالحة لـ Seaport Clearance.")
                    return 0, 100

                except Exception as e:
                    print(f"❌ [ERROR] أثناء معالجة Seaport Clearance: {e}")
                    return 0, 100

            # باقي التابات (افتراضي)
            else:
                print(
                    "ℹ️ [INFO] لم يتم تعريف حساب خاص لهذا التاب. سيتم إرجاع النسبة 0%."
                )
                return 0, 100

        except Exception:
            import traceback

            print("❌ [EXCEPTION] حدث خطأ غير متوقع في calculate_tab_progress():")
            traceback.print_exc()
            return 0, 100

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
            "total lead time performance": 98,
            "data logger measurement": 99,
            "dock to stock": 99,
            "airport clearance": 98,
            "seaport clearance": 98,
            "pods update": 98,
            "return & refusal": 100,
        }

        def process_tab(tab_name):
            detail_html, count, hit_pct = "", 0, 0
            try:
                res = {}
                tab_lower = tab_name.lower()
                month_for_filters = selected_month if not selected_months else None

                if tab_lower in ["rejections", "return & refusal"]:
                    res = self.filter_rejections_combined(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif tab_lower == "data logger measurement":
                    res = self.filter_data_logger_measurement(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif tab_lower == "dock to stock":
                    res = self.filter_dock_to_stock_combined(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )
                elif "airport clearance" in tab_lower:
                    res = self.filter_airport_combined(request, month_for_filters)
                elif "seaport clearance" in tab_lower:
                    res = self.filter_seaport_combined(request, month_for_filters)
                elif "pods update" in tab_lower:
                    res = self.filter_pods_update(request, month_for_filters)
                elif "total lead time performance" in tab_lower:
                    res = self.filter_total_lead_time_performance(
                        request,
                        month_for_filters,
                        selected_months=selected_months,
                    )

                # النسبة الحقيقية زي ما راجعة من الدالة
                hit_pct = res.get("hit_pct", 0)
                if isinstance(hit_pct, dict):
                    if selected_month and selected_month.capitalize() in hit_pct:
                        hit_pct_val = hit_pct[selected_month.capitalize()]
                    else:
                        # نحسب المتوسط
                        hit_pct_val = int(round(sum(hit_pct.values()) / len(hit_pct)))
                else:
                    try:
                        hit_pct_val = int(round(float(hit_pct)))
                    except:
                        hit_pct_val = 0

                hit_pct_val = max(0, min(hit_pct_val, 100))

                if tab_lower == "dock to stock":
                    target_pct = hit_pct_val
                else:
                    target_pct = target_manual.get(tab_lower, 100)
                color_class = "bg-success" if hit_pct >= target_pct else "bg-danger"

                progress_html = f"""
                    <div class='mb-3'>
                        <div class='d-flex justify-content-between align-items-center mb-1'>
                            <strong class='text-capitalize'>{tab_name}</strong>
                            <small>{hit_pct}% / Target: {target_pct}%</small>
                        </div>
                        <div class='progress' style='height: 20px;'>
                            <div class='progress-bar {color_class}' role='progressbar'
                                 style='width: {hit_pct}%;' aria-valuenow='{hit_pct}'
                                 aria-valuemin='0' aria-valuemax='100'>
                                 {hit_pct}%
                            </div>
                        </div>
                    </div>
                """

                detail_html = progress_html + (res.get("detail_html", "") or "")
                count = res.get("count", 0)

            except Exception:
                detail_html = "<p class='text-muted'>No data available.</p>"
                hit_pct = 0
                if tab_name.lower() == "dock to stock":
                    target_pct = hit_pct
                else:
                    target_pct = target_manual.get(tab_name.lower(), 100)

            return {
                "name": tab_name,
                "hit_pct": hit_pct_val,
                "target_pct": target_pct,
                "detail_html": detail_html,
                "count": count,
            }

        tabs_order = [
            "Dock to stock",
            "Data logger Measurement",
            "Total Lead Time Performance",
            "PODs update",
            "Return & Refusal",
            "Airport Clearance",
            "Seaport Clearance",
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

    def dashboard_tab(self, request):
        """
        🔹 تاب Dashboard مخصص للتصميم اليدوي (بدون قراءة شيت مباشر)
        """
        try:
            html = render_to_string(
                "dashboard_custom.html",
                {"title": self.DASHBOARD_TAB_NAME},
                request=request,
            )
            return {"detail_html": html}
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

    def cross_docking_tab(self, request):
        """
        🔹 عرض تاب Cross Docking يحتوي على تابين (رسمتين لجدة والرياض)
        """
        import traceback
        from django.template.loader import render_to_string
        from django.http import JsonResponse

        try:
            # ✅ تجهيز السياق لكل رسمه
            context_jeddah = {
                "title": "Jeddah Cross-Docking Performance",
            }
            context_riyadh = {
                "title": "Riyadh Cross-Docking Performance",
            }

            # ✅ تحميل التمبلتين (كل واحدة فيها رسمه Chart)
            jeddah_html = render_to_string(
                "includes_cross_docking/jeddah_crossdock.html",
                context_jeddah,
                request=request,
            )
            riyadh_html = render_to_string(
                "includes_cross_docking/riyadh_crossdock.html",
                context_riyadh,
                request=request,
            )

            # ✅ بناء التابات
            html = render_to_string(
                "cross_docking.html",
                {
                    "jeddah_html": jeddah_html,
                    "riyadh_html": riyadh_html,
                },
                request=request,
            )

            return JsonResponse({"detail_html": html}, safe=False)

        except Exception as e:
            traceback.print_exc()
            return JsonResponse(
                {"error": f"An error occurred while loading tabs: {e}"}, status=500
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
