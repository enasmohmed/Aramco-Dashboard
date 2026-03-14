import os
import shutil
import tempfile
from django.conf import settings
from django.contrib import admin
from django.shortcuts import render, redirect
from django.urls import path
from django import forms
from django.contrib import messages
from .models import MeetingPoint, InboundShipmentRemark, OutboundOrderRemark, TransportationKPI
from .transportation_import import parse_transportation_excel


@admin.register(InboundShipmentRemark)
class InboundShipmentRemarkAdmin(admin.ModelAdmin):
    list_display = ("shipment_nbr", "facility", "remark_short", "status_override", "updated_at")
    list_editable = ()
    list_filter = ("facility",)
    search_fields = ("shipment_nbr", "facility", "remark")
    ordering = ("-updated_at",)

    def remark_short(self, obj):
        return (obj.remark[:50] + "…") if obj.remark and len(obj.remark) > 50 else (obj.remark or "")

    remark_short.short_description = "Remark"


@admin.register(OutboundOrderRemark)
class OutboundOrderRemarkAdmin(admin.ModelAdmin):
    list_display = ("order_nbr", "facility", "remark_short", "status_override", "updated_at")
    list_editable = ()
    list_filter = ("facility",)
    search_fields = ("order_nbr", "facility", "remark")
    ordering = ("-updated_at",)

    def remark_short(self, obj):
        return (obj.remark[:50] + "…") if obj.remark and len(obj.remark) > 50 else (obj.remark or "")

    remark_short.short_description = "Remark"


@admin.register(MeetingPoint)
class MeetingPointAdmin(admin.ModelAdmin):
    list_display = ("description", "is_done", "created_at", "target_date")
    list_editable = ("is_done", "target_date",)
    list_filter = ("is_done", "created_at", "target_date")
    search_fields = ("description",)
    ordering = ("-created_at", "target_date", "assigned_to")

    # ✅ السماح بتعديل created_at من صفحة التفاصيل
    fields = ("description", "is_done", "created_at", "target_date", "assigned_to")


class ImportTransportationExcelForm(forms.Form):
    excel_file = forms.FileField(
        label="ملف الإكسل",
        help_text="اختر ملف إكسل يحتوي على شيت Transportation (Inbound / Outbound، January / February، الأعمدة: REGION, TOTAL, HIT, MISS, ACHIEVED)",
    )


@admin.register(TransportationKPI)
class TransportationKPIAdmin(admin.ModelAdmin):
    list_display = ("section", "month", "kpi", "region", "total", "sum_value", "hit", "miss", "total_submitted", "achieved_percent", "target_percent")
    list_editable = ("total", "sum_value", "hit", "miss", "total_submitted", "achieved_percent", "target_percent")
    list_filter = ("section", "month", "kpi", "region")
    search_fields = ("section", "month", "kpi", "region")
    ordering = ("section", "month", "kpi", "region")
    list_per_page = 50
    actions = ["create_missing_rows"]
    change_list_template = "admin/dashboard/transportationkpi/change_list.html"

    def get_urls(self):
        urls = super().get_urls()
        return [
            path("import-excel/", self.admin_site.admin_view(self.import_excel_view), name="dashboard_transportationkpi_import_excel"),
        ] + urls

    def import_excel_view(self, request):
        if request.method == "POST":
            form = ImportTransportationExcelForm(request.POST, request.FILES)
            if form.is_valid():
                f = request.FILES.get("excel_file")
                if not f or not f.name.lower().endswith((".xlsx", ".xls")):
                    messages.error(request, "يرجى رفع ملف إكسل (.xlsx أو .xls).")
                    return render(request, "admin/dashboard/transportationkpi/import_excel.html", {"form": form, "opts": self.model._meta})
                path = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        for chunk in f.chunks():
                            tmp.write(chunk)
                        path = tmp.name
                    rows, errors = parse_transportation_excel(path)
                    if errors:
                        for e in errors:
                            messages.error(request, e)
                        return render(request, "admin/dashboard/transportationkpi/import_excel.html", {"form": form, "opts": self.model._meta})
                    updated = 0
                    created = 0
                    for r in rows:
                        obj, created_flag = TransportationKPI.objects.update_or_create(
                            section=r["section"],
                            month=r["month"],
                            kpi=r["kpi"],
                            region=r["region"],
                            defaults={
                                "total": r.get("total", ""),
                                "sum_value": r.get("sum_value", ""),
                                "hit": r.get("hit", ""),
                                "miss": r.get("miss", ""),
                                "total_submitted": r.get("total_submitted", ""),
                                "achieved_percent": r.get("achieved_percent", ""),
                                "target_percent": r.get("target_percent", "98"),
                            },
                        )
                        if created_flag:
                            created += 1
                        else:
                            updated += 1
                    # حفظ نسخة من الملف في excel_uploads ليعرض تاب Transportation (From Factory to WH / Normal delivery) من الموقع
                    dest_dir = os.path.join(settings.MEDIA_ROOT, "excel_uploads")
                    os.makedirs(dest_dir, exist_ok=True)
                    dest_path = os.path.join(dest_dir, "transportation_from_admin.xlsx")
                    try:
                        shutil.copy2(path, dest_path)
                        messages.success(
                            request,
                            f"تم استيراد البيانات: {created} سجل جديد، {updated} سجل محدّث. تم حفظ الملف لعرضه في تاب Transportation من الموقع.",
                        )
                    except Exception as copy_err:
                        messages.success(request, f"تم استيراد البيانات: {created} سجل جديد، {updated} سجل محدّث. (تحذير: لم يُحفظ الملف للموقع: {copy_err})")
                    return redirect("admin:dashboard_transportationkpi_changelist")
                finally:
                    if path and os.path.exists(path):
                        try:
                            os.unlink(path)
                        except Exception:
                            pass
            else:
                for _field, err in form.errors.items():
                    messages.error(request, err)
        else:
            form = ImportTransportationExcelForm()
        return render(request, "admin/dashboard/transportationkpi/import_excel.html", {"form": form, "opts": self.model._meta})

    @admin.action(description="إنشاء الصفوف الناقصة لهذا القسم والشهر")
    def create_missing_rows(self, request, queryset):
        created = 0
        for section in ["Inbound", "Outbound"]:
            for month in ["January", "February"]:
                for kpi in ["Delivery Fulfilment", "On Time Delivery", "PODs submission"]:
                    for region in ["Fuchs-Yanbu", "Fuchs-Jeddah"]:
                        _, is_new = TransportationKPI.objects.get_or_create(
                            section=section,
                            month=month,
                            kpi=kpi,
                            region=region,
                            defaults={"target_percent": "98"},
                        )
                        if is_new:
                            created += 1
        self.message_user(request, f"تم إنشاء {created} صفًا جديدًا.")
