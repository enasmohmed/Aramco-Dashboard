from django.contrib import admin
from .models import MeetingPoint, InboundShipmentRemark, OutboundOrderRemark


@admin.register(InboundShipmentRemark)
class InboundShipmentRemarkAdmin(admin.ModelAdmin):
    list_display = ("shipment_nbr", "facility", "remark_short", "updated_at")
    list_editable = ()
    list_filter = ("facility",)
    search_fields = ("shipment_nbr", "facility", "remark")
    ordering = ("-updated_at",)

    def remark_short(self, obj):
        return (obj.remark[:50] + "…") if obj.remark and len(obj.remark) > 50 else (obj.remark or "")

    remark_short.short_description = "Remark"


@admin.register(OutboundOrderRemark)
class OutboundOrderRemarkAdmin(admin.ModelAdmin):
    list_display = ("order_nbr", "facility", "remark_short", "updated_at")
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
