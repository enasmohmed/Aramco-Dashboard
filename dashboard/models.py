
from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone


# Create your models here.


class UploadedFile(models.Model):
    file = models.FileField(upload_to='uploads/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.file.name} ({self.uploaded_at:%Y-%m-%d %H:%M})"




class UploadMonth(models.Model):
    month = models.CharField(max_length=20, unique=True)

    def __str__(self):
        return self.month



class MeetingPoint(models.Model):
    description = models.TextField()  # لازم يكون TextField أو CharField
    is_done = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    target_date = models.DateField(null=True, blank=True)
    assigned_to = models.CharField(max_length=255, blank=True, null=True)

    # def save(self, *args, **kwargs):
    #     # لو مفيش تاريخ هدف، حطيه بعد 7 أيام من الإنشاء
    #     if not self.target_date and not self.pk:
    #         from datetime import date
    #         self.target_date = date.today() + timedelta(days=7)
    #     super().save(*args, **kwargs)

    def __str__(self):
        return self.description[:50]


class InboundShipmentRemark(models.Model):
    """ملاحظات/أسباب وتعديل الحالة (Hit/Miss) لشحنات Inbound و Return — تُحرَّر من الجدول"""
    shipment_nbr = models.CharField(max_length=255, db_index=True)
    facility = models.CharField(
        max_length=255,
        db_index=True,
        help_text="المنطقة كما في الداشبورد؛ أو * / ALL لتطبيق التعديل على نفس رقم الشحنة في أي منطقة.",
    )
    remark = models.TextField(blank=True)
    status_override = models.CharField(max_length=10, blank=True, null=True)  # "Hit" or "Miss"; null = use computed
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Inbound Shipment Remark"
        verbose_name_plural = "Inbound Shipment Remarks"
        unique_together = [["shipment_nbr", "facility"]]

    def __str__(self):
        return f"{self.shipment_nbr} @ {self.facility}"


class OutboundOrderRemark(models.Model):
    """ملاحظات/أسباب وتعديل الحالة (Hit/Miss) لأوامر Outbound — تُحرَّر من الجدول"""
    order_nbr = models.CharField(max_length=255, db_index=True)
    facility = models.CharField(
        max_length=255,
        db_index=True,
        help_text="المنطقة كما في الداشبورد؛ أو * / ALL لتطبيق التعديل على نفس رقم الطلب في أي منطقة.",
    )
    remark = models.TextField(blank=True)
    status_override = models.CharField(max_length=10, blank=True, null=True)  # "Hit" or "Miss"; null = use computed
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Outbound Order Remark"
        verbose_name_plural = "Outbound Order Remarks"
        unique_together = [["order_nbr", "facility"]]

    def __str__(self):
        return f"{self.order_nbr} @ {self.facility}"


class ExcelSheetCache(models.Model):
    """كاش بيانات شيت إكسل: يُملأ عند الرفع ويُستخدم لتسريع فتح التابات."""
    sheet_name = models.CharField(max_length=255, unique=True, db_index=True)
    data = models.JSONField(default=list)  # list of dicts (rows)
    source_file_path = models.CharField(max_length=512, blank=True, null=True, db_index=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Excel Sheet Cache"
        verbose_name_plural = "Excel Sheet Caches"

    def __str__(self):
        return f"{self.sheet_name} ({len(self.data)} rows)"


class DashboardDataCache(models.Model):
    """كاش بيانات الداشبورد المُستخرجة من الإكسل: يُملأ عند رفع الملف ويُقرأ من الداتابيز لفتح الداشبورد بسرعة."""
    source_file_path = models.CharField(max_length=512, unique=True, db_index=True)
    data = models.JSONField(default=dict)  # inbound_kpi, pending_shipments, charts, outbound, pods, returns, inventory, warehouse, returns_region
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Dashboard Data Cache"
        verbose_name_plural = "Dashboard Data Caches"

    def __str__(self):
        return f"Dashboard cache ({self.source_file_path})"


class TransportationKPI(models.Model):
    """بيانات Transportation (Inbound/Outbound) القابلة للتعديل من الأدمن — تظهر في الجدول والشارت."""
    SECTION_CHOICES = [
        ("Inbound", "Inbound"),
        ("Outbound", "Outbound"),
    ]
    MONTH_CHOICES = [
        ("January", "January"),
        ("February", "February"),
    ]
    KPI_CHOICES = [
        ("Delivery Fulfilment", "Delivery Fulfilment"),
        ("On Time Delivery", "On Time Delivery"),
        ("PODs submission", "PODs submission"),
    ]
    REGION_CHOICES = [
        ("Fuchs-Yanbu", "Fuchs-Yanbu"),
        ("Fuchs-Jeddah", "Fuchs-Jeddah"),
    ]
    section = models.CharField(max_length=20, choices=SECTION_CHOICES, db_index=True)
    month = models.CharField(max_length=20, choices=MONTH_CHOICES, db_index=True)
    kpi = models.CharField(max_length=50, choices=KPI_CHOICES, db_index=True)
    region = models.CharField(max_length=30, choices=REGION_CHOICES, db_index=True)
    total = models.CharField(max_length=30, blank=True, default="")
    sum_value = models.CharField(max_length=30, blank=True, default="", help_text="SUM (يُعرض لـ On Time Delivery فقط)")
    hit = models.CharField(max_length=30, blank=True, default="")
    miss = models.CharField(max_length=30, blank=True, default="")
    total_submitted = models.CharField(max_length=30, blank=True, default="", help_text="TOTAL SUBMITTED (يُستخدم مع KPI = PODs submission)")
    achieved_percent = models.CharField(max_length=20, blank=True, default="", help_text="مثال: 90 أو 100 (نسبة ACHIEVED)")
    target_percent = models.CharField(max_length=20, blank=True, default="98", help_text="مثال: 98")

    class Meta:
        verbose_name = "Transportation KPI"
        verbose_name_plural = "Transportation KPIs"
        unique_together = [["section", "month", "kpi", "region"]]
        ordering = ["section", "month", "kpi", "region"]

    def __str__(self):
        return f"{self.section} / {self.month} / {self.kpi} / {self.region}"
