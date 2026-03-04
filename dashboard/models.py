
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
    created_at = models.DateField(default=timezone.now)
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
    """ملاحظات/أسباب تُضاف من الأدمن لشحنات Inbound وتظهر في جدول Inbound Shipments Detail"""
    shipment_nbr = models.CharField(max_length=255, db_index=True)
    facility = models.CharField(max_length=255, db_index=True)
    remark = models.TextField(blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Inbound Shipment Remark"
        verbose_name_plural = "Inbound Shipment Remarks"
        unique_together = [["shipment_nbr", "facility"]]

    def __str__(self):
        return f"{self.shipment_nbr} @ {self.facility}"
