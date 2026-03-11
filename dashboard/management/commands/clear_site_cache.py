# -*- coding: utf-8 -*-
"""
مسح كل كاش الموقع: Django cache، DashboardDataCache، ExcelSheetCache، وملف dashboard_cache.json
استخدام: python manage.py clear_site_cache
"""
import os
import json
from django.core.management.base import BaseCommand
from django.core.cache import cache
from django.conf import settings


class Command(BaseCommand):
    help = "مسح كل كاش الموقع (Django cache، Dashboard، Excel sheets، ملف JSON)"

    def add_arguments(self, parser):
        parser.add_argument(
            "--no-input",
            action="store_true",
            help="تشغيل بدون تأكيد",
        )

    def handle(self, *args, **options):
        if not options["no_input"]:
            self.stdout.write("سيتم مسح: Django cache، DashboardDataCache، ExcelSheetCache، وملف dashboard_cache.json")
        try:
            # 1) Django cache (file-based أو غيره)
            cache.clear()
            self.stdout.write(self.style.SUCCESS("✅ تم مسح Django cache"))

            # 2) DashboardDataCache من الداتابيز
            from dashboard.models import DashboardDataCache

            n = DashboardDataCache.objects.count()
            DashboardDataCache.objects.all().delete()
            self.stdout.write(self.style.SUCCESS(f"✅ تم حذف {n} سجل من DashboardDataCache"))

            # 3) ExcelSheetCache من الداتابيز
            from dashboard.models import ExcelSheetCache

            n2 = ExcelSheetCache.objects.count()
            ExcelSheetCache.objects.all().delete()
            self.stdout.write(self.style.SUCCESS(f"✅ تم حذف {n2} سجل من ExcelSheetCache"))

            # 4) ملف dashboard_cache.json
            media_root = getattr(settings, "MEDIA_ROOT", "")
            json_path = os.path.join(media_root, "excel_uploads", "dashboard_cache.json")
            if os.path.isfile(json_path):
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump({}, f)
                self.stdout.write(self.style.SUCCESS("✅ تم تفريغ ملف dashboard_cache.json"))
            else:
                self.stdout.write("ملف dashboard_cache.json غير موجود (لا مشكلة)")

            # 5) مجلد file-based cache إن وُجد
            cache_loc = getattr(settings, "CACHES", {}).get("default", {}).get("LOCATION")
            if cache_loc and os.path.isdir(cache_loc):
                for name in os.listdir(cache_loc):
                    p = os.path.join(cache_loc, name)
                    try:
                        if os.path.isfile(p):
                            os.remove(p)
                    except Exception:
                        pass
                self.stdout.write(self.style.SUCCESS("✅ تم حذف محتويات مجلد Django file cache"))

            self.stdout.write(self.style.SUCCESS("\nتم مسح كل الكاش. حدّث الصفحة أو أعد رفع الملف لرؤية التحديثات."))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"خطأ: {e}"))
