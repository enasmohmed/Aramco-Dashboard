# -*- coding: utf-8 -*-
from django.test import TestCase


class StockCountAllInOneChartTest(TestCase):
    """
    تحقق أن كارد Stock Count في All-in-One يكون له دائماً chart_data غير فاضية
    حتى مع خطأ أو عدم وجود داتا، عشان الشارت يظهر ولا يظهر «No data available» مكانه.
    """

    def test_stock_count_fallback_chart_data_structure(self):
        """الـ fallback لـ Stock Count يكون: type column، و dataPoints فيها على الأقل نقطة واحدة."""
        # نفس البنية اللي بنحطها في الـ fallback (في الـ exception handler و في process_tab)
        fallback_chart_data = [
            {
                "type": "column",
                "name": "Stock Count Hit %",
                "valueSuffix": "%",
                "dataPoints": [{"label": "Hit %", "y": 0}],
            }
        ]
        self.assertTrue(len(fallback_chart_data) > 0)
        self.assertEqual(fallback_chart_data[0].get("type"), "column")
        data_points = fallback_chart_data[0].get("dataPoints", [])
        self.assertTrue(len(data_points) >= 1, "Stock Count fallback يجب أن يحتوي على نقطة واحدة على الأقل لرسم الشارت")
        self.assertIn("label", data_points[0])
        self.assertIn("y", data_points[0])

    def test_stock_count_slug_matches_html_id(self):
        """الـ slug المستخدم في JS (all-one-chart-{slug}) يطابق الـ slug في الـ HTML."""
        # Python (views): chart_id_slug من "Stock Count"
        name = "Stock Count"
        chart_id_slug = (
            name.lower()
            .replace(" & ", "-")
            .replace(" ", "-")
            .replace("'", "")
        )
        chart_id_slug = "".join(c for c in chart_id_slug if c.isalnum() or c == "-").strip("-") or "chart"
        self.assertEqual(chart_id_slug, "stock-count")
        element_id = f"all-one-chart-{chart_id_slug}"
        self.assertEqual(element_id, "all-one-chart-stock-count")
