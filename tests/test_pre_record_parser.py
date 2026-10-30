"""预录单数量与价格关系的离线回归测试。"""
from __future__ import annotations

import unittest
from decimal import Decimal

from app.services.parser.declaration_parser import (
    _declaration_number,
    _pre_record_dates,
)
from app.services.parser.models import DeclarationItem
from app.services.parser.pre_record_parser import (
    correct_item_quantities,
    quantity_candidates,
)


class PreRecordParserTests(unittest.TestCase):
    def test_left_column_customs_number_is_read_when_value_is_not_adjacent(self):
        lines = [
            "预录入编号：E20260001865115627",
            "境内发货人(913304017964830942)",
            "海关编号",
            "境外收货人",
            "223120260004677495",
            "合同协议号",
            "26MT-10B365Y",
        ]
        self.assertEqual(
            _declaration_number(lines, "\n".join(lines)),
            "223120260004677495",
        )

    def test_customs_number_prefers_forward_value_over_unrelated_number_above(self):
        lines = [
            "913304017964830942",
            "海关编号",
            *(["其他左侧字段"] * 10),
            "223320260002002667",
        ]
        self.assertEqual(
            _declaration_number(lines, "\n".join(lines)),
            "223320260002002667",
        )

    def test_split_declare_and_export_labels_read_separate_dates_above(self):
        lines = [
            "*310120260514502678*",
            "备案号",
            "20261021",
            "申报日期",
            "20261026",
            "出口日期",
            "梅山港区",
        ]
        self.assertEqual(
            _pre_record_dates(lines),
            ("2026-10-21", "2026-10-26"),
        )

    def test_combined_date_label_keeps_legacy_behavior(self):
        lines = [
            "*310120260514502678*",
            "2026102120261026",
            "申报日期出口日期",
        ]
        self.assertEqual(
            _pre_record_dates(lines),
            ("2026-10-21", "2026-10-26"),
        )

    def test_price_relationship_corrects_reversed_text_quantities(self):
        text = """
        47.5千克
        47.5千克
        2.325千克
        2.325千克
        1 7318159001不锈钢螺栓
        18.0645
        42.00
        2 7505220090镍合金丝
        73.8000
        3505.50
        """
        items = [
            DeclarationItem(
                serial=1,
                product_name="不锈钢螺栓",
                unit_price="18.0645",
                total_price="42.00",
                quantity="47.5千克",
                unit="千克",
            ),
            DeclarationItem(
                serial=2,
                product_name="镍合金丝",
                unit_price="73.8000",
                total_price="3505.50",
                quantity="2.325千克",
                unit="千克",
            ),
        ]
        corrected, warnings = correct_item_quantities(text, items)
        self.assertEqual(warnings, [])
        self.assertEqual(corrected[0].quantity, "2.325千克")
        self.assertEqual(corrected[1].quantity, "47.5千克")

    def test_candidates_stop_at_first_item_and_deduplicate_pairs(self):
        text = """
        12.5千克
        12.5千克
        (1)
        1 7307210000不锈钢法兰
        8.5
        106.25
        """
        self.assertEqual(quantity_candidates(text), [(Decimal("12.5"), "千克")])


if __name__ == "__main__":
    unittest.main()
