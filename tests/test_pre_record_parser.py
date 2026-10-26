"""预录单数量与价格关系的离线回归测试。"""
from __future__ import annotations

import unittest
from decimal import Decimal

from app.services.parser.models import DeclarationItem
from app.services.parser.pre_record_parser import (
    correct_item_quantities,
    quantity_candidates,
)


class PreRecordParserTests(unittest.TestCase):
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
