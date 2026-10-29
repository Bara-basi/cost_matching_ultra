"""报关 PDF 到飞书写入计划的离线测试。"""
from __future__ import annotations

import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from types import SimpleNamespace

from app.services.customs_ingest import (
    ORIGINAL_AI_COST_FIELD,
    PROTECTED_PURCHASE_FIELD,
    _allocated_fee,
    _match_desired_records,
    build_records,
    candidate_declarations,
    date_millis,
    ensure_ai_result_fields,
    fee_number,
    original_ai_fields,
    write_backfill_to_original,
)


class CustomsIngestTests(unittest.TestCase):
    def test_candidate_requires_multiple_all_normal_rows(self):
        rows = [
            {"报关单号": "A", "异常类型": "正常", "采购金额": "10", "供应商简称": "甲", "采购订单号": "P1"},
            {"报关单号": "A", "异常类型": "正常", "采购金额": "20", "供应商简称": "乙", "采购订单号": "P2"},
            {"报关单号": "B", "异常类型": "正常", "采购金额": "30", "供应商简称": "甲", "采购订单号": "P3"},
            {"报关单号": "C", "异常类型": "正常", "采购金额": "10", "供应商简称": "甲", "采购订单号": "P4"},
            {"报关单号": "C", "异常类型": "异常", "采购金额": "20", "供应商简称": "乙", "采购订单号": "P5"},
        ]
        self.assertEqual(candidate_declarations(rows, 3), ["A"])

    def test_fee_and_date_conversion(self):
        self.assertEqual(fee_number("USD/2800/总价"), Decimal("2800"))
        self.assertEqual(fee_number("/0/"), Decimal(0))
        self.assertEqual(date_millis(datetime(2026, 10, 19)), 1792339200000)
        with self.assertRaises(ValueError):
            date_millis("")

    def test_fee_allocation_absorbs_last_cent(self):
        rows = [{"报关金额": "1"}, {"报关金额": "2"}]
        result = _allocated_fee(rows, {"运费": "USD/10/总价"}, "运费")
        self.assertEqual(sum(result), Decimal(10))
        self.assertEqual(result[0], Decimal("3.33"))
        self.assertEqual(result[1], Decimal("6.67"))

    def test_original_ai_fields_protects_original_purchase_amount(self):
        row = {
            "报关单号": "D1",
            "合同号_1": "26MT-01A100",
            "销售方式": "出口",
            "报关品名": "不锈钢管",
            "报关金额": "123.45",
            "报关重量": "20",
            PROTECTED_PURCHASE_FIELD: "123.45",
            "拆单状态": "正常",
            "成本匹配状态": "正常",
        }
        fields = original_ai_fields(row)
        self.assertNotIn(PROTECTED_PURCHASE_FIELD, fields)
        self.assertEqual(fields[ORIGINAL_AI_COST_FIELD], 123.45)
        self.assertEqual(fields["报关单号"], "D1")
        self.assertEqual(fields["合同号_1"], "26MT-01A100")
        self.assertEqual(fields["报关品名"], "不锈钢管")

    def test_business_key_match_uses_supplier_product_and_amount(self):
        desired = [
            {
                "报关单号": "D1",
                "供应商": "供应商甲",
                "产品类型": "法兰",
                "报关金额": "100",
                "报关重量": "10",
            },
            {
                "报关单号": "D1",
                "供应商": "供应商甲",
                "产品类型": "法兰",
                "报关金额": "200",
                "报关重量": "20",
            },
        ]
        records = [
            {
                "record_id": "R2",
                "fields": {
                    "供应商": [{"text": "供应商甲"}],
                    "产品类型(对内)": [{"text": "法兰"}],
                    "报关金额AI": 200,
                    "报关重量AI": 20,
                },
            },
            {
                "record_id": "R1",
                "fields": {
                    "供应商": [{"text": "供应商甲"}],
                    "产品类型(对内)": [{"text": "法兰"}],
                    "报关金额AI": 100,
                    "报关重量AI": 10,
                },
            },
        ]
        matched = _match_desired_records(desired, records)
        self.assertEqual([record["record_id"] for record in matched], ["R1", "R2"])

    def test_all_result_fields_get_ai_columns(self):
        class FakeClient:
            def __init__(self):
                self.fields = [
                    {"field_name": PROTECTED_PURCHASE_FIELD, "field_id": "fld_purchase", "type": 19},
                ]

            def list_fields(self, _app, _table):
                return list(self.fields)

            def create_field(self, _app, _table, name, field_type):
                created = {
                    "field_name": name,
                    "field_id": f"fld_{len(self.fields)}",
                    "type": field_type,
                }
                self.fields.append(created)
                return created

        client = FakeClient()
        fields = ensure_ai_result_fields(client, "app", "table")
        self.assertEqual(fields[ORIGINAL_AI_COST_FIELD]["type"], 2)
        self.assertEqual(fields["报关重量AI"]["type"], 2)
        self.assertEqual(fields["合同号AI"]["type"], 1)
        self.assertIn("采购金额分摊依据AI", fields)
        self.assertEqual(len(fields), 5)

    def test_single_split_row_is_valid_for_server_task(self):
        parsed = [
            {
                "报关单号": "D1",
                "来源文件": "a.pdf",
                "出口日期": "2026-10-01",
                "海关编码": "7304419000",
                "成交方式": "FOB",
                "币种": "USD",
                "报关重量单位": "千克",
                "目的国": "越南",
                "运费": "/0/",
                "保费": "/0/",
            }
        ]
        result = [
            {
                "报关单号": "D1",
                "合同号_1": "26MT-01A",
                "报关品名": "不锈钢管",
                "产品类型": "无缝管",
                "供应商简称": "供应商",
                "供应商名称": "供应商有限公司",
                "采购订单号": "PO1",
                "报关金额": "100",
                "报关重量": "20",
                "采购金额": "80",
                "分摊依据": "单条记录",
                "异常类型": "正常",
                "异常明细": "",
            }
        ]
        with TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "a.pdf").write_bytes(b"%PDF-1.7\nfixture")
            with patch("app.services.customs_ingest.read_sheet", return_value=parsed):
                source, ai, pdfs = build_records(
                    ["D1"], result, root / "parse.xlsx", root, "test"
                )
        self.assertEqual(len(source), 1)
        self.assertEqual(len(ai), 1)
        self.assertEqual(ai[0]["分摊依据"], "单条记录")
        self.assertEqual(pdfs["D1"].name, "a.pdf")

    def test_backfill_updates_matching_business_row_with_parsed_fields(self):
        record = {
            "record_id": "R1",
            "fields": {
                "供应商": [{"text": "供应商有限公司"}],
                "产品类型(对内)": [{"text": "法兰"}],
                "报关金额AI": 100,
                "报关重量AI": 10,
                "pdf原件": [{"file_token": "old-token"}],
            },
        }

        class FakeClient:
            def batch_update_records(self, _app, _table, updates):
                for update in updates:
                    record["fields"].update(update["fields"])
                return {"records": [record]}

            def batch_create_records(self, _app, _table, _rows):
                raise AssertionError("should update existing business row")

            def upload_bitable_file(self, _app, _path):
                raise AssertionError("should reuse existing PDF token")

        desired = {
            "报关单号": "D1",
            "合同号_1": "26MT-01A100",
            "合同号（应收表格）": "26MT-01A100",
            "销售方式": "出口",
            "出口日期": 1792339200000,
            "报关品名": "不锈钢法兰",
            "海关编码": "7307210000",
            "运输方式": "FOB",
            "运费": 0,
            "保费": 0,
            "报关金额": 100,
            "币种": "USD",
            "报关重量": 10,
            "报关重量单位": "千克",
            "国家": "越南",
            "供应商": "供应商有限公司",
            "产品类型": "法兰",
            "采购金额": 80,
            "分摊依据": "单条记录",
            "拆单状态": "正常",
            "成本匹配状态": "正常",
        }
        with TemporaryDirectory() as temp, patch(
            "app.services.customs_ingest.get_config",
            side_effect=lambda name, _default=None: {"MT_FINANCE_AI_APP_TOKEN": "app", "MT_FINANCE_AI_TABLE_ID": "table"}[name],
        ), patch("app.services.customs_ingest.FeishuClient", return_value=FakeClient()), patch(
            "app.services.customs_ingest.ensure_ai_result_fields",
            return_value={},
        ), patch(
            "app.services.customs_ingest._records_for_declarations",
            return_value={"D1": [record]},
        ):
            summary = write_backfill_to_original(
                declarations=["D1"],
                source_rows=[desired],
                ai_rows=[desired],
                pdfs={"D1": Path("a.pdf")},
                work_dir=Path(temp),
                marker="test",
                wait_seconds=1,
            )
        self.assertEqual(summary.mode, "original-backfill-by-business-key")
        self.assertEqual(record["fields"]["合同号_1"], "26MT-01A100")
        self.assertEqual(record["fields"]["报关品名"], "不锈钢法兰")
        self.assertEqual(record["fields"]["采购金额AI"], 80)
        self.assertNotIn(PROTECTED_PURCHASE_FIELD, record["fields"])


if __name__ == "__main__":
    unittest.main()
