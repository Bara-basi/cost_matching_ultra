"""工作台的金额守恒、模板和写回授权回归测试。"""
from __future__ import annotations

import json
import tempfile
import unittest
import urllib.error
from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.main import app
from app.services import contract_shipments, feishu_workflow, lookup, review, shipment_detail, template_import, workspace


class WorkbenchTests(unittest.TestCase):
    def test_manual_adjustment_stays_writable_and_flag_blocks_group(self):
        source = {"record_id": "rec_1", "报关单号": "223120260000174064",
                  "合同号_1": "26MT-01A001", "报关品名": "钢管", "报关金额": "100.00"}
        row = {"报关单号": source["报关单号"], "合同号_1": source["合同号_1"],
               "报关品名": "钢管", "供应商简称": "甲", "报关金额": "100.00",
               "采购金额": "80.00", "异常类型": "正常", "异常明细": ""}
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)):
            root = Path(directory) / "sample"
            root.mkdir()
            (root / "rows.json").write_text(json.dumps([row], ensure_ascii=False), encoding="utf-8")
            (root / "source_records.json").write_text(json.dumps([source], ensure_ascii=False), encoding="utf-8")
            workspace._update("sample", kind="sync", state="complete", createdAt="2026-09-30T00:00:00+00:00")
            self.assertEqual(feishu_workflow.plan("sample")["summary"]["ready"], 1)
            response = TestClient(app).post("/api/jobs/sample/override", json={
                "index": 0, "amount": "82.00", "operator": "财务甲", "reason": "凭证复核"})
            self.assertEqual(response.status_code, 200)
            proposal = feishu_workflow.plan("sample")
            self.assertEqual(proposal["summary"]["ready"], 1)
            self.assertTrue(proposal["groups"][0]["manualAdjusted"])
            # 上报异常仅允许正常记录；独立任务验证阻断与异常汇总。
            (root / "overrides.json").unlink()
            response = TestClient(app).post("/api/jobs/sample/flag-exception", json={
                "index": 0, "operator": "财务乙", "reason": "供应商凭证金额待核实"})
            self.assertEqual(response.status_code, 200)
            self.assertEqual(feishu_workflow.plan("sample")["summary"]["ready"], 0)
            self.assertEqual(workspace.all_exceptions()[0]["reason"], "供应商凭证金额待核实")
            self.assertEqual(workspace.get_job("sample")["summary"]["exceptionRows"], 1)
            self.assertEqual(TestClient(app).post("/api/jobs/sample/flag-exception", json={
                "index": 0, "operator": "财务乙", "reason": "重复"}).status_code, 400)
            (root / "flags.json").unlink()
            (root / "push_results.json").write_text(json.dumps([
                {"sourceId": "rec_1", "status": "written"}]), encoding="utf-8")
            self.assertEqual(TestClient(app).post("/api/jobs/sample/flag-exception", json={
                "index": 0, "operator": "财务乙", "reason": "写回后上报"}).status_code, 400)
            self.assertEqual(workspace.list_jobs()[0]["writebackReady"], 0)

    def test_sync_selects_uncosted_rows_without_pdf_and_completes_contract_context(self):
        def record(key, declaration, contract, amount, *, cost="", pdf=False, day=1780272000000):
            return {"record_id": key, "fields": {"报关单号": declaration, "合同号_1": contract,
                "报关品名": "钢管", "商品序号": key, "报关金额": amount, "报关重量": "10",
                "币种": "USD", "采购金额": cost, "pdf原件": [{"file_token": "pdf"}] if pdf else [],
                "出口日期": day}}

        target = record("target", "223120260000174064", "26MT-01A001", "100")
        extra = record("extra", "223120260000174065", "26MT-01A001", "50", cost="25")
        already = record("already", "223120260000174066", "26MT-01B001", "60", cost="30")

        class FakeClient:
            def iter_search_records(self, _app, _table, filter):
                self.candidate_filter = filter
                yield target
                yield already  # 防御性本地校验也应排除已核算行

            def search_records(self, _app, _table, filter):
                if filter["conditions"][0]["field_name"] == "报关单号":
                    return [target]
                return [target, extra]

        fake = FakeClient()
        with patch.object(feishu_workflow, "_target", return_value=(fake, "app", "table")):
            result = feishu_workflow.select_records(start_date="2026-01-01", limit=5)
        self.assertEqual([item["record_id"] for item in result["selected"]], ["target"])
        self.assertEqual([item["record_id"] for item in result["context"]], ["extra"])
        self.assertFalse(result["selected"][0]["pdfTokens"])
        self.assertEqual(result["selection"]["supplemental"], 1)
        self.assertIn("isEmpty", [entry["operator"] for entry in fake.candidate_filter["conditions"]])
        self.assertNotIn("pdf原件", [entry["field_name"] for entry in fake.candidate_filter["conditions"]])

    def test_sync_scan_uses_feishu_rows_and_never_downloads_pdf(self):
        source = {"record_id": "source", "报关单号": "223120260000174064",
                  "合同号_1": "26MT-01A001", "报关品名": "钢管", "报关金额": "100",
                  "pdfTokens": ["evidence"]}
        chosen = {"selected": [source], "context": [], "skipped": [], "selection": {"total": 1}}
        with patch.object(feishu_workflow, "select_records", return_value=chosen), \
             patch.object(workspace, "start_job", return_value="job") as start, \
             patch.object(workspace, "_update"):
            result = feishu_workflow.start_scan(limit=1)
        self.assertEqual(result["id"], "job")
        self.assertEqual(start.call_args.kwargs["declarations"][0]["总价"], "100")
        self.assertNotIn("files", start.call_args.kwargs)

    def test_historical_split_merges_amount_and_keeps_original_members(self):
        common = {"报关单号": "223120260000174064", "合同号_1": "26MT-01A001",
                  "报关品名": "钢管", "商品序号": "1", "报关重量": "10", "采购金额": ""}
        first = {"record_id": "first", "fields": {**common, "报关金额": "40",
                 "pdf原件": [{"file_token": "evidence"}]}}
        second = {"record_id": "second", "fields": {**common, "报关金额": "60",
                  "pdf原件": []}}

        class FakeClient:
            def iter_search_records(self, _app, _table, _filter):
                yield first
                yield second

            def search_records(self, _app, _table, _filter):
                return [first, second]

        with patch.object(feishu_workflow, "_target", return_value=(FakeClient(), "app", "table")):
            result = feishu_workflow.select_records(limit=5)
        self.assertEqual(result["selection"]["total"], 1)
        self.assertEqual(result["selection"]["merged"], 1)
        self.assertEqual(result["selected"][0]["报关金额"], "100.00")
        self.assertEqual(len(result["selected"][0]["sourceMembers"]), 2)
        self.assertNotIn("writebackBlocked", result["selected"][0])

    def test_domestic_row_is_labeled_by_contract_not_invalid_declaration(self):
        domestic = {"record_id": "domestic", "fields": {"报关单号": "",
                    "合同号_1": "26MT-01A001", "报关品名": "钢管",
                    "报关金额": "100", "采购金额": ""}}

        class FakeClient:
            def iter_search_records(self, _app, _table, _filter):
                yield domestic

        with patch.object(feishu_workflow, "_target", return_value=(FakeClient(), "app", "table")):
            result = feishu_workflow.select_records(limit=5)
        self.assertEqual(result["selection"]["total"], 0)
        self.assertEqual(result["skipped"][0]["tradeType"], "内销")
        self.assertEqual(result["skipped"][0]["合同号_1"], "26MT-01A001")
        self.assertNotIn("18 位", result["skipped"][0]["reason"])

    def test_sync_endpoints_never_surface_connection_failure_as_500(self):
        with patch.object(feishu_workflow, "select_records", side_effect=urllib.error.URLError("offline")):
            response = TestClient(app).post("/api/sync/preview", data={"limit": 1})
        self.assertEqual(response.status_code, 503)
        with patch.object(feishu_workflow, "start_scan", return_value={"id": None, "selection": {}}) as start:
            response = TestClient(app).post("/api/sync/scan", data={"limit": 1})
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("use_ai", start.call_args.kwargs)

    def test_feishu_paste_and_unique_source_enrichment(self):
        self.assertEqual(lookup.parse_paste("报关单号223120260000174064")["rows"][0]["报关单号"],
                         "223120260000174064")
        pasted = lookup.parse_paste(
            "报关单号\t采购订单号\t报关金额\n"
            "223120260000174064\t26MT-01A001-YH\t100.00"
        )
        self.assertEqual(pasted["rows"][0]["报关金额"], "100.00")
        sources = [{"fields": {"报关单号": "223120260000174064", "合同号_1": "26MT-01A001",
                               "报关品名": "钢管", "报关金额": "100.00", "报关重量": 10,
                               "币种": "USD", "pdf原件": [{"file_token": "pdf"}]}}]
        with patch("app.services.erp_cache_index.lookup", return_value={"shipments": [
                {"invoiceCode": "26MT-01A001"}]}):
            enriched = lookup.enrich_rows(pasted["rows"], sources=sources)
        self.assertEqual(enriched["rows"][0]["合同号_1"], "26MT-01A001")
        self.assertEqual(enriched["rows"][0]["报关品名"], "钢管")
        self.assertEqual(lookup.validate_rows(enriched["rows"]), [])

    def test_pasted_row_prefers_contract_with_batch(self):
        """飞书一行常同时有「合同号」和「合同号_1」，要取带批次的那一个。"""
        rows = lookup.parse_paste(
            "224420260014225760 26MT-03R036 26MT-03R036F 不锈钢管件 56 14"
        )["rows"]
        self.assertEqual(rows[0]["合同号_1"], "26MT-03R036F")
        # 同一行的采购订单号带工厂后缀时，基号仍然当合同号
        rows = lookup.parse_paste("26MT-03R036 26MT-03R036-HD")["rows"]
        self.assertEqual(rows[0]["合同号_1"], "26MT-03R036")
        self.assertEqual(rows[0]["采购订单号"], "26MT-03R036-HD")
        # 联合合同号比基号更完整
        rows = lookup.parse_paste("25MT-07F477 25MT-07F477&26MT-07C330")["rows"]
        self.assertEqual(rows[0]["合同号_1"], "25MT-07F477&26MT-07C330")

    def test_amount_claim_stays_inside_order_family(self):
        """按报关金额认领出运单时，只认同一订单家族的出运单。"""
        from scripts.run_split_knapsack import order_family

        self.assertEqual(order_family("26MT-03P315B"), order_family("25MT-03P315-ADD1-A"))
        self.assertNotEqual(order_family("26MT-03P315B"), order_family("25MT-03T614"))
        self.assertEqual(order_family("25MT-03P495Y-ADD1-A"), order_family("25MT-03P495Y-B"))

    def test_contract_batch_missing_in_erp_reports_instead_of_guessing(self):
        """报关合同号写了批次、睿贝却没有这批出运单时，不能拿同订单其它批次顶替。"""
        shipments = [
            {"invoiceCode": "26MT-03R036A", "shipmentId": "1113",
             "orderCode": "PI-26MT-03R036", "purchaseCode": "26MT-03R036-HD"},
            {"invoiceCode": "26MT-03R036E", "shipmentId": "1334",
             "orderCode": "PI-26MT-03R036", "purchaseCode": "26MT-03R036-JX"},
        ]
        with patch("app.services.erp_cache.read_jsonl", return_value=shipments), \
             patch.object(contract_shipments, "read_jsonl", return_value=shipments), \
             patch.object(contract_shipments, "_SHIPMENTS", None):
            self.assertTrue(shipment_detail.names_batch("26MT-03R036F"))
            self.assertFalse(shipment_detail.names_batch("26MT-03R036"))
            # 批次在睿贝里不存在 → 不定位、不兜底
            self.assertEqual(shipment_detail.shipment_invoices_for_contract("26MT-03R036F"), [])
            # 同名出运单 / 没写批次的合同照旧能定位
            self.assertEqual(shipment_detail.shipment_invoices_for_contract("26MT-03R036E"),
                             ["26MT-03R036E"])
            self.assertEqual(set(shipment_detail.shipment_invoices_for_contract("26MT-03R036")),
                             {"26MT-03R036A", "26MT-03R036E"})

    def test_local_split_failure_becomes_finance_readable_row(self):
        """拆单失败的报关行也要出现在结果里，且说明是给财务看的话。"""
        self.assertEqual(review._friendly_local_note("无出运产品行")[0], "找不到出运单")
        self.assertEqual(
            review._friendly_local_note("睿贝里没有这张出运单（26MT-03R036F），该订单只有 A/B/C/D/E 批")[0],
            "找不到出运单",
        )
        self.assertEqual(
            review._friendly_local_note("310120260519759032: 子集和无精确解（容量 73076.83）")[1],
            "报关金额与睿贝出运金额凑不出精确对应，请核对这一单的报关单与出运单。",
        )
        self.assertEqual(
            review._friendly_local_note("费用口径不唯一：小数点调整:add:by_amount、…")[0],
            "拆单失败",
        )

    def test_lookup_collapses_historical_split_rows(self):
        """财务把拆单结果补回表里的多行，粘贴补全时要并回一条报关行。"""
        def record(key, amount, weight):
            return {"record_id": key, "fields": {
                "报关单号": "223120260000174064", "合同号_1": "26MT-01A001",
                "报关品名": "钢管", "报关金额": amount, "报关重量": weight, "币种": "USD"}}

        sources = [record("first", "40", "10"), record("second", "60", "15")]
        # 稀疏粘贴（只给单号 / 合同号）→ 展开成一条合并后的报关行
        parsed = lookup.parse_paste("223120260000174064 26MT-01A001")
        enriched = lookup.enrich_rows(parsed["rows"], sources=sources)
        self.assertEqual(len(enriched["rows"]), 1)
        self.assertEqual(enriched["rows"][0]["报关金额"], "100")
        self.assertEqual(enriched["rows"][0]["报关重量"], "25")
        self.assertIn("并回一条", enriched["rows"][0]["_补全说明"])
        # 粘贴里就是同一报关单的两行拆完单记录 → 同样只留一条，不重复计金额
        parsed = lookup.parse_paste(
            "223120260000174064 26MT-01A001 钢管 40 10\n"
            "223120260000174064 26MT-01A001 钢管 60 15"
        )
        enriched = lookup.enrich_rows(parsed["rows"], sources=sources)
        self.assertEqual(len(enriched["rows"]), 1)
        self.assertEqual(enriched["rows"][0]["报关金额"], "100")

    def test_template_roundtrip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "模板.xlsx"
            path.write_bytes(template_import.template_bytes())
            rows, errors = template_import.read_template(path)
            self.assertEqual(rows, [])
            self.assertEqual(errors[0]["message"], "模板没有商品行")
            book = load_workbook(path)
            sheet = book.active
            columns = {cell.value: cell.column for cell in sheet[1]}
            values = {"报关单号": "223120260000174064", "合同号_1": "26MT-01A001",
                      "报关品名": "钢管", "报关金额": 100, "币种": "USD", "报关重量": 10}
            for key, value in values.items():
                sheet.cell(2, columns[key], value)
            book.save(path)
            book.close()
            rows, errors = template_import.read_template(path)
            self.assertEqual(errors, [])
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["总价"], "100")

    def test_writeback_plan_preserves_original_amount(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)):
            root = Path(directory) / "sample"
            root.mkdir()
            (root / "source_records.json").write_text(json.dumps([{
                "record_id": "rec_1", "报关单号": "123", "合同号_1": "MT",
                "报关品名": "钢管", "报关金额": "100.00"
            }], ensure_ascii=False), encoding="utf-8")
            rows = [
                {"报关单号": "123", "合同号_1": "MT", "报关品名": "钢管", "采购订单号": "PO1", "供应商简称": "甲",
                 "产品类型": "无缝管", "报关金额": "40.00", "采购金额": "20.00", "异常类型": "正常"},
                {"报关单号": "123", "合同号_1": "MT", "报关品名": "钢管", "采购订单号": "PO2", "供应商简称": "乙",
                 "产品类型": "无缝管", "报关金额": "60.00", "采购金额": "30.00", "异常类型": "正常"},
            ]
            (root / "rows.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
            proposal = feishu_workflow.plan("sample")
            self.assertEqual(proposal["summary"]["ready"], 1)
            self.assertEqual(proposal["groups"][0]["resultAmount"], "100.00")
            self.assertEqual(len(proposal["groups"][0]["children"]), 2)
            rows[1]["报关金额"] = "59.99"
            (root / "rows.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
            proposal = feishu_workflow.plan("sample")
            self.assertEqual(proposal["summary"]["blocked"], 1)
            self.assertIn("不守恒", proposal["groups"][0]["reasons"][-1])

    def test_writeback_api_requires_gateway_token(self):
        # 本机（127.0.0.1 / testclient）默认放行写回，只有远端来源才要求网关口令
        remote = TestClient(app, client=("203.0.113.9", 1234))
        response = remote.post("/api/sync/unknown/push", data={"operator": "测试员"})
        self.assertEqual(response.status_code, 403)
        evidence = remote.get("/api/jobs/unknown/erp-evidence/0")
        self.assertEqual(evidence.status_code, 403)
        comparison = remote.get("/api/jobs/unknown/evidence-comparison")
        self.assertEqual(comparison.status_code, 403)
        local = TestClient(app).post("/api/sync/unknown/push", data={"operator": "测试员"})
        self.assertEqual(local.status_code, 400)  # 本机放行，由业务层给出「没有可写回的任务」

    def test_pdf_evidence_compares_input_amount_field(self):
        with tempfile.TemporaryDirectory() as directory:
            def fake_parse(_pdfs, target):
                book = Workbook()
                sheet = book.active
                sheet.append(["报关单号", "合同号_1", "总价"])
                sheet.append(["223120260000174064", "26MT-01A001", 100])
                book.save(target)
                book.close()
                return {"rows": 1}

            with patch.object(workspace.pipeline, "parse_pdfs", side_effect=fake_parse):
                result = workspace._compare_pdf_evidence([{
                    "报关单号": "223120260000174064", "合同号_1": "26MT-01A001",
                    "报关金额": "100.00",
                }], [Path(directory) / "reference.pdf"], Path(directory))
            self.assertEqual(result["rows"][0]["状态"], "结果一致")

    def test_ambiguous_source_stays_unmatched_without_manual_picker(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)):
            root = Path(directory) / "sample"
            root.mkdir()
            sources = [{"record_id": key, "报关单号": "123", "报关品名": "钢管", "报关金额": "40.00"}
                       for key in ("rec_a", "rec_b")]
            (root / "source_records.json").write_text(json.dumps(sources, ensure_ascii=False), encoding="utf-8")
            (root / "rows.json").write_text(json.dumps([{"报关单号": "123", "报关品名": "钢管",
                "报关金额": "40.00", "采购金额": "20.00", "异常类型": "正常"}], ensure_ascii=False), encoding="utf-8")
            self.assertEqual(feishu_workflow.plan("sample")["summary"]["unmatched"], 1)
            with self.assertRaisesRegex(ValueError, "自动匹配"):
                feishu_workflow.set_mapping("sample", 0, "rec_a", "财务甲")

    def test_existing_split_rows_match_business_key_before_write(self):
        first = {"record_id": "rec_a", "fields": {"报关单号": "123", "合同号_1": "MT",
                 "供应商简称": "甲", "报关品名": "钢管", "报关金额": 40}}
        second = {"record_id": "rec_b", "fields": {"报关单号": "123", "合同号_1": "MT",
                  "供应商简称": "乙", "报关品名": "钢管", "报关金额": 60}}
        members = [{"record_id": "rec_a", "报关金额": "40"},
                   {"record_id": "rec_b", "报关金额": "60"}]
        children = [{"报关单号": "123", "合同号_1": "MT", "供应商简称": supplier,
                     "报关品名": "钢管"} for supplier in ("乙", "甲")]
        targets = feishu_workflow._write_targets(
            {"sourceId": "rec_a", "originalAmount": "100", "children": children},
            {"sourceMembers": members}, {"rec_a": first, "rec_b": second})
        self.assertEqual(targets, ["rec_b", "rec_a"])

    def test_writeback_preview_blocks_unmatched_legacy_member(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)):
            root = Path(directory) / "sample"
            root.mkdir()
            members = [{"record_id": key, "报关单号": "123", "合同号_1": "MT",
                        "报关品名": "钢管", "供应商简称": supplier, "报关金额": amount}
                       for key, supplier, amount in (("a", "甲", "40"), ("b", "乙", "60"))]
            sources = [{"record_id": "a", "报关单号": "123", "合同号_1": "MT",
                        "报关品名": "钢管", "报关金额": "100", "sourceMembers": members}]
            rows = [{"报关单号": "123", "合同号_1": "MT", "报关品名": "钢管",
                     "供应商简称": supplier, "报关金额": amount, "采购金额": "20",
                     "异常类型": "正常"}
                    for supplier, amount in (("甲", "40"), ("乙", "60"))]
            (root / "source_records.json").write_text(json.dumps(sources, ensure_ascii=False), encoding="utf-8")
            (root / "rows.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(feishu_workflow.plan("sample")["summary"]["ready"], 1)
            members[1]["供应商简称"] = "丙"
            (root / "source_records.json").write_text(json.dumps(sources, ensure_ascii=False), encoding="utf-8")
            proposal = feishu_workflow.plan("sample")
            self.assertEqual(proposal["summary"]["blocked"], 1)
            self.assertIn("历史拆分行无法全部", proposal["groups"][0]["reasons"][-1])

    def test_push_reuses_historical_split_rows_by_business_key(self):
        members = [{"record_id": "rec_a", "报关金额": "40"},
                   {"record_id": "rec_b", "报关金额": "60"}]
        live = [
            {"record_id": "rec_a", "fields": {"报关单号": "123", "合同号_1": "MT",
                "供应商简称": "甲", "报关品名": "钢管", "报关金额": 40}},
            {"record_id": "rec_b", "fields": {"报关单号": "123", "合同号_1": "MT",
                "供应商简称": "乙", "报关品名": "钢管", "报关金额": 60}},
        ]
        children = [{"splitKey": key, "报关单号": "123", "合同号_1": "MT",
                     "供应商简称": supplier, "报关品名": "钢管", "报关金额": amount,
                     "采购金额": cost, "报关重量": "5"}
                    for key, supplier, amount, cost in (("b", "乙", "60", "30"),
                                                        ("a", "甲", "40", "20"))]
        proposal = {"summary": {"ready": 1}, "groups": [{"sourceId": "rec_a",
                    "status": "ready", "originalAmount": "100", "children": children}]}

        class FakeClient:
            def __init__(self):
                self.updates = []

            def list_fields(self, _app, _table):
                names = ("报关单号", "合同号_1", "供应商简称", "报关品名", "报关金额",
                         "采购金额", "报关重量", feishu_workflow.META_SOURCE,
                         feishu_workflow.META_KEY, feishu_workflow.META_ORIGINAL)
                numeric = {"报关金额", "采购金额", "报关重量", feishu_workflow.META_ORIGINAL}
                return [{"field_name": name, "type": 2 if name in numeric else 1} for name in names]

            def iter_records(self, _app, _table):
                return iter(live)

            def create_record(self, *_args):
                raise AssertionError("已有联合主键行，不应新增")

            def update_record(self, _app, _table, record_id, values):
                self.updates.append((record_id, values))

        fake = FakeClient()
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)), \
             patch.object(feishu_workflow, "_target", return_value=(fake, "app", "table")), \
             patch.object(feishu_workflow, "plan", return_value=proposal), \
             patch.object(workspace, "append_audit"):
            root = Path(directory) / "sample"
            root.mkdir()
            (root / "source_records.json").write_text(json.dumps([{
                "record_id": "rec_a", "sourceMembers": members}], ensure_ascii=False), encoding="utf-8")
            result = feishu_workflow.push("sample", "财务甲")
        self.assertEqual((result["written"], result["failed"]), (1, 0))
        self.assertEqual(result["results"][0]["created"], 0)
        self.assertEqual([(record_id, values["报关金额"]) for record_id, values in fake.updates],
                         [("rec_b", 60.0), ("rec_a", 40.0)])

    def test_push_updates_original_with_first_split_amount(self):
        class FakeClient:
            def __init__(self):
                self.calls = []

            def list_fields(self, _app, _table):
                names = ("报关单号", "合同号_1", "供应商简称", "报关品名", "报关金额", "采购金额", "报关重量",
                         feishu_workflow.META_SOURCE, feishu_workflow.META_KEY,
                         feishu_workflow.META_ORIGINAL)
                return [{"field_name": name, "type": 2 if name in ("报关金额", "采购金额", "报关重量", feishu_workflow.META_ORIGINAL) else 1}
                        for name in names]

            def iter_records(self, _app, _table):
                return iter([{"record_id": "rec_1", "fields": {"报关金额": 100.0}}])

            def create_record(self, _app, _table, values):
                self.calls.append(("create", values))
                return {"record": {"record_id": "rec_child"}}

            def update_record(self, _app, _table, record_id, values):
                self.calls.append(("update", record_id, values))

        children = [
            {"splitKey": "first", "报关单号": "123", "合同号_1": "MT", "供应商简称": "甲", "报关品名": "钢管",
             "报关金额": "40.00", "采购金额": "20.00", "报关重量": "4"},
            {"splitKey": "second", "报关单号": "123", "合同号_1": "MT", "供应商简称": "乙", "报关品名": "钢管",
             "报关金额": "60.00", "采购金额": "30.00", "报关重量": "6"},
        ]
        proposal = {"summary": {"ready": 1}, "groups": [{"sourceId": "rec_1", "status": "ready",
                    "originalAmount": "100.00", "children": children}], "unmatched": []}
        fake = FakeClient()
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)), \
             patch.object(feishu_workflow, "_target", return_value=(fake, "app", "table")), \
             patch.object(feishu_workflow, "plan", return_value=proposal), \
             patch.object(workspace, "append_audit"):
            (Path(directory) / "sample").mkdir()
            result = feishu_workflow.push("sample", "测试员")
        self.assertEqual(result["written"], 1)
        self.assertEqual(fake.calls[0][0], "create")
        self.assertEqual(fake.calls[0][1]["报关金额"], 60.0)
        self.assertEqual(fake.calls[1][0:2], ("update", "rec_1"))
        self.assertEqual(fake.calls[1][2]["报关金额"], 40.0)
        self.assertEqual(fake.calls[1][2]["采购金额"], 20.0)

        class BrokenClient(FakeClient):
            def create_record(self, _app, _table, values):
                raise RuntimeError("模拟新增子行失败")

        broken = BrokenClient()
        with tempfile.TemporaryDirectory() as directory, patch.object(workspace, "JOBS_DIR", Path(directory)), \
             patch.object(feishu_workflow, "_target", return_value=(broken, "app", "table")), \
             patch.object(feishu_workflow, "plan", return_value=proposal), \
             patch.object(workspace, "append_audit"):
            (Path(directory) / "sample").mkdir()
            result = feishu_workflow.push("sample", "测试员")
        self.assertEqual(result["failed"], 1)
        self.assertFalse(any(call[0] == "update" and call[1] == "rec_1" for call in broken.calls))


if __name__ == "__main__":
    unittest.main()
