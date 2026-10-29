from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient

from app import main
from app.services import customs_tasks as queue
from app.services import customs_task_runner as runner

PDF = b"%PDF-1.7\nfixture\n%%EOF"


class CustomsTaskQueueTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.dir_patch = patch.object(queue, "TASKS_DIR", self.root / "tasks")
        self.dir_patch.start()

    def tearDown(self):
        self.dir_patch.stop()
        self.temp.cleanup()

    def test_create_claim_idempotency_and_heartbeat(self):
        payload = {
            "flow": "both",
            "startDate": "2026-10-01",
            "endDate": "20261027",
            "idempotencyKey": "button-1",
            "authToken": "must-not-be-stored",
        }
        task, duplicate = queue.create_task(payload, requested_by="finance")
        same, duplicate_again = queue.create_task(payload, requested_by="finance")
        self.assertFalse(duplicate)
        self.assertTrue(duplicate_again)
        self.assertEqual(task["id"], same["id"])
        self.assertEqual(task["parameters"]["startDate"], "2026-10-01")
        self.assertEqual(task["parameters"]["endDate"], "2026-10-27")
        self.assertNotIn("authToken", task["parameters"])

        claimed = queue.claim_next(agent_name="pc-1")
        self.assertEqual(claimed["state"], "claimed")
        self.assertIsNone(queue.claim_next(agent_name="pc-2"))
        heartbeat = queue.heartbeat(task["id"], agent_name="pc-1")
        self.assertEqual(heartbeat["agentName"], "pc-1")

    def test_upload_is_pdf_only_and_cache_is_cleared(self):
        task, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        with self.assertRaises(ValueError):
            queue.save_uploads(
                task["id"],
                [("../bad.pdf", b"<html></html>")],
                agent_name="pc-1",
            )
        self.assertEqual(queue.get_task(task["id"])["state"], "claimed")

        uploaded = queue.save_uploads(
            task["id"],
            [("../declaration.pdf", PDF)],
            agent_name="pc-1",
        )
        self.assertEqual(uploaded["state"], "uploaded")
        self.assertEqual(uploaded["fileCount"], 1)
        self.assertTrue((self.root / "tasks" / task["id"] / "uploads" / "declaration.pdf").exists())

        queue.mark_processing(task["id"])
        queue.mark_completed(task["id"], {"aiRecords": 1})
        queue.clear_pdf_cache(task["id"])
        task_dir = self.root / "tasks" / task["id"]
        self.assertFalse((task_dir / "uploads").exists())
        self.assertTrue((task_dir / "task.json").exists())

    def test_processing_task_recovers_with_existing_pdf(self):
        task, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        queue.save_uploads(task["id"], [("a.pdf", PDF)], agent_name="pc-1")
        queue.mark_processing(task["id"])
        self.assertEqual(queue.recover_interrupted(), [task["id"]])
        self.assertEqual(queue.get_task(task["id"])["state"], "uploaded")

    def test_new_queued_task_wins_over_expired_claim(self):
        stale, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-old")
        stale_task = queue.get_task(stale["id"])
        stale_task["leaseExpiresAt"] = "2000-01-01T00:00:00+00:00"
        queue._save_task(stale_task)
        fresh, _ = queue.create_task({"flow": "pre_record"})

        claimed = queue.claim_next(agent_name="pc-new")
        self.assertEqual(claimed["id"], fresh["id"])

    def test_report_error_blocks_task_and_is_idempotent(self):
        task, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        reported = queue.report_task_error(
            task["id"],
            agent_name="pc-1",
            error_code="CARD_READ_FAILED",
            message="卡介质未连接",
            retryable=True,
            reported_at="2026-10-29T13:24:33+08:00",
            details={"phase": "login"},
        )
        self.assertEqual(reported["state"], "blocked")
        self.assertEqual(reported["errorCode"], "CARD_READ_FAILED")
        self.assertEqual(reported["errorRetryable"], True)
        self.assertTrue(reported["blockedAt"])
        self.assertFalse((self.root / "tasks" / task["id"] / "uploads").exists())

        repeated = queue.report_task_error(
            task["id"],
            agent_name="pc-1",
            error_code="CARD_READ_FAILED",
            message="不同文字不影响同错误码幂等",
            retryable=True,
            reported_at="2026-10-29T13:25:00+08:00",
        )
        self.assertEqual(repeated["errorReceivedAt"], reported["errorReceivedAt"])
        with self.assertRaises(queue.TaskForbidden):
            queue.report_task_error(
                task["id"],
                agent_name="pc-2",
                error_code="CARD_READ_FAILED",
                message="其他电脑",
                retryable=True,
                reported_at="2026-10-29T13:25:00+08:00",
            )

    def test_non_retryable_error_becomes_failed(self):
        task, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        reported = queue.report_task_error(
            task["id"],
            agent_name="pc-1",
            error_code="DOWNLOAD_EMPTY",
            message="没有待下载数据",
            retryable=False,
            reported_at="2026-10-29T13:24:33+08:00",
        )
        self.assertEqual(reported["state"], "failed")
        self.assertEqual(reported["progress"], 100)

    def test_report_error_rejects_expired_lease_and_uploaded_task(self):
        task, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        raw = queue.get_task(task["id"])
        raw["leaseExpiresAt"] = "2000-01-01T00:00:00+00:00"
        queue._save_task(raw)
        with self.assertRaises(queue.TaskConflict):
            queue.report_task_error(
                task["id"],
                agent_name="pc-1",
                error_code="CARD_READ_FAILED",
                message="超时",
                retryable=True,
                reported_at="2026-10-29T13:24:33+08:00",
            )

        queue.mark_failed(task["id"], "lease expired")
        uploaded, _ = queue.create_task({"flow": "export"})
        queue.claim_next(agent_name="pc-1")
        queue.save_uploads(uploaded["id"], [("a.pdf", PDF)], agent_name="pc-1")
        with self.assertRaises(queue.TaskConflict):
            queue.report_task_error(
                uploaded["id"],
                agent_name="pc-1",
                error_code="CARD_READ_FAILED",
                message="已上传",
                retryable=True,
                reported_at="2026-10-29T13:24:33+08:00",
            )


class CustomsTaskHttpTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patches = [
            patch.object(queue, "TASKS_DIR", self.root / "tasks"),
            patch.dict(
                os.environ,
                {
                    "CUSTOMS_TASK_CREATE_TOKEN": "create-secret",
                    "CUSTOMS_AGENT_TOKEN": "agent-secret",
                    "CUSTOMS_FEISHU_VERIFICATION_TOKEN": "verify-secret",
                },
            ),
        ]
        for item in self.patches:
            item.start()
        self.client = TestClient(main.app)

    def tearDown(self):
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def test_feishu_create_claim_upload_and_status(self):
        created = self.client.post(
            "/api/customs/tasks",
            headers={"X-Customs-Task-Token": "create-secret"},
            json={
                "source": "feishu_button",
                "workflow": "customs",
                "flow": "both",
                "startDate": "2026-10-01",
                "endDate": "2026-10-27",
                "idempotencyKey": "record-1",
            },
        )
        self.assertEqual(created.status_code, 200, created.text)
        task_id = created.json()["taskId"]

        claimed = self.client.get(
            "/api/customs/tasks/claim",
            headers={"Authorization": "Bearer agent-secret", "X-Agent-Name": "pc-1"},
        )
        self.assertEqual(claimed.status_code, 200, claimed.text)
        self.assertEqual(claimed.json()["task"]["id"], task_id)

        with patch.object(runner, "start") as start:
            uploaded = self.client.post(
                f"/api/customs/tasks/{task_id}/upload",
                headers={"Authorization": "Bearer agent-secret"},
                data={"agent_name": "pc-1"},
                files={"files": ("declaration.pdf", PDF, "application/pdf")},
            )
        self.assertEqual(uploaded.status_code, 202, uploaded.text)
        start.assert_called_once_with(task_id)

        status = self.client.get(
            f"/api/customs/tasks/{task_id}",
            headers={"Authorization": "Bearer agent-secret"},
        )
        self.assertEqual(status.status_code, 200, status.text)
        self.assertEqual(status.json()["state"], "uploaded")

    def test_url_verification_and_auth_failure(self):
        verified = self.client.post(
            "/api/customs/tasks",
            json={
                "type": "url_verification",
                "token": "verify-secret",
                "challenge": "abc",
            },
        )
        self.assertEqual(verified.status_code, 200, verified.text)
        self.assertEqual(verified.json(), {"challenge": "abc"})

        denied = self.client.post(
            "/api/customs/tasks",
            headers={"X-Customs-Task-Token": "wrong"},
            json={"flow": "both"},
        )
        self.assertEqual(denied.status_code, 403)

    def test_report_error_json_without_files(self):
        created = self.client.post(
            "/api/customs/tasks",
            headers={"X-Customs-Task-Token": "create-secret"},
            json={"flow": "both", "idempotencyKey": "report-error-task"},
        )
        task_id = created.json()["taskId"]
        claimed = self.client.get(
            "/api/customs/tasks/claim",
            headers={"Authorization": "Bearer agent-secret", "X-Agent-Name": "pc-1"},
        )
        self.assertEqual(claimed.json()["task"]["id"], task_id)
        payload = {
            "agentName": "pc-1",
            "errorCode": "CARD_READ_FAILED",
            "message": "卡介质未连接或无法读取",
            "retryable": True,
            "reportedAt": "2026-10-29T13:24:33+08:00",
            "details": {"vendorErrorCode": "50070", "phase": "login"},
        }
        with patch.object(runner, "start") as start:
            response = self.client.post(
                f"/api/customs/tasks/{task_id}/report-error",
                headers={
                    "Authorization": "Bearer agent-secret",
                    "X-Agent-Name": "pc-1",
                },
                json=payload,
            )
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()["state"], "blocked")
        self.assertEqual(response.json()["errorCode"], "CARD_READ_FAILED")
        start.assert_not_called()

        repeated = self.client.post(
            f"/api/customs/tasks/{task_id}/report-error",
            headers={
                "Authorization": "Bearer agent-secret",
                "X-Agent-Name": "pc-1",
            },
            json=payload,
        )
        self.assertEqual(repeated.status_code, 200, repeated.text)
        self.assertEqual(repeated.json()["errorReceivedAt"], response.json()["errorReceivedAt"])

        changed = dict(payload)
        changed["errorCode"] = "LOGIN_TIMEOUT"
        conflict = self.client.post(
            f"/api/customs/tasks/{task_id}/report-error",
            headers={
                "Authorization": "Bearer agent-secret",
                "X-Agent-Name": "pc-1",
            },
            json=changed,
        )
        self.assertEqual(conflict.status_code, 409)

        mismatch = self.client.post(
            f"/api/customs/tasks/{task_id}/report-error",
            headers={
                "Authorization": "Bearer agent-secret",
                "X-Agent-Name": "pc-2",
            },
            json=payload,
        )
        self.assertEqual(mismatch.status_code, 403)


class CustomsTaskRunnerTests(unittest.TestCase):
    def test_runner_writes_to_formal_ai_table(self):
        from app.services.customs_ingest import WriteSummary

        with tempfile.TemporaryDirectory() as temp:
            tasks_dir = Path(temp) / "tasks"
            with patch.object(queue, "TASKS_DIR", tasks_dir):
                task, _ = queue.create_task({"flow": "export"})
                queue.claim_next(agent_name="pc-1")
                queue.save_uploads(task["id"], [("a.pdf", PDF)], agent_name="pc-1")

                def fake_parse(_pdfs, path):
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(b"xlsx")
                    return {"rows": 1, "failed": []}

                def fake_split(_parse, split_dir):
                    split_dir.mkdir(parents=True, exist_ok=True)
                    (split_dir / "拆单明细_全部.xlsx").write_bytes(b"xlsx")
                    return {"stats": {}}

                summary = WriteSummary(
                    marker="test",
                    declarations=["722020260000439783"],
                    sourceRecords=0,
                    aiRecords=1,
                    uploadedFiles=1,
                    workDir=str(tasks_dir / task["id"] / "work"),
                    mode="original-ai-fields-only",
                )
                with patch.object(runner.pipeline, "parse_pdfs", side_effect=fake_parse), \
                     patch.object(runner.pipeline, "run_round", side_effect=fake_split), \
                     patch.object(
                         runner.pipeline,
                         "review_rows",
                         return_value=[{"报关单号": "722020260000439783"}],
                     ), \
                     patch(
                         "app.services.customs_ingest.build_records",
                         return_value=([], [{"报关单号": "722020260000439783"}], {}),
                     ), \
                     patch(
                         "app.services.customs_ingest.write_backfill_to_cost_table",
                         return_value=summary,
                     ) as writer:
                    runner._process(task["id"])
                self.assertTrue(writer.called)
                self.assertEqual(writer.call_args.kwargs["declarations"], ["722020260000439783"])

    def test_runner_completes_and_clears_pdf_cache(self):
        with tempfile.TemporaryDirectory() as temp:
            tasks_dir = Path(temp) / "tasks"
            with patch.object(queue, "TASKS_DIR", tasks_dir):
                task, _ = queue.create_task({"flow": "export"})
                queue.claim_next(agent_name="pc-1")
                queue.save_uploads(task["id"], [("a.pdf", PDF)], agent_name="pc-1")
                fake_result = {
                    "taskId": task["id"],
                    "declarations": ["D1"],
                    "pdfCount": 1,
                    "write": {"aiRecords": 1, "uploadedFiles": 1},
                }
                with patch.object(runner, "_process", return_value=fake_result):
                    runner._run(task["id"])
                completed = queue.get_task(task["id"])
                self.assertEqual(completed["state"], "completed")
                self.assertEqual(completed["summary"]["aiRecords"], 1)
                self.assertFalse((tasks_dir / task["id"] / "uploads").exists())
                self.assertTrue((tasks_dir / task["id"] / "task.json").exists())


if __name__ == "__main__":
    unittest.main()
