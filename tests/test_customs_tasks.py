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
                         "app.services.customs_ingest.write_to_original",
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
