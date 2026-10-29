from __future__ import annotations

import json
import unittest
from unittest.mock import patch

from scripts.mcp_erp import McpClient


class _Response:
    def __init__(self):
        self.headers = {}

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        return json.dumps(
            {"jsonrpc": "2.0", "id": 1, "result": {"tools": []}}
        ).encode("utf-8")


class McpClientTests(unittest.TestCase):
    def test_requests_include_current_protocol_version(self):
        captured = []

        def fake_urlopen(request, timeout):
            captured.append(request)
            return _Response()

        client = McpClient("token")
        with patch("scripts.mcp_erp.urllib.request.urlopen", side_effect=fake_urlopen):
            client._post({"jsonrpc": "2.0", "id": 1, "method": "tools/list"})

        self.assertEqual(
            captured[0].get_header("Mcp-protocol-version"),
            "2025-06-18",
        )


if __name__ == "__main__":
    unittest.main()
