from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.services.erp_web import launch_browser


class ErpBrowserTests(unittest.TestCase):
    def test_launch_falls_back_to_configured_chromium(self):
        class Chromium:
            def __init__(self):
                self.calls = []

            def launch(self, **kwargs):
                self.calls.append(kwargs)
                if "executable_path" not in kwargs:
                    raise RuntimeError("bundled browser missing")
                return "browser"

        class Playwright:
            def __init__(self):
                self.chromium = Chromium()

        with tempfile.TemporaryDirectory() as temp:
            executable = Path(temp) / "chrome"
            executable.write_bytes(b"fixture")
            playwright = Playwright()
            with patch.dict(
                os.environ,
                {"PLAYWRIGHT_CHROMIUM_EXECUTABLE": str(executable)},
            ):
                browser = launch_browser(playwright, headless=True)
        self.assertEqual(browser, "browser")
        self.assertEqual(len(playwright.chromium.calls), 2)
        self.assertEqual(
            playwright.chromium.calls[1]["executable_path"],
            str(executable),
        )


if __name__ == "__main__":
    unittest.main()
