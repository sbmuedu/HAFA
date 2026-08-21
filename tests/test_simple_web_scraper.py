# -*- coding: utf-8 -*-
"""آزمون یکپارچهٔ ساده برای وب‌اسکرپر، فقط با یک وب‌سرور محلی."""
from __future__ import annotations

import importlib.util
import json
import sys
import threading
import unittest
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "simple_web_scraper.py"
spec = importlib.util.spec_from_file_location("simple_web_scraper", SCRIPT)
assert spec is not None and spec.loader is not None
scraper_module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = scraper_module
spec.loader.exec_module(scraper_module)
WebScraper = scraper_module.WebScraper


class QuietHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        pass


class RedirectingHandler(QuietHandler):
    """سرور آزمایشی که یک URL داخل محدوده را به یک sibling خارج از آن می‌فرستد."""

    def do_GET(self) -> None:
        if self.path == "/docs/redirect.html":
            self.send_response(302)
            self.send_header("Location", "/outside.html")
            self.end_headers()
            return
        if self.path == "/outside.html":
            self.server.outside_was_requested = True  # type: ignore[attr-defined]
        super().do_GET()


class SimpleWebScraperTest(unittest.TestCase):
    def test_downloads_descendants_but_not_sibling_or_external_link(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary) / "site"
            docs = root / "docs"
            docs.mkdir(parents=True)
            (docs / "index.html").write_text(
                '<a href="child.html">child</a>'
                '<a href="child.html#section">same child</a>'
                '<a href="../outside.html">outside</a>'
                '<a href="https://example.com/">external</a>',
                encoding="utf-8",
            )
            (docs / "child.html").write_text('<a href="nested/grandchild.html">grandchild</a>', encoding="utf-8")
            (docs / "nested").mkdir()
            (docs / "nested" / "grandchild.html").write_text("done", encoding="utf-8")
            (root / "outside.html").write_text("must not be downloaded", encoding="utf-8")

            handler = partial(QuietHandler, directory=str(root))
            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                output = Path(temporary) / "archive"
                start = "http://127.0.0.1:{}/docs/index.html".format(server.server_port)
                summary = WebScraper(
                    start,
                    output,
                    scope="path",
                    max_pages=10,
                    max_depth=5,
                    delay=0,
                    timeout=5,
                    max_file_bytes=1024 * 1024,
                    user_agent="SimpleWebArchive-test",
                ).run()
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=5)

            self.assertEqual(summary["saved_pages"], 3)
            manifest = json.loads((output / "manifest.json").read_text(encoding="utf-8"))
            saved_paths = {record["local_path"] for record in manifest["records"] if record["status"] == "saved"}
            self.assertEqual(len(saved_paths), 3)
            self.assertFalse(any("outside" in path for path in saved_paths))
            for saved_path in saved_paths:
                self.assertTrue((output / saved_path).is_file())

    def test_does_not_request_redirect_target_outside_path_scope(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary) / "site"
            docs = root / "docs"
            docs.mkdir(parents=True)
            (docs / "index.html").write_text('<a href="redirect.html">redirect</a>', encoding="utf-8")
            (root / "outside.html").write_text("must not be requested", encoding="utf-8")

            handler = partial(RedirectingHandler, directory=str(root))
            server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
            server.outside_was_requested = False  # type: ignore[attr-defined]
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                output = Path(temporary) / "archive"
                start = "http://127.0.0.1:{}/docs/index.html".format(server.server_port)
                summary = WebScraper(
                    start,
                    output,
                    scope="path",
                    max_pages=10,
                    max_depth=5,
                    delay=0,
                    timeout=5,
                    max_file_bytes=1024 * 1024,
                    user_agent="SimpleWebArchive-test",
                ).run()
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=5)

            self.assertEqual(summary["saved_pages"], 1)
            self.assertFalse(server.outside_was_requested)  # type: ignore[attr-defined]
            self.assertTrue(
                any("تغییرمسیر به خارج از محدوده" in str(record.get("reason", "")) for record in summary["records"])
            )


if __name__ == "__main__":
    unittest.main()
