#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""بایگانی‌کنندهٔ سادهٔ صفحات وب.

این ابزار از یک نشانی شروع می‌کند و فقط صفحات HTML مجازِ همان مبدأ را که
در محدودهٔ انتخاب‌شده قرار دارند دانلود می‌کند. برای نمونه:

    python scripts/simple_web_scraper.py https://example.org/docs/ \
        --output web_archive

به robots.txt احترام می‌گذارد، بین درخواست‌ها مکث می‌کند و برای جلوگیری از
خزش ناخواسته، حد صفحه، عمق و اندازهٔ هر پاسخ دارد. فقط از کتابخانهٔ استاندارد
پایتون استفاده می‌کند.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Deque, Dict, Iterable, List, Optional, Set, Tuple
from urllib import error, robotparser
from urllib.parse import quote, urljoin, urlsplit, urlunsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener, urlopen


DEFAULT_USER_AGENT = "SimpleWebArchive/1.0"
HTML_MIME_TYPES = {"text/html", "application/xhtml+xml"}
PAGE_LINK_TAGS = {"a", "area", "frame", "iframe"}
SAFE_PATH_CHARS = "/%:@!$&'()*+,;=-._~"
SAFE_QUERY_CHARS = "%:@!$&'()*+,;=/?-._~"


class ScraperError(Exception):
    """خطای قابل‌نمایش برای کاربر ابزار."""


class OutOfScopeRedirect(ScraperError):
    """تغییرمسیر HTTP که پیش از درخواستِ مقصد، از محدوده خارج شده است."""

    def __init__(self, target_url: str) -> None:
        self.target_url = target_url
        super().__init__("تغییرمسیر به خارج از محدوده: {}".format(target_url))


@dataclass(frozen=True)
class FetchResult:
    """نتیجهٔ دریافت یک پاسخ HTTP."""

    requested_url: str
    final_url: str
    http_status: int
    content_type: str
    body: bytes
    charset: str


class LinkParser(HTMLParser):
    """استخراج‌کنندهٔ پیوندهای صفحه، بدون اجرای JavaScript."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.links: List[str] = []
        self.base_href: Optional[str] = None

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        values = {key.lower(): value for key, value in attrs}
        if tag.lower() == "base" and self.base_href is None and values.get("href"):
            self.base_href = values["href"]
        if tag.lower() in PAGE_LINK_TAGS and values.get("href"):
            self.links.append(values["href"] or "")


def normalize_url(raw_url: str, base_url: Optional[str] = None) -> Optional[str]:
    """یک نشانی HTTP(S) مطلق و بدون fragment برمی‌گرداند.

    ``None`` برای نشانی‌های نامعتبر، غیر HTTP(S)، دارای مشخصات ورود، یا پیوند
    صرفاً به یک fragment بازگردانده می‌شود. بخش query حفظ می‌شود؛ زیرا در بسیاری
    از وب‌سایت‌ها صفحه‌بندی به آن وابسته است.
    """
    if not raw_url or not raw_url.strip():
        return None

    try:
        resolved = urljoin(base_url, raw_url) if base_url else raw_url
        parsed = urlsplit(resolved)
        scheme = parsed.scheme.lower()
        if scheme not in {"http", "https"} or not parsed.hostname:
            return None
        if parsed.username is not None or parsed.password is not None:
            return None
        port = parsed.port  # باعث می‌شود شماره‌درگاه نامعتبر همین‌جا رد شود.
    except (TypeError, ValueError):
        return None

    try:
        hostname = parsed.hostname.encode("idna").decode("ascii").lower()
    except UnicodeError:
        return None

    # IPv6 در netloc باید در براکت بماند.
    host = "[{}]".format(hostname) if ":" in hostname else hostname
    if port is not None and not ((scheme == "http" and port == 80) or (scheme == "https" and port == 443)):
        host = "{}:{}".format(host, port)

    # quote فقط نویسه‌های غیر ASCII را کدگذاری می‌کند و درصدهای موجود را حفظ می‌کند.
    path = quote(parsed.path or "/", safe=SAFE_PATH_CHARS)
    query = quote(parsed.query, safe=SAFE_QUERY_CHARS)
    return urlunsplit((scheme, host, path, query, ""))


def origin_key(url: str) -> Tuple[str, str, Optional[int]]:
    """کلید مبدأ URL با درنظرگرفتن درگاه پیش‌فرض."""
    parsed = urlsplit(url)
    return parsed.scheme, (parsed.hostname or "").lower(), parsed.port


def crawl_prefix(start_url: str) -> Tuple[str, str]:
    """URL آغاز و پیشوند مسیر مجاز در حالت ``path`` را تعیین می‌کند."""
    path = urlsplit(start_url).path or "/"
    if path.endswith("/"):
        return start_url, path

    leaf = path.rsplit("/", 1)[-1]
    # برای /guide، صفحه و فرزندان /guide/... مناسب‌تر از کل دامنه‌اند؛ اما یک
    # فایل مشخص مانند /guide/index.html باید محتوای پوشه‌اش را دنبال کند.
    if "." in leaf:
        return start_url, path.rsplit("/", 1)[0] + "/"
    return start_url, path + "/"


def is_in_scope(url: str, start_url: str, path_prefix: str, scope: str) -> bool:
    """بررسی می‌کند که URL در مبدأ و محدودهٔ مورد نظر باقی مانده است."""
    if origin_key(url) != origin_key(start_url):
        return False
    if scope == "site":
        return True

    candidate_path = urlsplit(url).path or "/"
    initial_path = urlsplit(start_url).path or "/"
    return candidate_path == initial_path or candidate_path.startswith(path_prefix)


class ScopeRedirectHandler(HTTPRedirectHandler):
    """اجازهٔ تغییرمسیر HTTP را فقط در محدودهٔ خزنده می‌دهد.

    بازگرداندن ``None`` از redirect handler باعث پاسخ‌خوانی پیش‌فرض می‌شود؛ به
    همین علت برای تغییرمسیر بیرون از محدوده یک خطای روشن پرتاب می‌کنیم تا حتی
    یک درخواست GET به مقصد بیرونی ارسال نشود.
    """

    def __init__(self, start_url: str, path_prefix: str, scope: str) -> None:
        super().__init__()
        self.start_url = start_url
        self.path_prefix = path_prefix
        self.scope = scope

    def redirect_request(self, req: Request, fp: object, code: int, msg: str, headers: object, newurl: str) -> Request:
        target_url = normalize_url(newurl, req.full_url)
        if target_url is None or not is_in_scope(target_url, self.start_url, self.path_prefix, self.scope):
            raise OutOfScopeRedirect(target_url or newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


class RobotsPolicy:
    """سیاست robots.txt با رفتار محافظه‌کارانه هنگام خطا."""

    def __init__(self, parser: Optional[robotparser.RobotFileParser], status: str) -> None:
        self._parser = parser
        self.status = status

    @classmethod
    def load(cls, start_url: str, user_agent: str, timeout: float) -> "RobotsPolicy":
        parsed = urlsplit(start_url)
        robots_url = urlunsplit((parsed.scheme, parsed.netloc, "/robots.txt", "", ""))
        request = Request(
            robots_url,
            headers={"User-Agent": user_agent, "Accept": "text/plain,*/*;q=0.1"},
        )
        try:
            with urlopen(request, timeout=timeout) as response:
                status = response.getcode() or 200
                raw = response.read(1_000_000)
                charset = response.headers.get_content_charset() or "utf-8"
                text = raw.decode(charset, errors="replace")
        except error.HTTPError as exc:
            if exc.code == 404:
                return cls(None, "not-found (همهٔ URLها مجاز فرض شدند)")
            return cls(None, "HTTP {} (خزش متوقف شد)".format(exc.code))
        except (error.URLError, OSError, ValueError) as exc:
            # اگر نتوان robots را بررسی کرد، به جای حدسِ مجازبودن خزش نمی‌کنیم.
            return cls(None, "غیرقابل‌دسترسی (خزش متوقف شد): {}".format(exc.reason if isinstance(exc, error.URLError) else exc))

        if not 200 <= status < 300:
            return cls(None, "HTTP {} (خزش متوقف شد)".format(status))

        parser = robotparser.RobotFileParser()
        parser.set_url(robots_url)
        parser.parse(text.splitlines())
        return cls(parser, "loaded")

    def can_fetch(self, user_agent: str, url: str) -> bool:
        # فقدان robots.txt (۴۰۴) به‌معنای نبود قانون است؛ ولی همهٔ خطاهای دیگر
        # در load با وضعیت «خزش متوقف شد» ثبت و اینجا رد می‌شوند.
        if self.status.startswith("not-found"):
            return True
        if self._parser is None:
            return False
        return self._parser.can_fetch(user_agent, url)


def safe_component(value: str, fallback: str) -> str:
    """نام یک بخش مسیر را برای فایل‌سیستم امن و قابل‌حمل می‌کند."""
    value = re.sub(r'[<>:"\\|?*\x00-\x1f]', "_", value).strip(". ")
    value = value or fallback
    if len(value) > 100:
        digest = hashlib.sha256(value.encode("utf-8")).hexdigest()[:12]
        value = "{}__{}".format(value[:80], digest)
    return value


class ArchiveWriter:
    """تبدیل URLها به مسیرهای یکتا و نوشتن اتمی فایل‌ها."""

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self._allocated: Dict[Path, str] = {}

    def relative_path_for(self, url: str) -> Path:
        parsed = urlsplit(url)
        host = safe_component(parsed.netloc, "site")
        raw_parts = [part for part in parsed.path.split("/") if part]
        parts = [safe_component(part, "page") for part in raw_parts]

        if not parts or parsed.path.endswith("/"):
            directories, filename = parts, "index.html"
        else:
            directories, filename = parts[:-1], parts[-1]
            suffix = Path(filename).suffix.lower()
            if suffix not in {".html", ".htm", ".xhtml"}:
                filename += ".html"

        if parsed.query:
            query_hash = hashlib.sha256(parsed.query.encode("utf-8")).hexdigest()[:12]
            stem = Path(filename).stem
            suffix = Path(filename).suffix or ".html"
            filename = "{}__q_{}{}".format(stem, query_hash, suffix)

        candidate = Path(host, *directories, filename)
        previously_allocated_to = self._allocated.get(candidate)
        if previously_allocated_to not in (None, url):
            suffix = candidate.suffix or ".html"
            unique = hashlib.sha256(url.encode("utf-8")).hexdigest()[:12]
            candidate = candidate.with_name("{}__{}{}".format(candidate.stem, unique, suffix))
        self._allocated[candidate] = url
        return candidate

    def save(self, url: str, body: bytes) -> str:
        relative_path = self.relative_path_for(url)
        target = self.output_dir / relative_path
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(target.name + ".part")
        temporary.write_bytes(body)
        temporary.replace(target)
        return relative_path.as_posix()


def decode_html(body: bytes, charset: str) -> str:
    """بدنهٔ HTML را حتی در صورت charset نامعتبر، با جایگزینی امن decode می‌کند."""
    try:
        return body.decode(charset, errors="replace")
    except (LookupError, TypeError):
        return body.decode("utf-8", errors="replace")


class WebScraper:
    """خزندهٔ ترتیبی و محدود به یک مبدأ برای بایگانی HTML."""

    def __init__(
        self,
        start_url: str,
        output_dir: Path,
        *,
        scope: str,
        max_pages: int,
        max_depth: int,
        delay: float,
        timeout: float,
        max_file_bytes: int,
        user_agent: str,
    ) -> None:
        normalized_start = normalize_url(start_url)
        if normalized_start is None:
            raise ScraperError("نشانی آغاز باید یک URL معتبر http:// یا https:// باشد.")
        if max_pages < 0 or max_depth < 0 or delay < 0 or timeout <= 0 or max_file_bytes <= 0:
            raise ScraperError("حدها نامعتبرند؛ زمان/اندازه باید مثبت و عمق/تعداد صفحه باید صفر یا بیشتر باشند.")

        self.start_url, self.path_prefix = crawl_prefix(normalized_start)
        self.output_dir = output_dir
        self.scope = scope
        self.max_pages = max_pages
        self.max_depth = max_depth
        self.delay = delay
        self.timeout = timeout
        self.max_file_bytes = max_file_bytes
        self.user_agent = user_agent
        self.writer = ArchiveWriter(output_dir)
        self.opener = build_opener(ScopeRedirectHandler(self.start_url, self.path_prefix, self.scope))
        self.records: List[Dict[str, object]] = []
        self._last_request_at: Optional[float] = None

    def _wait_before_request(self) -> None:
        if self._last_request_at is None or self.delay == 0:
            return
        remaining = self.delay - (time.monotonic() - self._last_request_at)
        if remaining > 0:
            time.sleep(remaining)

    def _fetch_html(self, url: str) -> FetchResult:
        self._wait_before_request()
        request = Request(
            url,
            headers={
                "User-Agent": self.user_agent,
                "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.1",
            },
        )
        try:
            with self.opener.open(request, timeout=self.timeout) as response:
                self._last_request_at = time.monotonic()
                status = response.getcode() or 200
                content_type = response.headers.get_content_type().lower()
                if content_type not in HTML_MIME_TYPES:
                    raise ScraperError("نوع محتوا HTML نیست: {}".format(content_type or "نامشخص"))
                body = response.read(self.max_file_bytes + 1)
                if len(body) > self.max_file_bytes:
                    raise ScraperError("اندازهٔ پاسخ از حد مجاز بیشتر است")
                final_url = normalize_url(response.geturl())
                if final_url is None:
                    raise ScraperError("سرور پس از تغییرمسیر، URL نامعتبر برگرداند")
                return FetchResult(
                    requested_url=url,
                    final_url=final_url,
                    http_status=status,
                    content_type=content_type,
                    body=body,
                    charset=response.headers.get_content_charset() or "utf-8",
                )
        except error.HTTPError as exc:
            self._last_request_at = time.monotonic()
            raise ScraperError("HTTP {}: {}".format(exc.code, exc.reason)) from exc
        except error.URLError as exc:
            self._last_request_at = time.monotonic()
            raise ScraperError("خطای شبکه: {}".format(exc.reason)) from exc
        except (OSError, ValueError) as exc:
            self._last_request_at = time.monotonic()
            raise ScraperError("خطا در دریافت: {}".format(exc)) from exc

    def _record(self, **values: object) -> None:
        self.records.append(values)

    def run(self) -> Dict[str, object]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        robots = RobotsPolicy.load(self.start_url, self.user_agent, self.timeout)
        if robots.status != "loaded" and not robots.status.startswith("not-found"):
            self._record(url=self.start_url, depth=0, status="stopped", reason="robots.txt: {}".format(robots.status))
            return self._write_manifest(robots.status)

        pending: Deque[Tuple[str, int]] = deque([(self.start_url, 0)])
        queued: Set[str] = {self.start_url}
        final_urls: Set[str] = set()
        attempted = 0

        while pending and (self.max_pages == 0 or attempted < self.max_pages):
            url, depth = pending.popleft()
            attempted += 1

            if not robots.can_fetch(self.user_agent, url):
                self._record(url=url, depth=depth, status="skipped", reason="robots.txt اجازهٔ دریافت نمی‌دهد")
                continue

            try:
                fetched = self._fetch_html(url)
            except OutOfScopeRedirect as exc:
                self._record(url=url, depth=depth, status="skipped", reason=str(exc))
                continue
            except ScraperError as exc:
                self._record(url=url, depth=depth, status="failed", reason=str(exc))
                continue

            if not is_in_scope(fetched.final_url, self.start_url, self.path_prefix, self.scope):
                self._record(
                    url=url,
                    final_url=fetched.final_url,
                    depth=depth,
                    status="skipped",
                    reason="تغییرمسیر به خارج از محدوده",
                )
                continue
            if fetched.final_url in final_urls:
                self._record(
                    url=url,
                    final_url=fetched.final_url,
                    depth=depth,
                    status="skipped",
                    reason="محتوای تغییرمسیرشده قبلاً ذخیره شده است",
                )
                continue

            final_urls.add(fetched.final_url)
            local_path = self.writer.save(fetched.final_url, fetched.body)
            self._record(
                url=url,
                final_url=fetched.final_url,
                depth=depth,
                status="saved",
                http_status=fetched.http_status,
                content_type=fetched.content_type,
                local_path=local_path,
                bytes=len(fetched.body),
            )
            print("[saved] {} -> {}".format(fetched.final_url, local_path))

            if depth >= self.max_depth:
                continue

            parser = LinkParser()
            parser.feed(decode_html(fetched.body, fetched.charset))
            parser.close()
            base_url = urljoin(fetched.final_url, parser.base_href) if parser.base_href else fetched.final_url
            for href in parser.links:
                candidate = normalize_url(href, base_url)
                if candidate is None or candidate in queued:
                    continue
                if not is_in_scope(candidate, self.start_url, self.path_prefix, self.scope):
                    continue
                queued.add(candidate)
                pending.append((candidate, depth + 1))

        if pending and self.max_pages:
            self._record(
                status="limit-reached",
                reason="حد {} صفحه رسید؛ {} URL دیگر در صف ماند.".format(self.max_pages, len(pending)),
            )
        return self._write_manifest(robots.status)

    def _write_manifest(self, robots_status: str) -> Dict[str, object]:
        saved = sum(1 for item in self.records if item.get("status") == "saved")
        failed = sum(1 for item in self.records if item.get("status") == "failed")
        summary: Dict[str, object] = {
            "created_at": datetime.now(timezone.utc).isoformat(),
            "start_url": self.start_url,
            "scope": self.scope,
            "path_prefix": self.path_prefix if self.scope == "path" else None,
            "robots_txt": robots_status,
            "saved_pages": saved,
            "failed_pages": failed,
            "records": self.records,
        }
        manifest_path = self.output_dir / "manifest.json"
        manifest_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return summary


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="دانلود کنترل‌شدهٔ صفحه‌های HTML یک وب‌سایت در یک پوشهٔ محلی.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("url", help="نشانی صفحهٔ آغاز، شامل http:// یا https://")
    parser.add_argument("--output", "-o", type=Path, default=Path("web_archive"), help="پوشهٔ بایگانی")
    parser.add_argument(
        "--scope",
        choices=("path", "site"),
        default="path",
        help="path: صفحه و مسیر زیر آن؛ site: همهٔ صفحه‌های همان مبدأ",
    )
    parser.add_argument("--max-pages", type=int, default=250, help="بیشترین صفحه؛ صفر یعنی بدون حد")
    parser.add_argument("--max-depth", type=int, default=8, help="بیشترین عمق پیوند؛ صفر یعنی فقط صفحهٔ آغاز")
    parser.add_argument("--delay", type=float, default=0.5, help="حداقل مکث میان درخواست‌های صفحه، بر حسب ثانیه")
    parser.add_argument("--timeout", type=float, default=20.0, help="مهلت هر درخواست، بر حسب ثانیه")
    parser.add_argument("--max-file-mb", type=float, default=10.0, help="بیشترین اندازهٔ هر صفحهٔ HTML، بر حسب مگابایت")
    parser.add_argument("--user-agent", default=DEFAULT_USER_AGENT, help="User-Agent ارسالی")
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)
    max_file_bytes = int(args.max_file_mb * 1024 * 1024)
    try:
        scraper = WebScraper(
            args.url,
            args.output,
            scope=args.scope,
            max_pages=args.max_pages,
            max_depth=args.max_depth,
            delay=args.delay,
            timeout=args.timeout,
            max_file_bytes=max_file_bytes,
            user_agent=args.user_agent,
        )
        summary = scraper.run()
    except ScraperError as exc:
        print("خطا: {}".format(exc), file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("\nخزش توسط کاربر متوقف شد.", file=sys.stderr)
        return 130

    print(
        "پایان: {} صفحه ذخیره شد، {} خطا. گزارش: {}".format(
            summary["saved_pages"], summary["failed_pages"], args.output / "manifest.json"
        )
    )
    # خطا در یک صفحه نباید بایگانیِ موفق را به خطای اجرای کلی تبدیل کند؛ جزئیات
    # کامل همهٔ پاسخ‌ها در manifest.json ذخیره می‌شود.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
