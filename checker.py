"""
URL checking engine — no GUI code, so it can be tested and reused from scripts.
"""
from __future__ import annotations

import csv
import io
import ssl
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional
from urllib.parse import urlsplit

import httpx

APP_VERSION = "1.1.0"
USER_AGENT = f"Mozilla/5.0 (compatible; URLHealthChecker/{APP_VERSION})"


@dataclass
class CheckResult:
    """Result of a single URL health check."""

    url: str
    status: int | str
    response_time_ms: float
    final_url: str = ""
    content_type: str = ""
    redirects: int = 0
    error: str = ""
    checked_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    @property
    def category(self) -> str:
        """ok (2xx) · redirect (ended on a different URL) · client (4xx) · server (5xx) · timeout · error."""
        if isinstance(self.status, int):
            if 200 <= self.status < 300:
                return "redirect" if self.redirects else "ok"
            if 300 <= self.status < 400:
                return "redirect"
            return "client" if self.status < 500 else "server"
        return "timeout" if self.status == "TIMEOUT" else "error"

    @property
    def is_ok(self) -> bool:
        return isinstance(self.status, int) and 200 <= self.status < 400


@dataclass
class CheckOptions:
    timeout: float = 10.0
    max_redirects: int = 5
    retries: int = 1           # extra attempts for timeouts / connection errors
    workers: int = 8
    verify_ssl: bool = True


# ---------------------------------------------------------------------------
# Input parsing
# ---------------------------------------------------------------------------

def normalize_url(raw: str) -> Optional[str]:
    """Trim, add https:// when missing, and reject things that are clearly not URLs."""
    url = raw.strip().strip('"').strip("'")
    if not url or url.startswith("#"):
        return None
    if not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    parts = urlsplit(url)
    host = parts.hostname or ""
    if not host or (" " in host) or ("." not in host and host != "localhost"):
        return None
    return url


def parse_url_list(text: str) -> tuple[list[str], list[str]]:
    """Return (unique valid URLs in original order, rejected lines). Accepts commas, spaces, newlines."""
    seen: set[str] = set()
    urls: list[str] = []
    rejected: list[str] = []
    for line in text.replace(",", "\n").splitlines():
        if line.strip().startswith("#"):
            continue
        for token in line.split():  # URLs never contain whitespace
            url = normalize_url(token)
            if url is None:
                rejected.append(token)
            elif url.lower() not in seen:
                seen.add(url.lower())
                urls.append(url)
    return urls, rejected


def read_url_file(path: str | Path) -> str:
    """Read URLs from .txt, .csv or .xlsx (every cell is considered)."""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl

        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        values = [str(c) for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for c in row if c]
        wb.close()
        return "\n".join(values)
    text = p.read_text(encoding="utf-8-sig", errors="replace")
    if p.suffix.lower() == ".csv":
        return "\n".join(cell for row in csv.reader(io.StringIO(text)) for cell in row)
    return text


# ---------------------------------------------------------------------------
# Checking
# ---------------------------------------------------------------------------

def _describe(exc: Exception) -> tuple[str, str]:
    if isinstance(exc, httpx.TimeoutException):
        return "TIMEOUT", "Request timed out"
    if isinstance(exc, httpx.TooManyRedirects):
        return "REDIRECT_LOOP", "Too many redirects"
    if isinstance(exc, httpx.ConnectError):
        text = str(exc)
        if "CERTIFICATE" in text.upper() or isinstance(exc.__cause__, ssl.SSLError):
            return "SSL_ERR", "SSL certificate problem: " + text[:150]
        if "getaddrinfo" in text or "Name or service not known" in text or "nodename" in text:
            return "DNS_ERR", "Domain could not be resolved"
        return "CONN_ERR", text[:200] or "Connection failed"
    return "ERROR", f"{type(exc).__name__}: {exc}"[:200]


def check_url(url: str, client: Optional[httpx.Client] = None, options: Optional[CheckOptions] = None) -> CheckResult:
    """GET the URL (body is not downloaded), following redirects; retries network failures."""
    opts = options or CheckOptions()
    own = client is None
    client = client or make_client(opts)
    try:
        attempts = 1 + max(0, opts.retries)
        for attempt in range(1, attempts + 1):
            start = time.perf_counter()
            try:
                with client.stream("GET", url) as response:
                    elapsed = (time.perf_counter() - start) * 1000
                    return CheckResult(
                        url=url,
                        status=response.status_code,
                        response_time_ms=round(elapsed, 1),
                        final_url=str(response.url),
                        content_type=response.headers.get("content-type", "").split(";")[0],
                        redirects=len(response.history),
                    )
            except (httpx.TimeoutException, httpx.ConnectError, httpx.RemoteProtocolError) as exc:
                status, message = _describe(exc)
                if attempt == attempts or status in ("DNS_ERR", "SSL_ERR"):
                    return CheckResult(url, status, -1, url, "", 0, message)
                time.sleep(0.5 * attempt)
            except Exception as exc:  # invalid URL, too many redirects, …
                status, message = _describe(exc)
                return CheckResult(url, status, -1, url, "", 0, message)
        return CheckResult(url, "ERROR", -1, url, "", 0, "Unknown failure")
    finally:
        if own:
            client.close()


def make_client(opts: CheckOptions) -> httpx.Client:
    return httpx.Client(
        follow_redirects=True,
        max_redirects=opts.max_redirects,
        timeout=opts.timeout,
        verify=opts.verify_ssl,
        headers={"User-Agent": USER_AGENT, "Accept": "*/*"},
        limits=httpx.Limits(max_connections=max(4, opts.workers * 2)),
    )


def check_many(
    urls: Iterable[str],
    options: Optional[CheckOptions] = None,
    on_result: Optional[Callable[[int, CheckResult], None]] = None,
    cancel: Optional[threading.Event] = None,
    client: Optional[httpx.Client] = None,
) -> list[CheckResult]:
    """Check URLs in parallel. `on_result(index, result)` is called from worker threads as results arrive."""
    opts = options or CheckOptions()
    url_list = list(urls)
    results: list[Optional[CheckResult]] = [None] * len(url_list)
    own = client is None
    client = client or make_client(opts)
    try:
        with ThreadPoolExecutor(max_workers=max(1, opts.workers)) as pool:
            futures = {}
            for i, url in enumerate(url_list):
                if cancel and cancel.is_set():
                    break
                futures[pool.submit(check_url, url, client, opts)] = i
            for fut in as_completed(futures):
                i = futures[fut]
                results[i] = fut.result()
                if on_result:
                    on_result(i, results[i])
                if cancel and cancel.is_set():
                    for f in futures:
                        f.cancel()
                    break
    finally:
        if own:
            client.close()
    return [r for r in results if r is not None]


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

HEADERS = ["URL", "Status", "Result", "Time (ms)", "Redirects", "Final URL", "Content-Type", "Error", "Checked At"]


def _row(r: CheckResult) -> list:
    return [r.url, r.status, r.category, r.response_time_ms if r.response_time_ms >= 0 else "", r.redirects,
            r.final_url, r.content_type, r.error, r.checked_at]


def export_csv(results: list[CheckResult], path: str | Path) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # BOM → opens correctly in Excel
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerows(_row(r) for r in results)


def export_xlsx(results: list[CheckResult], path: str | Path) -> None:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    fills = {
        "ok": "C6EFCE", "redirect": "DDEBF7", "client": "FFC7CE", "server": "FFC7CE",
        "timeout": "FFEB9C", "error": "FFC7CE",
    }
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in results:
        ws.append(_row(r))
        fill = PatternFill("solid", fgColor=fills.get(r.category, "FFFFFF"))
        for cell in ws[ws.max_row][1:3]:
            cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, width in enumerate([45, 10, 10, 10, 10, 45, 22, 40, 20], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    summary = wb.create_sheet("Summary")
    counts = summarize(results)
    summary.append(["Result", "Count"])
    for key, value in counts.items():
        summary.append([key, value])
    summary["A1"].font = summary["B1"].font = Font(bold=True)
    wb.save(path)


def summarize(results: list[CheckResult]) -> dict[str, int]:
    counts = {"total": len(results), "ok": 0, "redirect": 0, "client": 0, "server": 0, "timeout": 0, "error": 0}
    for r in results:
        counts[r.category] += 1
    return counts


def to_dict(r: CheckResult) -> dict:
    return {**asdict(r), "category": r.category}


# ---------------------------------------------------------------------------
# Command line (for scheduled checks without the GUI)
# ---------------------------------------------------------------------------

def main() -> None:
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Check a list of URLs (headless).")
    parser.add_argument("input", help=".txt / .csv / .xlsx file with URLs")
    parser.add_argument("-o", "--output", help="Write results to .xlsx or .csv")
    parser.add_argument("--timeout", type=float, default=10)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--retries", type=int, default=1)
    parser.add_argument("--insecure", action="store_true", help="Do not verify SSL certificates")
    args = parser.parse_args()

    urls, rejected = parse_url_list(read_url_file(args.input))
    for bad in rejected:
        print(f"skipped (not a URL): {bad}", file=sys.stderr)
    opts = CheckOptions(timeout=args.timeout, workers=args.workers, retries=args.retries, verify_ssl=not args.insecure)
    results = check_many(urls, opts, on_result=lambda i, r: print(f"{str(r.status):>13}  {r.response_time_ms:>8}  {r.url}"))
    if args.output:
        (export_xlsx if args.output.lower().endswith(".xlsx") else export_csv)(results, args.output)
        print(f"Saved {len(results)} results to {args.output}")
    s = summarize(results)
    print(f"OK {s['ok']} | redirect {s['redirect']} | 4xx {s['client']} | 5xx {s['server']} | timeout {s['timeout']} | error {s['error']}")
    sys.exit(0 if s["client"] + s["server"] + s["timeout"] + s["error"] == 0 else 1)


if __name__ == "__main__":
    main()
