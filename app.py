"""
URL Health Checker — Desktop Tool
Cross-platform GUI app built with tkinter.
Checks a list of URLs for status, response time, and redirects.
Exports results to Excel/CSV. Packaged as .EXE with PyInstaller.
"""
from __future__ import annotations

import csv
import os
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List

import httpx
import openpyxl

APP_TITLE = "URL Health Checker"
APP_VERSION = "1.0.0"
TIMEOUT = 10
MAX_REDIRECTS = 5


class CheckResult:
    """Holds the result of a single URL health check."""

    def __init__(
        self,
        url: str,
        status: int | str,
        response_time_ms: float,
        final_url: str,
        content_type: str,
        error: str = "",
    ) -> None:
        self.url = url
        self.status = status
        self.response_time_ms = response_time_ms
        self.final_url = final_url
        self.content_type = content_type
        self.error = error
        self.checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @property
    def is_ok(self) -> bool:
        return isinstance(self.status, int) and 200 <= self.status < 400


def check_url(url: str) -> CheckResult:
    """Perform an HTTP GET request and return a CheckResult."""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    start = time.perf_counter()
    try:
        with httpx.Client(
            follow_redirects=True,
            max_redirects=MAX_REDIRECTS,
            timeout=TIMEOUT,
            headers={"User-Agent": f"URLHealthChecker/{APP_VERSION}"},
        ) as client:
            response = client.get(url)
            elapsed = (time.perf_counter() - start) * 1000
            return CheckResult(
                url=url,
                status=response.status_code,
                response_time_ms=round(elapsed, 1),
                final_url=str(response.url),
                content_type=response.headers.get("content-type", ""),
                error="",
            )
    except httpx.TimeoutException:
        return CheckResult(url, "TIMEOUT", -1, url, "", "Request timed out")
    except httpx.ConnectError as e:
        return CheckResult(url, "CONN_ERR", -1, url, "", str(e))
    except Exception as e:
        return CheckResult(url, "ERROR", -1, url, "", str(e))


class URLCheckerApp(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("900x640")
        self.resizable(True, True)
        self.configure(bg="#1e1e2e")
        self.results: List[CheckResult] = []
        self._build_ui()

    def _build_ui(self) -> None:
        """Build the complete UI layout."""
        # ── Header ──────────────────────────────────────────────
        header = tk.Frame(self, bg="#313244", pady=12)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text=f"🔍 {APP_TITLE}",
            bg="#313244",
            fg="#cdd6f4",
            font=("Segoe UI", 16, "bold"),
        ).pack(side=tk.LEFT, padx=20)
        tk.Label(
            header,
            text="Bulk URL status checker — exports to Excel/CSV",
            bg="#313244",
            fg="#a6adc8",
            font=("Segoe UI", 10),
        ).pack(side=tk.LEFT, padx=5)

        # ── Input area ──────────────────────────────────────────
        input_frame = tk.Frame(self, bg="#1e1e2e", padx=15, pady=10)
        input_frame.pack(fill=tk.X)

        tk.Label(
            input_frame,
            text="Enter URLs (one per line):",
            bg="#1e1e2e",
            fg="#a6adc8",
            font=("Segoe UI", 10),
        ).pack(anchor=tk.W)

        self.url_text = tk.Text(
            input_frame,
            height=6,
            bg="#313244",
            fg="#cdd6f4",
            insertbackground="#cdd6f4",
            font=("Consolas", 10),
            relief=tk.FLAT,
            padx=8,
            pady=8,
        )
        self.url_text.pack(fill=tk.X, pady=(4, 0))
        self.url_text.insert("1.0", "https://example.com\nhttps://github.com\nhttps://httpbin.org/status/404")

        # ── Controls ────────────────────────────────────────────
        controls = tk.Frame(self, bg="#1e1e2e", padx=15, pady=8)
        controls.pack(fill=tk.X)

        self.check_btn = tk.Button(
            controls,
            text="▶  Check URLs",
            command=self._start_check,
            bg="#89b4fa",
            fg="#1e1e2e",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT,
            padx=18,
            pady=6,
            cursor="hand2",
        )
        self.check_btn.pack(side=tk.LEFT)

        tk.Button(
            controls,
            text="📂 Load from file",
            command=self._load_file,
            bg="#45475a",
            fg="#cdd6f4",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(8, 0))

        tk.Button(
            controls,
            text="💾 Export Excel",
            command=lambda: self._export("xlsx"),
            bg="#a6e3a1",
            fg="#1e1e2e",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(8, 0))

        tk.Button(
            controls,
            text="📄 Export CSV",
            command=lambda: self._export("csv"),
            bg="#89dceb",
            fg="#1e1e2e",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(8, 0))

        self.status_label = tk.Label(
            controls,
            text="Ready",
            bg="#1e1e2e",
            fg="#a6adc8",
            font=("Segoe UI", 9),
        )
        self.status_label.pack(side=tk.RIGHT, padx=10)

        # ── Progress bar ────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill=tk.X, padx=15)

        # ── Results table ────────────────────────────────────────
        table_frame = tk.Frame(self, bg="#1e1e2e", padx=15, pady=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("URL", "Status", "Time (ms)", "Final URL", "Content-Type", "Checked At")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=16)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#313244", foreground="#cdd6f4",
                         fieldbackground="#313244", rowheight=26)
        style.configure("Treeview.Heading", background="#45475a", foreground="#cdd6f4",
                         font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#585b70")])

        for col in cols:
            self.tree.heading(col, text=col)
        self.tree.column("URL", width=220)
        self.tree.column("Status", width=80, anchor=tk.CENTER)
        self.tree.column("Time (ms)", width=90, anchor=tk.CENTER)
        self.tree.column("Final URL", width=200)
        self.tree.column("Content-Type", width=150)
        self.tree.column("Checked At", width=140)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("ok", foreground="#a6e3a1")
        self.tree.tag_configure("warn", foreground="#f9e2af")
        self.tree.tag_configure("err", foreground="#f38ba8")

    def _get_urls(self) -> List[str]:
        raw = self.url_text.get("1.0", tk.END)
        return [line.strip() for line in raw.splitlines() if line.strip()]

    def _start_check(self) -> None:
        urls = self._get_urls()
        if not urls:
            messagebox.showwarning("No URLs", "Please enter at least one URL.")
            return

        self.results.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)

        self.check_btn.config(state=tk.DISABLED)
        self.progress["maximum"] = len(urls)
        self.progress["value"] = 0
        threading.Thread(target=self._run_checks, args=(urls,), daemon=True).start()

    def _run_checks(self, urls: List[str]) -> None:
        for i, url in enumerate(urls, 1):
            self.status_label.config(text=f"Checking {i}/{len(urls)}: {url[:60]}")
            result = check_url(url)
            self.results.append(result)
            self.after(0, self._add_row, result)
            self.progress["value"] = i

        self.after(0, self._on_done, len(urls))

    def _add_row(self, r: CheckResult) -> None:
        tag = "ok" if r.is_ok else ("warn" if r.status == "TIMEOUT" else "err")
        self.tree.insert(
            "",
            tk.END,
            values=(r.url, r.status, r.response_time_ms, r.final_url, r.content_type, r.checked_at),
            tags=(tag,),
        )

    def _on_done(self, total: int) -> None:
        ok_count = sum(1 for r in self.results if r.is_ok)
        self.status_label.config(text=f"Done — {ok_count}/{total} OK")
        self.check_btn.config(state=tk.NORMAL)

    def _load_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.url_text.delete("1.0", tk.END)
            self.url_text.insert("1.0", Path(path).read_text(encoding="utf-8").strip())

    def _export(self, fmt: str) -> None:
        if not self.results:
            messagebox.showinfo("No data", "Run a check first.")
            return

        ext = "xlsx" if fmt == "xlsx" else "csv"
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=f"url_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        )
        if not path:
            return

        headers = ["URL", "Status", "Time (ms)", "Final URL", "Content-Type", "Error", "Checked At"]
        rows = [
            [r.url, r.status, r.response_time_ms, r.final_url, r.content_type, r.error, r.checked_at]
            for r in self.results
        ]

        if fmt == "xlsx":
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(headers)
            for row in rows:
                ws.append(row)
            wb.save(path)
        else:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(rows)

        messagebox.showinfo("Exported", f"Results saved to:\n{path}")


if __name__ == "__main__":
    app = URLCheckerApp()
    app.mainloop()
