"""
URL Health Checker — Desktop Tool
Cross-platform GUI app built with tkinter.
Checks a list of URLs in parallel for status, response time and redirects.
Exports results to Excel/CSV. Packaged as .EXE with PyInstaller.
"""
from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from checker import (
    APP_VERSION,
    CheckOptions,
    CheckResult,
    check_many,
    export_csv,
    export_xlsx,
    parse_url_list,
    read_url_file,
    summarize,
)

APP_TITLE = "URL Health Checker"

BG = "#1e1e2e"
SURFACE = "#313244"
OVERLAY = "#45475a"
TEXT = "#cdd6f4"
SUBTEXT = "#a6adc8"
COLORS = {"ok": "#a6e3a1", "redirect": "#89dceb", "client": "#f38ba8", "server": "#f38ba8", "timeout": "#f9e2af", "error": "#f38ba8"}
COLUMNS = ("URL", "Status", "Time (ms)", "Redirects", "Final URL", "Content-Type", "Error", "Checked At")


class URLCheckerApp(tk.Tk):
    """Main application window. Worker threads never touch widgets — they post to a queue."""

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1100x700")
        self.minsize(820, 520)
        self.configure(bg=BG)
        self.results: list[CheckResult] = []
        self._queue: queue.Queue = queue.Queue()
        self._cancel = threading.Event()
        self._running = False
        self._total = 0
        self._sort_state: dict[str, bool] = {}
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------ UI
    def _button(self, parent, text, command, bg, fg=BG) -> tk.Button:
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=bg,
                         font=("Segoe UI", 10), relief=tk.FLAT, padx=14, pady=6, cursor="hand2")

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=SURFACE, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text=f"🔍 {APP_TITLE}", bg=SURFACE, fg=TEXT, font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(header, text="Bulk URL status checker — exports to Excel/CSV", bg=SURFACE, fg=SUBTEXT,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5)

        # ── Input area
        input_frame = tk.Frame(self, bg=BG, padx=15, pady=10)
        input_frame.pack(fill=tk.X)
        tk.Label(input_frame, text="Enter URLs (one per line, commas also work; lines starting with # are ignored):",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 10)).pack(anchor=tk.W)
        self.url_text = tk.Text(input_frame, height=6, bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                                font=("Consolas", 10), relief=tk.FLAT, padx=8, pady=8, undo=True)
        self.url_text.pack(fill=tk.X, pady=(4, 0))
        self.url_text.insert("1.0", "https://example.com\nhttps://github.com\nhttps://httpbin.org/status/404")

        # ── Settings
        settings = tk.Frame(self, bg=BG, padx=15)
        settings.pack(fill=tk.X)
        self.timeout_var = tk.DoubleVar(value=10)
        self.workers_var = tk.IntVar(value=8)
        self.retries_var = tk.IntVar(value=1)
        self.verify_var = tk.BooleanVar(value=True)
        for label, var, frm, to in (("Timeout (s)", self.timeout_var, 2, 60), ("Parallel", self.workers_var, 1, 32),
                                    ("Retries", self.retries_var, 0, 3)):
            tk.Label(settings, text=label, bg=BG, fg=SUBTEXT, font=("Segoe UI", 9)).pack(side=tk.LEFT)
            tk.Spinbox(settings, from_=frm, to=to, textvariable=var, width=4, bg=SURFACE, fg=TEXT,
                       buttonbackground=OVERLAY, relief=tk.FLAT, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(4, 14))
        tk.Checkbutton(settings, text="Verify SSL certificates", variable=self.verify_var, bg=BG, fg=SUBTEXT,
                       selectcolor=SURFACE, activebackground=BG, activeforeground=TEXT,
                       font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # ── Controls
        controls = tk.Frame(self, bg=BG, padx=15, pady=8)
        controls.pack(fill=tk.X)
        self.check_btn = self._button(controls, "▶  Check URLs", self._start_check, "#89b4fa")
        self.check_btn.config(font=("Segoe UI", 10, "bold"))
        self.check_btn.pack(side=tk.LEFT)
        self.stop_btn = self._button(controls, "■ Stop", self._stop, "#f38ba8")
        self.stop_btn.pack(side=tk.LEFT, padx=(8, 0))
        self.stop_btn.config(state=tk.DISABLED)
        self._button(controls, "📂 Load file", self._load_file, OVERLAY, TEXT).pack(side=tk.LEFT, padx=(8, 0))
        self._button(controls, "💾 Export Excel", lambda: self._export("xlsx"), "#a6e3a1").pack(side=tk.LEFT, padx=(8, 0))
        self._button(controls, "📄 Export CSV", lambda: self._export("csv"), "#89dceb").pack(side=tk.LEFT, padx=(8, 0))
        self.status_label = tk.Label(controls, text="Ready", bg=BG, fg=SUBTEXT, font=("Segoe UI", 9))
        self.status_label.pack(side=tk.RIGHT, padx=10)

        # ── Progress + summary
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill=tk.X, padx=15)
        self.summary_label = tk.Label(self, text="", bg=BG, fg=SUBTEXT, font=("Segoe UI", 9), anchor=tk.W)
        self.summary_label.pack(fill=tk.X, padx=15, pady=(6, 0))

        # ── Results table
        table_frame = tk.Frame(self, bg=BG, padx=15, pady=10)
        table_frame.pack(fill=tk.BOTH, expand=True)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, foreground=TEXT, fieldbackground=SURFACE, rowheight=26)
        style.configure("Treeview.Heading", background=OVERLAY, foreground=TEXT, font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#585b70")])

        self.tree = ttk.Treeview(table_frame, columns=COLUMNS, show="headings", height=16)
        widths = (240, 90, 80, 75, 220, 130, 200, 140)
        for col, width in zip(COLUMNS, widths):
            self.tree.heading(col, text=col, command=lambda c=col: self._sort_by(c))
            anchor = tk.CENTER if col in ("Status", "Time (ms)", "Redirects") else tk.W
            self.tree.column(col, width=width, anchor=anchor)
        yscroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        for category, color in COLORS.items():
            self.tree.tag_configure(category, foreground=color)
        self.tree.bind("<Double-1>", self._open_selected)

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open in browser", command=self._open_selected)
        menu.add_command(label="Copy URL", command=self._copy_selected)
        self.tree.bind("<Button-3>", lambda e: (self.tree.selection_set(self.tree.identify_row(e.y)), menu.tk_popup(e.x_root, e.y_root)))

    # ------------------------------------------------------------ actions
    def _check_options(self) -> CheckOptions:
        def safe(var, default, lo, hi):
            try:
                return min(max(type(default)(var.get()), lo), hi)
            except (tk.TclError, ValueError):
                return default
        return CheckOptions(timeout=safe(self.timeout_var, 10.0, 2, 60), workers=safe(self.workers_var, 8, 1, 32),
                            retries=safe(self.retries_var, 1, 0, 3), verify_ssl=bool(self.verify_var.get()))

    def _start_check(self) -> None:
        urls, rejected = parse_url_list(self.url_text.get("1.0", tk.END))
        if rejected:
            preview = "\n".join(rejected[:10]) + ("\n…" if len(rejected) > 10 else "")
            messagebox.showwarning("Skipped lines", f"{len(rejected)} line(s) are not valid URLs and will be skipped:\n\n{preview}")
        if not urls:
            messagebox.showwarning("No URLs", "Please enter at least one valid URL.")
            return

        self.results.clear()
        self.tree.delete(*self.tree.get_children())
        self._total = len(urls)
        self.progress.config(maximum=len(urls), value=0)
        self.summary_label.config(text="")
        self._cancel.clear()
        self._set_running(True)
        opts = self._check_options()
        threading.Thread(target=self._worker, args=(urls, opts), daemon=True).start()

    def _worker(self, urls: list[str], opts: CheckOptions) -> None:
        try:
            check_many(urls, opts, on_result=lambda i, r: self._queue.put(("result", r)), cancel=self._cancel)
        except Exception as exc:  # never leave the UI stuck in "running"
            self._queue.put(("error", str(exc)))
        self._queue.put(("done", None))

    def _stop(self) -> None:
        self._cancel.set()
        self.status_label.config(text="Stopping…")

    def _drain_queue(self) -> None:
        """Runs on the Tk thread: apply everything the worker produced."""
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "result":
                    self._add_row(payload)
                elif kind == "error":
                    messagebox.showerror("Error", payload)
                elif kind == "done":
                    self._on_done()
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _add_row(self, r: CheckResult) -> None:
        self.results.append(r)
        self.tree.insert("", tk.END, values=(r.url, r.status, r.response_time_ms if r.response_time_ms >= 0 else "—",
                                              r.redirects, r.final_url, r.content_type, r.error, r.checked_at),
                         tags=(r.category,))
        self.progress.config(value=len(self.results))
        self.status_label.config(text=f"Checked {len(self.results)}/{self._total}")

    def _on_done(self) -> None:
        self._set_running(False)
        s = summarize(self.results)
        stopped = " (stopped)" if self._cancel.is_set() else ""
        self.status_label.config(text=f"Done{stopped} — {s['ok'] + s['redirect']}/{s['total']} reachable")
        self.summary_label.config(
            text=f"✔ OK {s['ok']}   ↪ Redirect {s['redirect']}   ✖ 4xx {s['client']}   ✖ 5xx {s['server']}   "
                 f"⏱ Timeout {s['timeout']}   ⚠ Error {s['error']}"
        )

    def _set_running(self, running: bool) -> None:
        self._running = running
        self.check_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL if running else tk.DISABLED)

    def _sort_by(self, col: str) -> None:
        desc = self._sort_state[col] = not self._sort_state.get(col, False)

        def key(item):
            value = self.tree.set(item, col)
            try:
                return (0, float(value))
            except ValueError:
                return (1, value.lower())

        items = sorted(self.tree.get_children(""), key=key, reverse=desc)
        for index, item in enumerate(items):
            self.tree.move(item, "", index)
        for c in COLUMNS:
            self.tree.heading(c, text=c + ((" ▼" if desc else " ▲") if c == col else ""))

    def _selected_url(self) -> str | None:
        sel = self.tree.selection()
        return self.tree.set(sel[0], "URL") if sel else None

    def _open_selected(self, _event=None) -> None:
        url = self._selected_url()
        if url:
            webbrowser.open(url)

    def _copy_selected(self) -> None:
        url = self._selected_url()
        if url:
            self.clipboard_clear()
            self.clipboard_append(url)

    def _load_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("URL lists", "*.txt *.csv *.xlsx"), ("All files", "*.*")])
        if not path:
            return
        try:
            text = read_url_file(path)
        except Exception as exc:
            messagebox.showerror("Could not read file", str(exc))
            return
        urls, rejected = parse_url_list(text)
        self.url_text.delete("1.0", tk.END)
        self.url_text.insert("1.0", "\n".join(urls))
        self.status_label.config(text=f"Loaded {len(urls)} URL(s)" + (f", skipped {len(rejected)}" if rejected else ""))

    def _export(self, fmt: str) -> None:
        if not self.results:
            messagebox.showinfo("No data", "Run a check first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=f".{fmt}",
            filetypes=[(f"{fmt.upper()} files", f"*.{fmt}")],
            initialfile=f"url_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        )
        if not path:
            return
        try:
            (export_xlsx if fmt == "xlsx" else export_csv)(self.results, path)
        except PermissionError:
            messagebox.showerror("Export failed", "The file is open in another program (Excel?). Close it and try again.")
            return
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            return
        messagebox.showinfo("Exported", f"Results saved to:\n{path}")

    def _on_close(self) -> None:
        if self._running and not messagebox.askyesno("Quit", "A check is still running. Quit anyway?"):
            return
        self._cancel.set()
        self.destroy()


def main() -> None:
    if sys.platform == "win32":
        try:  # crisp text on high-DPI screens
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    URLCheckerApp().mainloop()


if __name__ == "__main__":
    main()
