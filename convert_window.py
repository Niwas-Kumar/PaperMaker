"""PDF → editable Word window for Paper Maker.

Runs ``tools/pdf_math_to_docx.py`` once per PDF in a background process, so the
window never freezes. In the shared (frozen) build the converter runs inside the
portable engine that sits next to PaperMaker.exe::

    PaperMaker.exe
    engine/python.exe, engine/tools/pdf_math_to_docx.py   (MinerU, CPU, no install)
    mineru_home/                                           (its models, offline)

From source it runs with the current Python (MinerU found as in the script).
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

SECONDS_PER_PAGE = 25  # CPU estimate for the progress hint (≈12 s with an NVIDIA GPU)
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def app_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent


def cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "PaperMaker" / "engine_cache"


def engine_command() -> tuple[list[str], dict] | None:
    """(command prefix, environment) that runs the converter, or None if missing."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    if getattr(sys, "frozen", False):
        base = app_dir()
        py, script = base / "engine" / "python.exe", base / "engine" / "tools" / "pdf_math_to_docx.py"
        if not (py.exists() and script.exists()):
            return None
        for k in ("PYTHONPATH", "PYTHONHOME"):
            env.pop(k, None)
        env.update(MINERU_HOME=str(base / "mineru_home"), HF_HUB_OFFLINE="1",
                   MODELSCOPE_OFFLINE="1")
        # -E ignores PYTHONIOENCODING; -X utf8 keeps the output UTF-8 anyway
        return [str(py), "-E", "-s", "-X", "utf8", str(script)], env
    script = app_dir() / "tools" / "pdf_math_to_docx.py"
    return ([sys.executable, str(script)], env) if script.exists() else None


def page_count(pdf: Path) -> int | None:
    try:
        import fitz  # PyMuPDF is part of Paper Maker already
        with fitz.open(pdf) as d:
            return d.page_count
    except Exception:  # noqa: BLE001
        return None


class ConvertWindow:
    """Toplevel: list of PDFs → <name>_editable.docx each, in the house style."""

    def __init__(self, app) -> None:
        self.app = app
        self.win = tk.Toplevel(app.root)
        self.win.title("PDF → Editable Word")
        self.win.geometry("900x620")
        self.win.minsize(760, 520)
        self.win.configure(bg=app.root["bg"])
        self.items: list[dict] = []          # {"pdf", "pages", "status", "out", "iid"}
        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.proc: subprocess.Popen | None = None
        self.running = False
        self.cancelled = False
        self.out_dir: Path | None = None
        self.template: Path | None = None
        self._build()
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── layout ────────────────────────────────────────────────────────────
    def _build(self) -> None:
        a, w = self.app, self.win
        bg = a.root["bg"]
        head = ttk.Frame(w, style="Dark.TFrame")
        head.pack(fill="x", padx=20, pady=(16, 4))
        ttk.Label(head, text="🧮  PDF → Editable Word", style="Title.TLabel").pack(side="left")
        ttk.Label(w, style="DarkText.TLabel", justify="left", text=(
            "Turns question-paper PDFs into Word files with real, editable equations. "
            "If the paper's original .docx sits next to the PDF (same name), the output "
            "copies its exact format.")).pack(fill="x", padx=20, pady=(0, 8))

        card = ttk.Frame(w, style="Card.TFrame")
        card.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", padx=14, pady=(12, 6))
        ttk.Label(row, text="PDF files", style="Heading.TLabel").pack(side="left")
        ttk.Button(row, text="✕ Remove", style="Danger.TButton",
                   command=self._on_remove).pack(side="right")
        ttk.Button(row, text="➕ Add PDFs", style="Accent.TButton",
                   command=self._on_add).pack(side="right", padx=(0, 8))

        wrap = tk.Frame(card, bg="#252836", highlightthickness=1, highlightbackground="#2d3148")
        wrap.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        self.tree = ttk.Treeview(wrap, style="Files.Treeview", height=6,
                                 columns=("pages", "format", "status"), show="tree headings")
        self.tree.heading("#0", text="File")
        self.tree.heading("pages", text="Pages")
        self.tree.heading("format", text="Format")
        self.tree.heading("status", text="Status")
        self.tree.column("#0", width=380, anchor="w")
        self.tree.column("pages", width=60, anchor="center")
        self.tree.column("format", width=170, anchor="center")
        self.tree.column("status", width=150, anchor="center")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-Button-1>", self._on_open_output)

        opts = ttk.Frame(card, style="Card.TFrame")
        opts.pack(fill="x", padx=14, pady=(0, 12))
        ttk.Label(opts, text="Format", style="CardText.TLabel").grid(row=0, column=0, sticky="w")
        self.fmt_var = tk.StringVar(value="auto")
        for i, (val, text) in enumerate((("auto", "Original .docx next to the PDF (if any)"),
                                         ("choose", "One Word file for all…"),
                                         ("plain", "Plain layout"))):
            tk.Radiobutton(opts, text=text, value=val, variable=self.fmt_var,
                           command=self._on_format, bg="#1a1d28", fg="#e4e6eb",
                           selectcolor="#252836", activebackground="#1a1d28",
                           activeforeground="#e4e6eb", font=a.font_normal,
                           highlightthickness=0).grid(row=0, column=1 + i, sticky="w", padx=(10, 0))
        self.tpl_var = tk.StringVar(value="")
        ttk.Label(opts, textvariable=self.tpl_var, style="CardText.TLabel").grid(
            row=1, column=1, columnspan=3, sticky="w", padx=(10, 0))
        # line drawings as vector art (SVG): sharp at any zoom; in Word 365 a right-click
        # > "Convert to Shape" turns one into editable lines
        self.vector_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="Sharp vector diagrams (Word 2019 / 365)",
                       variable=self.vector_var, bg="#1a1d28", fg="#e4e6eb",
                       selectcolor="#252836", activebackground="#1a1d28",
                       activeforeground="#e4e6eb", font=a.font_normal,
                       highlightthickness=0).grid(row=2, column=1, columnspan=3,
                                                  sticky="w", padx=(10, 0))

        bottom = ttk.Frame(w, style="Dark.TFrame")
        bottom.pack(fill="x", side="bottom", padx=20, pady=(0, 14))
        out_row = ttk.Frame(bottom, style="Dark.TFrame")
        out_row.pack(fill="x", pady=(0, 8))
        ttk.Label(out_row, text="Save to", style="DarkText.TLabel").pack(side="left", padx=(0, 8))
        self.out_var = tk.StringVar(value="Same folder as each PDF")
        tk.Entry(out_row, textvariable=self.out_var, state="readonly", font=a.font_normal,
                 bg="#252836", fg="#9ca3af", readonlybackground="#252836", relief="flat",
                 highlightthickness=1, highlightbackground="#2d3148").pack(
            side="left", fill="x", expand=True, ipady=5)
        ttk.Button(out_row, text="📂 Browse", style="Ghost.TButton",
                   command=self._on_out).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Add one or more PDF files")
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").pack(
            fill="x", pady=(0, 4))
        self.bar = ttk.Progressbar(bottom, maximum=100, style="Custom.Horizontal.TProgressbar")
        self.bar.pack(fill="x", pady=(0, 8))
        btns = ttk.Frame(bottom, style="Dark.TFrame")
        btns.pack(fill="x")
        self.cancel_btn = ttk.Button(btns, text="Cancel", style="Danger.TButton",
                                     command=self._on_cancel)
        self.cancel_btn.pack(side="right")
        self.cancel_btn.state(["disabled"])
        self.go_btn = ttk.Button(btns, text="⬇  Convert to Word", style="Success.TButton",
                                 command=self._on_convert)
        self.go_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

        logwrap = tk.Frame(w, bg=bg)
        logwrap.pack(fill="x", padx=20, pady=(0, 8), side="bottom")
        self.log = tk.Text(logwrap, height=6, bg="#151821", fg="#9ca3af", relief="flat",
                           font=("Consolas", 9), wrap="none", state="disabled")
        self.log.pack(fill="x")

    # ── list ──────────────────────────────────────────────────────────────
    def _format_label(self, pdf: Path) -> str:
        mode = self.fmt_var.get()
        if mode == "plain":
            return "plain"
        if mode == "choose":
            return self.template.name[:24] if self.template else "(choose a file)"
        return "original .docx ✓" if pdf.with_suffix(".docx").exists() else "plain (no .docx)"

    def _refresh(self) -> None:
        for it in self.items:
            pages = "?" if it["pages"] is None else str(it["pages"])
            self.tree.item(it["iid"], values=(pages, self._format_label(it["pdf"]), it["status"]))

    def _on_add(self) -> None:
        if self.running:
            return
        paths = filedialog.askopenfilenames(parent=self.win, title="Select PDF papers",
                                            filetypes=[("PDF files", "*.pdf")])
        known = {it["pdf"] for it in self.items}
        for p in map(Path, paths):
            if p in known:
                continue
            iid = self.tree.insert("", "end", text="  " + p.name)
            self.items.append({"pdf": p, "pages": page_count(p), "status": "waiting",
                               "out": None, "iid": iid})
        self._refresh()
        self._update_status()

    def _on_remove(self) -> None:
        if self.running:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.items = [it for it in self.items if it["iid"] != iid]
        self._update_status()

    def _on_format(self) -> None:
        if self.fmt_var.get() == "choose":
            p = filedialog.askopenfilename(parent=self.win, title="Word file whose format to copy",
                                           filetypes=[("Word documents", "*.docx")])
            if p:
                self.template = Path(p)
                self.tpl_var.set(f"Format from: {self.template.name}")
            elif self.template is None:
                self.fmt_var.set("auto")
        else:
            self.tpl_var.set("")
        self._refresh()

    def _on_out(self) -> None:
        d = filedialog.askdirectory(parent=self.win, title="Save converted files to")
        if d:
            self.out_dir = Path(d)
            self.out_var.set(d)

    def _on_open_output(self, _e=None) -> None:
        sel = self.tree.selection()
        it = next((i for i in self.items if sel and i["iid"] == sel[0]), None)
        if it and it["out"] and Path(it["out"]).exists():
            os.startfile(it["out"])  # noqa: S606 - opens the user's own result in Word

    def _update_status(self) -> None:
        if not self.running:
            n = len(self.items)
            pages = sum(it["pages"] or 0 for it in self.items)
            self.status_var.set(
                f"{n} PDF(s), {pages} page(s) · about {max(1, round(pages * SECONDS_PER_PAGE / 60))}"
                " min on a normal PC" if n else "Add one or more PDF files")

    def _write_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── run ───────────────────────────────────────────────────────────────
    def _on_convert(self) -> None:
        if self.running:
            return
        todo = [it for it in self.items if it["status"] != "✓ done"]
        if not todo:
            messagebox.showinfo("Nothing to convert", "Add PDF files first.", parent=self.win)
            return
        eng = engine_command()
        if eng is None:
            messagebox.showerror("Converter missing",
                                 "The 'engine' folder was not found next to PaperMaker.exe.\n"
                                 "Unzip the complete PaperMaker folder and run it from there.",
                                 parent=self.win)
            return
        if self.fmt_var.get() == "choose" and self.template is None:
            messagebox.showwarning("Choose a Word file", "Pick the Word file whose format "
                                   "to copy, or choose another format option.", parent=self.win)
            return
        jobs = []
        for it in todo:
            pdf = it["pdf"].resolve()
            out_dir = (self.out_dir or pdf.parent).resolve()
            out = out_dir / f"{pdf.stem}_editable.docx"
            args = [str(pdf), "-o", str(out), "--cache-dir", str(cache_dir())]
            if self.fmt_var.get() == "plain":
                args.append("--no-template")
            elif self.fmt_var.get() == "choose":
                args += ["--template", str(self.template.resolve())]
            if self.vector_var.get():
                args.append("--vector-figures")

            jobs.append((it, out, args))
        self.running, self.cancelled = True, False
        self.go_btn.state(["disabled"])
        self.cancel_btn.state(["!disabled"])
        threading.Thread(target=self._worker, args=(eng, jobs), daemon=True).start()
        self.win.after(150, self._poll)

    def _worker(self, eng, jobs) -> None:
        cmd, env = eng
        total_pages = sum(it["pages"] or 10 for it, _, _ in jobs)
        done_pages = 0
        for k, (it, out, args) in enumerate(jobs, 1):
            if self.cancelled:
                break
            self.events.put(("status", it, "converting…"))
            self.events.put(("log", f"── {it['pdf'].name}"))
            start = time.time()
            try:
                self.proc = subprocess.Popen(cmd + args, env=env, stdout=subprocess.PIPE,
                                             stderr=subprocess.STDOUT, creationflags=NO_WINDOW,
                                             cwd=str(it["pdf"].parent))
            except OSError as exc:
                self.events.put(("status", it, "⚠ failed"))
                self.events.put(("log", f"could not start the converter: {exc}"))
                continue
            buf, tail = b"", []
            while True:
                chunk = self.proc.stdout.read1(4096) if hasattr(self.proc.stdout, "read1") \
                    else self.proc.stdout.read(4096)
                if not chunk:
                    break
                buf += chunk
                *lines, buf = re.split(rb"[\r\n]", buf)
                for raw in lines:
                    line = raw.decode("utf-8", "replace").strip()
                    if not line:
                        continue
                    tail = (tail + [line])[-15:]
                    if line.startswith("[") and not line.startswith("[3"):  # our own steps
                        self.events.put(("log", line))
                    elapsed = time.time() - start
                    est = (it["pages"] or 10) * SECONDS_PER_PAGE
                    frac = (done_pages + min(0.95, elapsed / est) * (it["pages"] or 10)) / total_pages
                    self.events.put(("progress", frac * 100,
                                     f"File {k} of {len(jobs)} · {it['pdf'].name} · "
                                     f"{int(elapsed // 60)}:{int(elapsed % 60):02d} elapsed"))
            rc = self.proc.wait()
            done_pages += it["pages"] or 10
            if self.cancelled:
                self.events.put(("status", it, "cancelled"))
            elif rc == 0 and out.exists():
                it["out"] = str(out)
                self.events.put(("status", it, "✓ done"))
                self.events.put(("log", f"saved {out}  ({time.time() - start:.0f} s)"))
            else:
                self.events.put(("status", it, "⚠ failed"))
                self.events.put(("log", "\n".join(tail[-8:])))
        self.proc = None
        self.events.put(("finished",))

    def _poll(self) -> None:
        try:
            while True:
                ev = self.events.get_nowait()
                if ev[0] == "status":
                    ev[1]["status"] = ev[2]
                    self._refresh()
                elif ev[0] == "log":
                    self._write_log(ev[1])
                elif ev[0] == "progress":
                    self.bar["value"] = ev[1]
                    self.status_var.set(ev[2])
                elif ev[0] == "finished":
                    self._finished()
                    return
        except queue.Empty:
            pass
        self.win.after(150, self._poll)

    def _finished(self) -> None:
        self.running = False
        self.go_btn.state(["!disabled"])
        self.cancel_btn.state(["disabled"])
        ok = [it for it in self.items if it["status"] == "✓ done"]
        bad = [it for it in self.items if it["status"] == "⚠ failed"]
        self.bar["value"] = 100 if ok and not bad else self.bar["value"]
        self.status_var.set(f"Done · {len(ok)} converted" + (f", {len(bad)} failed" if bad else "")
                            + " · double-click a file to open it")
        if not self.cancelled:
            msg = f"{len(ok)} Word file(s) saved."
            if bad:
                msg += "\n\nFailed:\n" + "\n".join(f"• {it['pdf'].name}" for it in bad) + \
                    "\n\nThe log at the bottom of the window shows why."
            msg += "\n\nPlease proofread the formulas: OCR can misread one now and then."
            (messagebox.showwarning if bad else messagebox.showinfo)(
                "Conversion finished", msg, parent=self.win)

    def _on_cancel(self) -> None:
        if not self.running:
            return
        self.cancelled = True
        p = self.proc
        if p and p.poll() is None:  # the converter and the MinerU process it started
            subprocess.run(["taskkill", "/PID", str(p.pid), "/T", "/F"],
                           capture_output=True, creationflags=NO_WINDOW)
        self.status_var.set("Cancelling…")

    def _on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("Stop conversion?",
                                       "A conversion is running. Stop it and close?",
                                       parent=self.win):
                return
            self._on_cancel()
        self.win.destroy()
        self.app._convert_window = None
