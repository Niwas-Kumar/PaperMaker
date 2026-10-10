#!/usr/bin/env python3
"""Paper Maker — build custom question papers from DOCX source files.

A dark-themed desktop app that parses Word (.docx) papers, lets you pick and
reorder questions per file, and generates trimmed papers that keep 100% of the
original formatting — two-column layout, header on every page, footer, images
and equations.

Work on several files at once: each source file produces its own output paper,
and all of them are generated in parallel. Parsing and generation run off the UI
thread, so the window never freezes (even during the first-open malware scan of a
large file).

The document engine lives in :mod:`paper_core`.
"""

from __future__ import annotations

import json
import multiprocessing
import queue
import re
import sys
import threading
import tkinter as tk
from collections import OrderedDict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from paper_core import (
    MODE_ANSWERS,
    MODE_BOTH,
    MODE_QUESTIONS,
    BatchJob,
    QuestionInfo,
    generate_papers_batch,
    group_by_subject,
    parse_paper,
)

# Output selector labels → BatchJob modes (order = combobox order).
MODE_LABELS: "OrderedDict[str, str]" = OrderedDict([
    ("Questions only", MODE_QUESTIONS),
    ("Answer key only", MODE_ANSWERS),
    ("Questions + Answer key", MODE_BOTH),
])
LABEL_FOR_MODE = {v: k for k, v in MODE_LABELS.items()}

# ═══════════════════════════════════════════════════════════════════════════════
# THEME
# ═══════════════════════════════════════════════════════════════════════════════

BG_DARK = "#0f1117"
BG_CARD = "#1a1d28"
BG_INPUT = "#252836"
BG_ROW_ALT = "#1f2230"
FG_PRIMARY = "#e4e6eb"
FG_SECONDARY = "#9ca3af"
FG_MUTED = "#6b7280"
ACCENT = "#6c63ff"
ACCENT_HOVER = "#7f78ff"
ACCENT_PRESSED = "#5a52e0"
SUCCESS = "#22c55e"
SUCCESS_HOVER = "#16a34a"
DANGER = "#ef4444"
DANGER_HOVER = "#dc2626"
BORDER = "#2d3148"
WARN = "#f59e0b"

CONFIG_PATH = Path.home() / ".paper_maker.json"
# bundled next to the code by PaperMaker.spec (datas)
ICON_PATH = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "assets" / "paper_maker.ico"


# ═══════════════════════════════════════════════════════════════════════════════
# PER-FILE STATE
# ═══════════════════════════════════════════════════════════════════════════════


class FileState:
    """Everything the app tracks about one loaded source file."""

    __slots__ = ("path", "subjects", "export", "status", "error",
                 "has_answer_key", "mode")

    def __init__(self, path: Path) -> None:
        self.path = path
        self.subjects: OrderedDict[str, list[QuestionInfo]] = OrderedDict()
        self.export: list[QuestionInfo] = []
        self.status = "parsing"  # parsing | ready | error
        self.error: str | None = None
        self.has_answer_key = False
        self.mode = MODE_QUESTIONS  # what to write; see MODE_LABELS

    @property
    def total_questions(self) -> int:
        return sum(len(v) for v in self.subjects.values())


# ═══════════════════════════════════════════════════════════════════════════════
# APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════


class PaperMakerApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Paper Maker")
        try:  # default= also gives every later window (sign-in, PDF → Word) the icon
            self.root.iconbitmap(default=str(ICON_PATH))
        except tk.TclError:
            pass  # icon file missing: keep Tk's default
        self.root.geometry("1120x760")
        self.root.minsize(960, 600)
        self.root.configure(bg=BG_DARK)

        # State
        self.files: "OrderedDict[Path, FileState]" = OrderedDict()
        self.current: Path | None = None
        self.output_dir: Path | None = None
        self._events: "queue.Queue[tuple]" = queue.Queue()
        self._workers: list[threading.Thread] = []
        self._generating = False
        self._tree_ids: dict[str, Path] = {}

        # Fonts
        self.font_title = ("Segoe UI Semibold", 17)
        self.font_heading = ("Segoe UI Semibold", 11)
        self.font_normal = ("Segoe UI", 10)
        self.font_small = ("Segoe UI", 9)

        self._config = self._load_config()
        self._configure_styles()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Config persistence ────────────────────────────────────────────────
    def _load_config(self) -> dict:
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_config(self) -> None:
        try:
            CONFIG_PATH.write_text(json.dumps(self._config), encoding="utf-8")
        except Exception:
            pass

    # ── Styles ────────────────────────────────────────────────────────────
    def _configure_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Dark.TFrame", background=BG_DARK)
        style.configure("Card.TFrame", background=BG_CARD)

        style.configure("Title.TLabel", background=BG_DARK,
                        foreground=FG_PRIMARY, font=self.font_title)
        style.configure("Heading.TLabel", background=BG_CARD,
                        foreground=FG_PRIMARY, font=self.font_heading)
        style.configure("CardText.TLabel", background=BG_CARD,
                        foreground=FG_SECONDARY, font=self.font_normal)
        style.configure("Status.TLabel", background=BG_DARK,
                        foreground=FG_SECONDARY, font=self.font_small)
        style.configure("DarkText.TLabel", background=BG_DARK,
                        foreground=FG_SECONDARY, font=self.font_normal)

        style.configure("Custom.TCombobox", fieldbackground=BG_INPUT,
                        background=BG_INPUT, foreground=FG_PRIMARY,
                        arrowcolor=ACCENT, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER)
        style.map("Custom.TCombobox",
                  fieldbackground=[("readonly", BG_INPUT)],
                  foreground=[("readonly", FG_PRIMARY)],
                  selectbackground=[("readonly", BG_INPUT)],
                  selectforeground=[("readonly", FG_PRIMARY)])

        for name, bg, hover, pressed, font in (
            ("Accent.TButton", ACCENT, ACCENT_HOVER, ACCENT_PRESSED, self.font_normal),
            ("Ghost.TButton", BG_INPUT, BORDER, BORDER, self.font_normal),
            ("Danger.TButton", DANGER, DANGER_HOVER, DANGER_HOVER, self.font_normal),
        ):
            style.configure(name, background=bg, foreground="white",
                            font=font, padding=(12, 7), borderwidth=0, focuscolor=bg)
            style.map(name, background=[("active", hover), ("pressed", pressed),
                                        ("disabled", BG_INPUT)],
                      foreground=[("disabled", FG_MUTED)])

        style.configure("Success.TButton", background=SUCCESS, foreground="white",
                        font=("Segoe UI Semibold", 12), padding=(20, 11),
                        borderwidth=0, focuscolor=SUCCESS)
        style.map("Success.TButton",
                  background=[("active", SUCCESS_HOVER), ("disabled", BG_INPUT)],
                  foreground=[("disabled", FG_MUTED)])

        style.configure("Custom.Horizontal.TProgressbar", troughcolor=BG_INPUT,
                        background=ACCENT, borderwidth=0, thickness=8)

        # Treeview (files list)
        style.configure("Files.Treeview", background=BG_INPUT,
                        fieldbackground=BG_INPUT, foreground=FG_PRIMARY,
                        borderwidth=0, rowheight=26, font=self.font_normal)
        style.configure("Files.Treeview.Heading", background=BG_CARD,
                        foreground=FG_SECONDARY, font=self.font_small,
                        borderwidth=0, relief="flat")
        style.map("Files.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])
        style.map("Files.Treeview.Heading", background=[("active", BG_CARD)])

    # ── UI layout ─────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        root = self.root

        # Title bar
        title_frame = ttk.Frame(root, style="Dark.TFrame")
        title_frame.pack(fill="x", padx=22, pady=(18, 2))
        ttk.Label(title_frame, text="📝  Paper Maker",
                  style="Title.TLabel").pack(side="left")
        ttk.Button(title_frame, text="🧮  PDF → Editable Word", style="Accent.TButton",
                   command=self._on_open_converter).pack(side="left", padx=(16, 0))
        ttk.Button(title_frame, text="Sign out", style="Ghost.TButton",
                   command=self._on_sign_out).pack(side="right", padx=(16, 0))
        self.status_var = tk.StringVar(value="Add one or more .docx files to begin")
        ttk.Label(title_frame, textvariable=self.status_var,
                  style="Status.TLabel").pack(side="right", pady=(8, 0))

        # ── Files strip ──
        files_card = ttk.Frame(root, style="Card.TFrame")
        files_card.pack(fill="x", padx=22, pady=(10, 8))

        fhead = ttk.Frame(files_card, style="Card.TFrame")
        fhead.pack(fill="x", padx=14, pady=(12, 6))
        ttk.Label(fhead, text="Source Files",
                  style="Heading.TLabel").pack(side="left")
        ttk.Button(fhead, text="✕ Remove", style="Danger.TButton",
                   command=self._on_remove_file).pack(side="right")
        ttk.Button(fhead, text="➕ Add Files", style="Accent.TButton",
                   command=self._on_add_files).pack(side="right", padx=(0, 8))

        tree_wrap = tk.Frame(files_card, bg=BG_INPUT, highlightthickness=1,
                             highlightbackground=BORDER)
        tree_wrap.pack(fill="x", padx=14, pady=(0, 14))
        self.files_tree = ttk.Treeview(
            tree_wrap, style="Files.Treeview", height=4,
            columns=("questions", "queued", "answerkey", "status"),
            show="tree headings", selectmode="browse")
        self.files_tree.heading("#0", text="File")
        self.files_tree.heading("questions", text="Questions")
        self.files_tree.heading("queued", text="Queued")
        self.files_tree.heading("answerkey", text="Answer Key")
        self.files_tree.heading("status", text="Status")
        self.files_tree.column("#0", width=400, anchor="w")
        self.files_tree.column("questions", width=80, anchor="center")
        self.files_tree.column("queued", width=70, anchor="center")
        self.files_tree.column("answerkey", width=90, anchor="center")
        self.files_tree.column("status", width=110, anchor="center")
        tsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                            command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=tsb.set)
        tsb.pack(side="right", fill="y")
        self.files_tree.pack(side="left", fill="x", expand=True)
        self.files_tree.bind("<<TreeviewSelect>>", self._on_file_select)

        # ── Bottom bar ──  (packed to the window bottom FIRST so the Generate
        # button and output row always stay visible, even on short screens; the
        # working area below then fills only the space that remains.)
        bottom = ttk.Frame(root, style="Dark.TFrame")
        bottom.pack(side="bottom", fill="x", padx=22, pady=(0, 14))

        out_row = ttk.Frame(bottom, style="Dark.TFrame")
        out_row.pack(fill="x", pady=(0, 8))
        ttk.Label(out_row, text="Output folder",
                  style="DarkText.TLabel").pack(side="left", padx=(0, 8))
        self.output_var = tk.StringVar(
            value=self._config.get("last_output_dir", "") or "Choose where papers are saved…")
        if self._config.get("last_output_dir"):
            self.output_dir = Path(self._config["last_output_dir"])
        out_entry = tk.Entry(
            out_row, textvariable=self.output_var, state="readonly",
            font=self.font_normal, bg=BG_INPUT, fg=FG_SECONDARY,
            readonlybackground=BG_INPUT, relief="flat", highlightthickness=1,
            highlightbackground=BORDER)
        out_entry.pack(side="left", fill="x", expand=True, ipady=5)
        ttk.Button(out_row, text="📂 Browse", style="Ghost.TButton",
                   command=self._on_choose_output).pack(side="left", padx=(8, 0))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            bottom, variable=self.progress_var, maximum=100,
            style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(0, 8))

        self.generate_btn = ttk.Button(
            bottom, text="⬇  Generate All Papers", style="Success.TButton",
            command=self._on_generate)
        self.generate_btn.pack(fill="x")

        # ── Working area (two columns) ──
        content = ttk.Frame(root, style="Dark.TFrame")
        content.pack(side="top", fill="both", expand=True, padx=22, pady=(0, 10))
        # uniform → always two equal halves, so a wide label in one card can't
        # squeeze the other card's button row off the window.
        content.columnconfigure(0, weight=1, uniform="cards")
        content.columnconfigure(1, weight=1, uniform="cards")
        content.rowconfigure(0, weight=1)
        # Size the cards to the space the window actually has, not to what
        # their children ask for; otherwise on a short screen the cards overflow
        # and the button rows pinned to their bottoms end up off-screen. With
        # this, the question/paper lists shrink instead.
        content.grid_propagate(False)

        # LEFT — selector
        left = ttk.Frame(content, style="Card.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.left_title = tk.StringVar(value="Select Questions")
        ttk.Label(left, textvariable=self.left_title,
                  style="Heading.TLabel").pack(anchor="w", padx=16, pady=(14, 2))
        ttk.Label(left, text="Pick a file above, choose a subject, then add questions.",
                  style="CardText.TLabel").pack(anchor="w", padx=16, pady=(0, 10))

        subj_row = ttk.Frame(left, style="Card.TFrame")
        subj_row.pack(fill="x", padx=16, pady=(0, 8))
        ttk.Label(subj_row, text="Subject",
                  style="CardText.TLabel").pack(anchor="w")
        self.subject_var = tk.StringVar()
        self.subject_combo = ttk.Combobox(
            subj_row, textvariable=self.subject_var, state="readonly",
            style="Custom.TCombobox", font=self.font_normal)
        self.subject_combo.pack(fill="x", pady=(4, 0))
        self.subject_combo.bind("<<ComboboxSelected>>", self._on_subject_change)

        search_row = ttk.Frame(left, style="Card.TFrame")
        search_row.pack(fill="x", padx=16, pady=(8, 0))
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_row, textvariable=self.search_var, font=self.font_normal,
            bg=BG_INPUT, fg=FG_PRIMARY, insertbackground=FG_PRIMARY,
            relief="flat", highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT)
        search_entry.pack(fill="x", ipady=5)
        self._placeholder(search_entry, self.search_var, "🔍  Filter questions…")

        # Button row is packed to the card's bottom BEFORE the list so it can
        # never be pushed out of view; the list absorbs any shortfall instead.
        lbtn = ttk.Frame(left, style="Card.TFrame")
        lbtn.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        ttk.Button(lbtn, text="➕ Add Selected", style="Accent.TButton",
                   command=self._on_add_selected).pack(side="left")
        ttk.Button(lbtn, text="Add All in Subject", style="Ghost.TButton",
                   command=self._on_add_all_subject).pack(side="left", padx=(8, 0))

        q_wrap = tk.Frame(left, bg=BG_INPUT, highlightthickness=1,
                          highlightbackground=BORDER)
        q_wrap.pack(side="top", fill="both", expand=True, padx=16, pady=(8, 8))
        self.question_listbox = tk.Listbox(
            q_wrap, selectmode="extended", font=self.font_normal, height=4,
            bg=BG_INPUT, fg=FG_PRIMARY, selectbackground=ACCENT,
            selectforeground="white", activestyle="none", relief="flat",
            highlightthickness=0, borderwidth=0)
        qsb = ttk.Scrollbar(q_wrap, orient="vertical",
                            command=self.question_listbox.yview)
        self.question_listbox.configure(yscrollcommand=qsb.set)
        qsb.pack(side="right", fill="y")
        self.question_listbox.pack(side="left", fill="both", expand=True)
        self.question_listbox.bind("<Double-Button-1>",
                                   lambda _e: self._on_add_selected())
        # Wire live filtering now that the listbox exists.
        self.search_var.trace_add("write", lambda *_: self._refresh_question_list())

        # RIGHT — export list
        right = ttk.Frame(content, style="Card.TFrame")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(right, text="This File's Paper",
                  style="Heading.TLabel").pack(anchor="w", padx=16, pady=(14, 2))
        self.export_count_var = tk.StringVar(value="No file selected")
        ttk.Label(right, textvariable=self.export_count_var,
                  style="CardText.TLabel").pack(anchor="w", padx=16, pady=(0, 6))

        # Output selector — only enabled when the file carries an answer key.
        mode_row = ttk.Frame(right, style="Card.TFrame")
        mode_row.pack(fill="x", padx=16, pady=(0, 8))
        ttk.Label(mode_row, text="Output",
                  style="CardText.TLabel").pack(side="left", padx=(0, 8))
        self.mode_var = tk.StringVar(value=LABEL_FOR_MODE[MODE_QUESTIONS])
        self.mode_combo = ttk.Combobox(
            mode_row, textvariable=self.mode_var, state="disabled",
            values=list(MODE_LABELS), style="Custom.TCombobox",
            font=self.font_normal, width=24)
        self.mode_combo.pack(side="left")
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_change)
        self.mode_hint_var = tk.StringVar(value="")
        ttk.Label(mode_row, textvariable=self.mode_hint_var,
                  style="CardText.TLabel").pack(side="left", padx=(10, 0))

        ectrl = ttk.Frame(right, style="Card.TFrame")
        ectrl.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        # "Clear" is packed first so it is allocated room before the left-hand
        # buttons; pack drops the LAST-packed widget when a row runs out of width.
        ttk.Button(ectrl, text="Clear", style="Danger.TButton",
                   command=self._on_clear_all).pack(side="right")
        ttk.Button(ectrl, text="▲ Up", style="Ghost.TButton",
                   command=self._on_move_up).pack(side="left", padx=(0, 4))
        ttk.Button(ectrl, text="▼ Down", style="Ghost.TButton",
                   command=self._on_move_down).pack(side="left", padx=(0, 4))
        ttk.Button(ectrl, text="✕ Remove", style="Danger.TButton",
                   command=self._on_remove_selected).pack(side="left", padx=(8, 4))

        e_wrap = tk.Frame(right, bg=BG_INPUT, highlightthickness=1,
                          highlightbackground=BORDER)
        e_wrap.pack(side="top", fill="both", expand=True, padx=16, pady=(0, 8))
        self.export_listbox = tk.Listbox(
            e_wrap, selectmode="extended", font=self.font_normal, height=4,
            bg=BG_INPUT, fg=FG_PRIMARY, selectbackground=ACCENT,
            selectforeground="white", activestyle="none", relief="flat",
            highlightthickness=0, borderwidth=0)
        esb = ttk.Scrollbar(e_wrap, orient="vertical",
                            command=self.export_listbox.yview)
        self.export_listbox.configure(yscrollcommand=esb.set)
        esb.pack(side="right", fill="y")
        self.export_listbox.pack(side="left", fill="both", expand=True)

    # ── Small helpers ─────────────────────────────────────────────────────
    def _placeholder(self, entry: tk.Entry, var: tk.StringVar, text: str) -> None:
        """Lightweight placeholder behaviour for an Entry."""
        def show() -> None:
            if not var.get():
                entry._ph = True
                entry.config(fg=FG_MUTED)
                var.set(text)

        def on_focus_in(_e) -> None:
            if getattr(entry, "_ph", False):
                var.set("")
                entry.config(fg=FG_PRIMARY)
                entry._ph = False

        def on_focus_out(_e) -> None:
            if not var.get():
                show()

        entry._ph = False
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
        show()

    def _search_text(self) -> str:
        if getattr(self.question_listbox, "_ignore", False):
            return ""
        val = self.search_var.get()
        if val.startswith("🔍"):
            return ""
        return val.strip().lower()

    def _current_state(self) -> FileState | None:
        if self.current is None:
            return None
        return self.files.get(self.current)

    # ── File management ───────────────────────────────────────────────────
    def _on_add_files(self) -> None:
        initial = self._config.get("last_input_dir") or None
        paths = filedialog.askopenfilenames(
            title="Select Source Documents",
            filetypes=[("Papers (Word or PDF)", "*.docx *.pdf"),
                       ("Word Documents", "*.docx"),
                       ("PDF Documents", "*.pdf")],
            initialdir=initial)
        if not paths:
            return

        new_paths: list[Path] = []
        for p in paths:
            path = Path(p)
            if path in self.files:
                continue
            self.files[path] = FileState(path)
            self._insert_file_row(path)
            new_paths.append(path)

        if new_paths:
            self._config["last_input_dir"] = str(new_paths[0].parent)
            self._save_config()
            if self.current is None:
                self._select_file(new_paths[0])
            self._start_worker(self._parse_worker, new_paths)
            self._update_status()

    def _insert_file_row(self, path: Path) -> None:
        iid = self.files_tree.insert(
            "", "end", text="  " + path.name,
            values=("…", "0", "…", "parsing…"))
        self._tree_ids[iid] = path

    def _row_for(self, path: Path) -> str | None:
        for iid, p in self._tree_ids.items():
            if p == path:
                return iid
        return None

    def _refresh_file_row(self, path: Path) -> None:
        iid = self._row_for(path)
        st = self.files.get(path)
        if iid is None or st is None:
            return
        if st.status == "parsing":
            vals = ("…", str(len(st.export)), "…", "parsing…")
        elif st.status == "error":
            vals = ("—", "0", "—", "⚠ error")
        else:
            vals = (str(st.total_questions), str(len(st.export)),
                    "✓ found" if st.has_answer_key else "—", "✓ ready")
        self.files_tree.item(iid, values=vals)

    def _on_remove_file(self) -> None:
        sel = self.files_tree.selection()
        if not sel:
            return
        iid = sel[0]
        path = self._tree_ids.get(iid)
        if path is None:
            return
        self.files_tree.delete(iid)
        self._tree_ids.pop(iid, None)
        self.files.pop(path, None)
        if self.current == path:
            self.current = None
            remaining = next(iter(self.files), None)
            if remaining:
                self._select_file(remaining)
            else:
                self._clear_working_panels()
        self._update_status()

    def _on_file_select(self, _event=None) -> None:
        sel = self.files_tree.selection()
        if not sel:
            return
        path = self._tree_ids.get(sel[0])
        if path and path != self.current:
            self._select_file(path)

    def _select_file(self, path: Path) -> None:
        self.current = path
        iid = self._row_for(path)
        if iid and iid not in self.files_tree.selection():
            self.files_tree.selection_set(iid)
        st = self.files.get(path)
        name = path.name if len(path.name) <= 40 else path.name[:37] + "…"
        self.left_title.set(f"Select Questions — {name}")
        self._sync_mode_combo(st)
        if st and st.subjects:
            names = list(st.subjects.keys())
            self.subject_combo["values"] = names
            self.subject_combo.current(0)
        else:
            self.subject_combo["values"] = []
            self.subject_var.set("")
        self._refresh_question_list()
        self._refresh_export_list()

    def _clear_working_panels(self) -> None:
        self.left_title.set("Select Questions")
        self.subject_combo["values"] = []
        self.subject_var.set("")
        self.question_listbox.delete(0, "end")
        self.export_listbox.delete(0, "end")
        self.export_count_var.set("No file selected")

    # ── Output mode (questions / answer key / both) ───────────────────────
    def _sync_mode_combo(self, st: FileState | None) -> None:
        """Show the selected file's output mode; enable it only with a key."""
        if st is not None and st.has_answer_key:
            self.mode_combo.configure(state="readonly")
            self.mode_var.set(LABEL_FOR_MODE.get(st.mode, LABEL_FOR_MODE[MODE_BOTH]))
            self.mode_hint_var.set("answer key detected")
        else:
            self.mode_combo.configure(state="disabled")
            self.mode_var.set(LABEL_FOR_MODE[MODE_QUESTIONS])
            self.mode_hint_var.set("" if st is None else "no answer key in this file")

    def _on_mode_change(self, _event=None) -> None:
        st = self._current_state()
        if st is not None:
            st.mode = MODE_LABELS.get(self.mode_var.get(), MODE_QUESTIONS)

    # ── Question selection ────────────────────────────────────────────────
    def _on_subject_change(self, _event=None) -> None:
        self._refresh_question_list()

    def _visible_questions(self) -> list[QuestionInfo]:
        st = self._current_state()
        if not st:
            return []
        subj = self.subject_var.get()
        items = st.subjects.get(subj, [])
        needle = self._search_text()
        if needle:
            items = [q for q in items
                     if needle in q.preview.lower() or needle in str(q.number)]
        return items

    def _refresh_question_list(self) -> None:
        self.question_listbox.delete(0, "end")
        for qi in self._visible_questions():
            self.question_listbox.insert("end", qi.label())

    def _add_questions(self, questions: list[QuestionInfo]) -> None:
        st = self._current_state()
        if not st:
            return
        existing = {(e.subject, e.number) for e in st.export}
        added = 0
        for qi in questions:
            key = (qi.subject, qi.number)
            if key in existing:
                continue
            st.export.append(qi)
            existing.add(key)
            added += 1
        if added:
            # one block per subject, the way the paper is built
            st.export[:] = group_by_subject(st.export)
            self._refresh_export_list()
            self._refresh_file_row(st.path)

    def _on_add_selected(self) -> None:
        if not self._current_state():
            messagebox.showinfo("No File", "Add and select a file first.")
            return
        visible = self._visible_questions()
        idxs = self.question_listbox.curselection()
        if not idxs:
            messagebox.showinfo("No Selection", "Select one or more questions "
                                "(or double-click one).")
            return
        self._add_questions([visible[i] for i in idxs])

    def _on_add_all_subject(self) -> None:
        st = self._current_state()
        if not st:
            return
        self._add_questions(list(st.subjects.get(self.subject_var.get(), [])))

    # ── Export list editing ───────────────────────────────────────────────
    def _refresh_export_list(self) -> None:
        self.export_listbox.delete(0, "end")
        st = self._current_state()
        if not st:
            self.export_count_var.set("No file selected")
            return
        for i, qi in enumerate(st.export, 1):
            self.export_listbox.insert(
                "end", f"{i}.  [{qi.subject}] Q{qi.number}   {qi.preview[:46]}")
        n = len(st.export)
        self.export_count_var.set(
            f"{n} question{'s' if n != 1 else ''} in this paper"
            if n else "No questions yet — add some from the left")

    def _on_remove_selected(self) -> None:
        st = self._current_state()
        if not st:
            return
        for idx in sorted(self.export_listbox.curselection(), reverse=True):
            del st.export[idx]
        self._refresh_export_list()
        self._refresh_file_row(st.path)

    def _on_clear_all(self) -> None:
        st = self._current_state()
        if not st or not st.export:
            return
        if messagebox.askyesno("Clear", "Remove all questions from this file's paper?"):
            st.export.clear()
            self._refresh_export_list()
            self._refresh_file_row(st.path)

    def _on_move_up(self) -> None:
        st = self._current_state()
        sel = self.export_listbox.curselection()
        if not st or not sel or sel[0] == 0:
            return
        moved = list(st.export)
        for idx in sel:
            moved[idx - 1], moved[idx] = moved[idx], moved[idx - 1]
        if group_by_subject(moved) != moved:
            return  # a question stays inside its subject's block
        st.export[:] = moved
        self._refresh_export_list()
        for idx in sel:
            self.export_listbox.selection_set(idx - 1)

    def _on_move_down(self) -> None:
        st = self._current_state()
        sel = self.export_listbox.curselection()
        if not st or not sel or sel[-1] >= len(st.export) - 1:
            return
        moved = list(st.export)
        for idx in reversed(sel):
            moved[idx + 1], moved[idx] = moved[idx], moved[idx + 1]
        if group_by_subject(moved) != moved:
            return  # a question stays inside its subject's block
        st.export[:] = moved
        self._refresh_export_list()
        for idx in sel:
            self.export_listbox.selection_set(idx + 1)

    # ── Output + generation ───────────────────────────────────────────────
    def _on_choose_output(self) -> None:
        initial = (str(self.output_dir) if self.output_dir
                   else self._config.get("last_output_dir") or None)
        folder = filedialog.askdirectory(title="Choose Output Folder",
                                         initialdir=initial)
        if not folder:
            return
        self.output_dir = Path(folder)
        self.output_var.set(folder)
        self._config["last_output_dir"] = folder
        self._save_config()

    def _on_generate(self) -> None:
        if self._generating:
            return
        jobs: list[BatchJob] = []
        for path, st in self.files.items():
            if st.status != "ready" or not st.export:
                continue
            mode = st.mode if st.has_answer_key else MODE_QUESTIONS
            jobs.append(BatchJob(path, list(st.export), Path(), mode=mode))  # output set below

        if not jobs:
            messagebox.showinfo(
                "Nothing to Generate",
                "Add questions to at least one file's paper first.")
            return
        if self.output_dir is None:
            messagebox.showwarning("No Output Folder",
                                   "Choose an output folder first.")
            return
        if not self.output_dir.exists():
            messagebox.showerror("Bad Folder",
                                 f"Output folder does not exist:\n{self.output_dir}")
            return

        # Resolve output paths (avoid overwriting a source that lives there).
        for job in jobs:
            # a source made by the PDF converter is called "<paper>_editable.docx":
            # the new paper is named after the paper, not after that working file
            stem = re.sub(r"_editable(?:_[A-Za-z0-9]+)?$", "", job.source.stem) or job.source.stem
            out = self.output_dir / f"{stem}_paper.docx"
            if out.resolve() == job.source.resolve():
                out = self.output_dir / f"{stem}_paper_out.docx"
            job.output = out

        self._generating = True
        self.generate_btn.state(["disabled"])
        self.progress_var.set(0)
        self.status_var.set(f"Generating {len(jobs)} paper(s)…")
        self._start_worker(self._gen_worker, jobs)

    # ── Background workers ────────────────────────────────────────────────
    def _start_worker(self, target, arg) -> None:
        t = threading.Thread(target=target, args=(arg,), daemon=True)
        self._workers.append(t)
        t.start()
        self.root.after(60, self._poll_events)

    def _parse_worker(self, paths: list[Path]) -> None:
        for path in paths:
            try:
                if path.suffix.lower() == ".pdf":
                    from pdf_core import parse_pdf
                    paper = parse_pdf(path)
                else:
                    paper = parse_paper(path)
                self._events.put(("parsed", path, paper, None))
            except Exception as exc:  # noqa: BLE001
                self._events.put(("parsed", path, None, str(exc)))

    def _gen_worker(self, jobs: list[BatchJob]) -> None:
        def cb(done, total, job, error) -> None:
            self._events.put(("gen_progress", done, total, job.output.name,
                              str(error) if error else None))
        try:
            results = generate_papers_batch(jobs, progress_callback=cb)
            self._events.put(("gen_done", results, None))
        except Exception as exc:  # noqa: BLE001
            self._events.put(("gen_done", None, str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self._events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self._workers = [t for t in self._workers if t.is_alive()]
        if self._workers or not self._events.empty():
            self.root.after(60, self._poll_events)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "parsed":
            _, path, paper, error = event
            st = self.files.get(path)
            if not st:
                return
            if error is not None:
                st.status = "error"
                st.error = error
            else:
                st.subjects = paper.subjects
                st.has_answer_key = paper.has_answer_key
                st.mode = MODE_BOTH if paper.has_answer_key else MODE_QUESTIONS
                st.status = "ready"
            self._refresh_file_row(path)
            if path == self.current:
                self._select_file(path)
            self._update_status()

        elif kind == "gen_progress":
            _, done, total, name, error = event
            self.progress_var.set(done / total * 100)
            tag = "⚠ " if error else ""
            self.status_var.set(f"{tag}Generated {done}/{total}: {name}")

        elif kind == "gen_done":
            _, results, error = event
            self._generating = False
            self.generate_btn.state(["!disabled"])
            if error is not None:
                self.progress_var.set(0)
                self.status_var.set("Generation failed")
                messagebox.showerror("Error", f"Generation failed:\n{error}")
                return
            self.progress_var.set(100)
            ok = [j for j, e in results if e is None]
            failed = [(j, e) for j, e in results if e is not None]
            saved = [p for j in ok for p in j.saved]
            self.status_var.set(
                f"✅ Done — {len(saved)} file(s) saved"
                + (f", {len(failed)} failed" if failed else ""))
            msg = f"{len(saved)} file(s) saved to:\n{self.output_dir}"
            if saved:
                msg += "\n\n" + "\n".join(f"• {p.name}" for p in saved)
            renamed = [p for p in saved if re.search(r" \(\d+\)\.docx$", p.name)]
            if renamed:
                msg += ("\n\nNote: a file with the same name was open in Word, "
                        "so the new copy was saved under a numbered name.")
            if failed:
                msg += "\n\nFailed:\n" + "\n".join(
                    f"• {j.source.name}: {e}" for j, e in failed)
                messagebox.showwarning("Completed with errors", msg)
            else:
                messagebox.showinfo("Success", msg)

    def _update_status(self) -> None:
        if self._generating:
            return
        n = len(self.files)
        ready = sum(1 for s in self.files.values() if s.status == "ready")
        queued = sum(len(s.export) for s in self.files.values())
        if n == 0:
            self.status_var.set("Add one or more .docx files to begin")
        else:
            self.status_var.set(
                f"{n} file(s) · {ready} ready · {queued} question(s) queued")

    # ── PDF → editable Word (separate window) ─────────────────────────────
    def _on_open_converter(self) -> None:
        win = getattr(self, "_convert_window", None)
        if win is not None and win.win.winfo_exists():
            win.win.deiconify()
            win.win.lift()
            return
        from convert_window import ConvertWindow
        self._convert_window = ConvertWindow(self)

    def _on_sign_out(self) -> None:
        win = getattr(self, "_convert_window", None)
        if win is not None and win.win.winfo_exists():
            messagebox.showinfo("Sign out", "Close the PDF → Editable Word window first.")
            return
        from access_control import sign_out
        sign_out(self)

    # ── Lifecycle ─────────────────────────────────────────────────────────
    def _on_close(self) -> None:
        win = getattr(self, "_convert_window", None)
        if win is not None and win.running:
            if not messagebox.askyesno("Stop conversion?",
                                       "A PDF conversion is still running. Stop it and quit?"):
                return
            win._on_cancel()
        self._save_config()
        self.root.destroy()

    def run(self) -> None:
        # Only emails on the owner's access sheet get past this point.
        from access_control import ensure_access
        if not ensure_access(self):
            self.root.destroy()
            return
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        screen_w, screen_h = (self.root.winfo_screenwidth(),
                              self.root.winfo_screenheight())
        # Never open taller/wider than the screen (leave room for the taskbar),
        # so the bottom bar with the Generate button is always on-screen.
        w = min(w, screen_w - 40)
        h = min(h, screen_h - 80)
        x = max(0, (screen_w - w) // 2)
        y = max(0, (screen_h - 80 - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()  # safe no-op; guards frozen builds
    PaperMakerApp().run()
