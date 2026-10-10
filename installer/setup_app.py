"""PaperMaker_Setup.exe — small online installer / updater for Paper Maker.

Reads ``manifest.json`` from the download bucket (see pack_release.py), downloads
the parts of every component that is missing or out of date (resumable, retried,
SHA-256 checked), unpacks them into %LOCALAPPDATA%\\Programs\\PaperMaker, runs a
self-test of the engine and creates the Desktop / Start-menu shortcuts.  It copies
itself into the install folder as "PaperMaker Setup.exe", which then serves as
the updater ("Update Paper Maker" in the Start menu) and the uninstaller
(Settings > Apps).  No admin rights are needed.

    PaperMaker_Setup.exe [--url https://.../] [--dir FOLDER] [--uninstall]
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import queue
import shutil
import subprocess
import sys
import tarfile
import threading
import time
import traceback
import urllib.error
import urllib.request
import uuid
import winreg
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from _setup_config import BASE_URL  # written by build_setup.py
except ImportError:
    BASE_URL = ""

APP = "Paper Maker"
EXE = "PaperMaker.exe"
SETUP_COPY = "PaperMaker Setup.exe"
OWNED = [EXE, "READ ME FIRST.txt", "engine", "mineru_home", ".setup"]  # removed on uninstall
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\PaperMaker"
FORMAT = 1
SPACE_MARGIN = 300 * 2**20
# the engine's deepest file is 184 characters below the install folder; Windows
# programs (Python too) fail beyond 259 characters unless long paths are enabled
MAX_DIR = 70
NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW
SELFTEST = ("import mineru, onnxruntime, fitz, docx, lxml, cv2, numpy, pypandoc; "
            "pypandoc.get_pandoc_version(); print('SELFTEST-OK')")


class Cancelled(Exception):
    pass


class SetupError(Exception):
    """Problem with a message meant for the user."""


# ---------------------------------------------------------------- helpers
def default_dir() -> Path:
    d = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "Programs" / "PaperMaker"
    if len(str(d)) > MAX_DIR:  # very long user name: fall back to a short folder
        d = Path(os.environ.get("SystemDrive", "C:") + "\\") / "PaperMaker"
    return d


def long_path(p: Path) -> Path:
    s = str(p.resolve())
    return Path(s if s.startswith("\\\\?\\") else "\\\\?\\" + s)


def known_folder(guid: str) -> Path | None:
    """Real Desktop / Start-menu folder (follows OneDrive redirection)."""
    class GUID(ctypes.Structure):
        _fields_ = [("d1", ctypes.c_ulong), ("d2", ctypes.c_ushort),
                    ("d3", ctypes.c_ushort), ("d4", ctypes.c_ubyte * 8)]
    u = uuid.UUID(guid)
    g = GUID(u.fields[0], u.fields[1], u.fields[2], (ctypes.c_ubyte * 8)(*u.bytes[8:]))
    p = ctypes.c_wchar_p()
    if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(p)) != 0:
        return None
    try:
        return Path(p.value)
    finally:
        ctypes.windll.ole32.CoTaskMemFree(p)


DESKTOP = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
PROGRAMS = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"


def shortcut_paths() -> list[Path]:
    out = []
    if d := known_folder(DESKTOP):
        out.append(d / f"{APP}.lnk")
    if p := known_folder(PROGRAMS):
        out += [p / f"{APP}.lnk", p / f"Update {APP}.lnk"]
    return out


def make_shortcut(lnk: Path, target: Path, args: str = "") -> None:
    """.lnk via WScript.Shell; falls back to a .url file if PowerShell/COM is blocked."""
    lnk.parent.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ, PM_LNK=str(lnk), PM_TARGET=str(target), PM_ARGS=args,
               PM_DIR=str(target.parent))
    ps = ("$s=(New-Object -ComObject WScript.Shell).CreateShortcut($env:PM_LNK);"
          "$s.TargetPath=$env:PM_TARGET;$s.Arguments=$env:PM_ARGS;"
          "$s.WorkingDirectory=$env:PM_DIR;$s.IconLocation=$env:PM_TARGET+',0';$s.Save()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], env=env,
                       creationflags=NO_WINDOW, timeout=60, check=True, capture_output=True)
        if lnk.exists():
            return
    except Exception:
        pass
    if not args:
        lnk.with_suffix(".url").write_text(
            f"[InternetShortcut]\nURL={target.as_uri()}\nIconFile={target}\nIconIndex=0\n", "utf-8")


def app_running() -> bool:
    try:
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {EXE}", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, creationflags=NO_WINDOW, timeout=30).stdout
        return EXE.lower() in out.lower()
    except Exception:
        return False


def mb(n: float) -> str:
    return f"{n / 2**20:,.0f} MB" if n < 2**30 else f"{n / 2**30:.1f} GB"


class ChainReader:
    """Reads several part files as one stream (for tarfile's streaming mode)."""

    def __init__(self, files: list[Path], on_read):
        self.files, self.fp, self.on_read = list(files), None, on_read

    def read(self, n: int = -1) -> bytes:
        while True:
            if self.fp is None:
                if not self.files:
                    return b""
                self.fp = self.files.pop(0).open("rb")
            data = self.fp.read(n if n and n > 0 else 1 << 20)
            if data:
                self.on_read(len(data))
                return data
            self.fp.close()
            self.fp = None

    def close(self) -> None:
        if self.fp:
            self.fp.close()


# ---------------------------------------------------------------- installer
class Installer:
    def __init__(self, base_url: str, target: Path, post, cancel: threading.Event):
        self.base = base_url.rstrip("/") + "/"
        self.dir = target
        self.meta = target / ".setup"
        self.post = post  # post(kind, payload) -> GUI thread
        self.cancel = cancel
        self.log_fp = None

    # -- logging / progress
    def log(self, msg: str) -> None:
        line = f"{time.strftime('%H:%M:%S')}  {msg}"
        if self.log_fp:
            self.log_fp.write(line + "\n")
            self.log_fp.flush()

    def status(self, msg: str) -> None:
        self.log(msg)
        self.post("status", msg)

    def check_cancel(self) -> None:
        if self.cancel.is_set():
            raise Cancelled

    # -- network
    def open_url(self, url: str, start: int = 0):
        req = urllib.request.Request(url, headers={"User-Agent": "PaperMakerSetup/1",
                                                   "Cache-Control": "no-cache"})
        if start:
            req.add_header("Range", f"bytes={start}-")
        return urllib.request.urlopen(req, timeout=30)

    def fetch_manifest(self) -> dict:
        url = f"{self.base}manifest.json?t={int(time.time())}"
        last = None
        for attempt in range(4):
            self.check_cancel()
            try:
                with self.open_url(url) as r:
                    m = json.loads(r.read().decode("utf-8"))
                if m.get("format") != FORMAT:
                    raise SetupError("This setup is too old for the current release.\n"
                                     "Please ask for the new PaperMaker_Setup.exe.")
                return m
            except SetupError:
                raise
            except Exception as e:
                last = e
                self.log(f"manifest attempt {attempt + 1}: {e!r}")
                time.sleep(2 * (attempt + 1))
        raise SetupError("Could not reach the download server.\n\n"
                         "Check the internet connection (office networks sometimes block "
                         f"downloads; try another network or a hotspot).\n\nDetails: {last}")

    def download(self, part: dict, progress) -> Path:
        dest = self.meta / "download" / part["file"]
        if dest.exists() and dest.stat().st_size == part["size"]:
            progress(part["size"])  # verified when it was downloaded
            return dest
        tmp = dest.with_name(dest.name + ".partial")
        url = self.base + part["file"]
        for full_try in range(3):
            fails = 0
            while True:
                self.check_cancel()
                have = tmp.stat().st_size if tmp.exists() else 0
                if have > part["size"]:
                    tmp.unlink()
                    have = 0
                if have == part["size"]:
                    break
                try:
                    with self.open_url(url, have) as r:
                        mode = "ab" if have and r.status == 206 else "wb"
                        if mode == "wb":
                            have = 0
                        progress(have)
                        with tmp.open(mode) as fp:
                            while chunk := r.read(1 << 20):
                                fp.write(chunk)
                                have += len(chunk)
                                fails = 0
                                progress(have)
                                self.check_cancel()
                except Cancelled:
                    raise
                except Exception as e:
                    fails += 1
                    self.log(f"{part['file']}: {e!r} (try {fails})")
                    if isinstance(e, urllib.error.HTTPError) and e.code in (403, 404):
                        raise SetupError(f"The download file {part['file']} is missing on the "
                                         f"server (HTTP {e.code}). Please tell the tool owner.")
                    if fails >= 8:
                        raise SetupError("The download keeps failing - the internet connection "
                                         "seems unstable.\nRun the setup again later: it continues "
                                         f"where it stopped.\n\nDetails: {e}")
                    self.post("status", f"Connection problem, retrying in a moment ... ({fails}/8)")
                    for _ in range(min(2 ** fails, 30)):
                        self.check_cancel()
                        time.sleep(1)
            h = hashlib.sha256()
            with tmp.open("rb") as fp:
                while chunk := fp.read(4 << 20):
                    h.update(chunk)
            if h.hexdigest() == part["sha256"]:
                tmp.replace(dest)
                return dest
            self.log(f"{part['file']}: checksum mismatch, downloading again")
            tmp.unlink()
        raise SetupError(f"{part['file']} arrived damaged three times.\n"
                         "Try another network, or tell the tool owner.")

    # -- state
    def load_state(self) -> dict:
        try:
            return json.loads((self.meta / "state.json").read_text("utf-8"))
        except Exception:
            return {"components": {}}

    def save_state(self, state: dict) -> None:
        tmp = self.meta / "state.json.tmp"
        tmp.write_text(json.dumps(state, indent=1), "utf-8")
        tmp.replace(self.meta / "state.json")

    def installed(self, c: dict, state: dict) -> bool:
        return (state["components"].get(c["name"]) == c["id"]
                and all((self.dir / p).exists() for p in c["paths"]))

    # -- unpack
    def unpack(self, c: dict, parts: list[Path]) -> None:
        # \\?\ paths: the staging folder adds depth that would cross MAX_PATH
        stage = long_path(self.meta / "staging" / c["name"])
        old = long_path(self.meta / "old" / c["name"])
        for d in (stage, old):
            shutil.rmtree(d, ignore_errors=True)
        stage.mkdir(parents=True)
        total, done, shown = sum(p.stat().st_size for p in parts), 0, 0

        def on_read(n):  # tarfile reads ~10 KB at a time: post every 0.5 %
            nonlocal done, shown
            done += n
            if done - shown > total / 200:
                shown = done
                self.post("progress", done / total)

        reader = ChainReader(parts, on_read)
        try:
            with tarfile.open(fileobj=reader, mode="r|" + c["compression"]) as tar:
                for member in tar:
                    self.check_cancel()
                    tar.extract(member, stage, filter="data")
        finally:
            reader.close()
        for p in c["paths"]:
            if not (stage / p).exists():
                raise SetupError(f"The download of '{c['name']}' is incomplete ({p} missing).")
        # swap: current -> old, staged -> current; put the old one back on failure
        moved, cur = [], long_path(self.dir)
        try:
            for p in c["paths"]:
                if (cur / p).exists():
                    (old / p).parent.mkdir(parents=True, exist_ok=True)
                    os.replace(cur / p, old / p)
                    moved.append(p)
            for p in c["paths"]:
                (cur / p).parent.mkdir(parents=True, exist_ok=True)
                os.replace(stage / p, cur / p)
        except OSError as e:
            for p in c["paths"]:
                if p in moved and not (cur / p).exists():
                    os.replace(old / p, cur / p)
            raise SetupError(f"Could not replace {c['name']} files - is {APP} still open?\n"
                             f"Close it and run the setup again.\n\nDetails: {e}")
        shutil.rmtree(old, ignore_errors=True)
        shutil.rmtree(stage, ignore_errors=True)

    # -- after install
    def selftest(self) -> None:
        py = self.dir / "engine" / "python.exe"
        env = {k: v for k, v in os.environ.items() if k not in ("PYTHONPATH", "PYTHONHOME")}
        try:
            r = subprocess.run([str(py), "-E", "-s", "-X", "utf8", "-c", SELFTEST], env=env,
                               capture_output=True, text=True, encoding="utf-8", errors="replace",
                               timeout=300, creationflags=NO_WINDOW, cwd=self.dir)
        except Exception as e:
            raise SetupError(f"The PDF engine could not start.\n\nDetails: {e}")
        self.log(f"selftest exit {r.returncode}: {r.stdout.strip()[-300:]} {r.stderr.strip()[-1500:]}")
        if "SELFTEST-OK" not in r.stdout:
            raise SetupError("The PDF engine did not pass its check (antivirus may have "
                             "blocked or removed a file).\nUse 'Repair' to download it again.")

    def finish(self, version: str, size: int) -> None:
        exe = self.dir / EXE
        me = Path(sys.executable)
        setup = self.dir / SETUP_COPY
        if getattr(sys, "frozen", False) and me.resolve() != setup.resolve():
            try:
                shutil.copy2(me, setup)
            except OSError as e:
                self.log(f"copy setup: {e!r}")
        links = shortcut_paths()
        for lnk in links:
            if lnk.name.startswith("Update"):
                if setup.exists():
                    make_shortcut(lnk, setup)
            else:
                make_shortcut(lnk, exe)
        try:  # SHCNE_ASSOCCHANGED: Explorer re-reads icons (a new app icon shows at once)
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
        except Exception:
            pass
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY) as k:
                for name, val in [("DisplayName", APP), ("DisplayVersion", version),
                                  ("Publisher", "PW"), ("InstallLocation", str(self.dir)),
                                  ("DisplayIcon", str(exe)),
                                  ("UninstallString", f'"{setup}" --uninstall --dir "{self.dir}"')]:
                    winreg.SetValueEx(k, name, 0, winreg.REG_SZ, val)
                winreg.SetValueEx(k, "EstimatedSize", 0, winreg.REG_DWORD, size // 1024)
                winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)
        except OSError as e:
            self.log(f"registry: {e!r}")

    # -- main
    def run(self, repair: bool) -> str:
        self.meta.mkdir(parents=True, exist_ok=True)
        (self.meta / "download").mkdir(exist_ok=True)
        self.log_fp = (self.meta / "setup.log").open("a", encoding="utf-8")
        self.log(f"==== setup start, dir={self.dir}, url={self.base}, repair={repair}")
        self.status("Contacting the download server ...")
        m = self.fetch_manifest()
        state = self.load_state()
        need = [c for c in m["components"] if repair or not self.installed(c, state)]
        size = sum(c["unpacked_size"] for c in m["components"])
        self.log(f"version {m['version']}; need {[c['name'] for c in need]}")

        if need:
            pending = [p for c in need for p in c["parts"]
                       if not ((self.meta / "download" / p["file"]).exists()
                               and (self.meta / "download" / p["file"]).stat().st_size == p["size"])]
            want = sum(p["size"] for p in pending) + sum(c["unpacked_size"] for c in need) + SPACE_MARGIN
            free = shutil.disk_usage(self.dir).free
            if free < want:
                raise SetupError(f"Not enough free space on drive {self.dir.drive or self.dir}.\n"
                                 f"Needed: {mb(want)}   Free: {mb(free)}\n\n"
                                 "Free up space, or choose another folder.")
            total = sum(p["size"] for c in need for p in c["parts"])
            got = 0
            for c in need:
                for i, part in enumerate(c["parts"]):
                    def progress(n, base=got):  # n = bytes of this part so far
                        self.post("progress", (base + n) / total)
                        self.post("detail", f"{mb(base + n)} of {mb(total)}")

                    self.status(f"Downloading {c['name']} ({i + 1} of {len(c['parts'])}) ...")
                    self.download(part, progress)
                    got += part["size"]
            self.post("detail", "")
            for c in need:
                self.status(f"Unpacking {c['name']} ...")
                files = [self.meta / "download" / p["file"] for p in c["parts"]]
                self.unpack(c, files)
                state["components"][c["name"]] = c["id"]
                state["version"] = m["version"]
                self.save_state(state)
                for f in files:
                    f.unlink(missing_ok=True)
        self.post("progress", None)
        self.status("Checking the PDF engine ...")
        self.selftest()
        self.status("Creating shortcuts ...")
        self.finish(m["version"], size)
        state["version"] = m["version"]
        self.save_state(state)
        self.log("==== done")
        return m["version"] if need else ""


# ---------------------------------------------------------------- GUI
class SetupWindow:
    def __init__(self, base_url: str, target: Path):
        self.base_url, self.target = base_url, target
        self.q: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.worker = None

        self.root = root = tk.Tk()
        root.title(f"{APP} Setup")
        try:
            root.iconbitmap(default=str(Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
                                        / "assets" / "paper_maker.ico"))
        except tk.TclError:
            pass
        root.resizable(False, False)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        f = ttk.Frame(root, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"{APP}", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        self.msg = ttk.Label(f, text="Installs or updates Paper Maker on this PC "
                             "(about 1 GB download, internet needed only now).",
                             wraplength=520, font=("Segoe UI", 10))
        self.msg.pack(anchor="w", pady=(4, 12))
        row = ttk.Frame(f)
        row.pack(fill="x")
        ttk.Label(row, text="Folder:", font=("Segoe UI", 9)).pack(side="left")
        self.dir_var = tk.StringVar(value=str(target))
        ttk.Entry(row, textvariable=self.dir_var, state="readonly", width=60).pack(
            side="left", fill="x", expand=True, padx=6)
        self.change_btn = ttk.Button(row, text="Change...", command=self.change_dir)
        self.change_btn.pack(side="left")
        self.bar = ttk.Progressbar(f, mode="determinate", maximum=1.0)
        self.bar.pack(fill="x", pady=(18, 4))
        self.status = ttk.Label(f, text="", font=("Segoe UI", 9))
        self.status.pack(anchor="w")
        self.detail = ttk.Label(f, text="", font=("Segoe UI", 9), foreground="#666")
        self.detail.pack(anchor="w")
        btns = ttk.Frame(f)
        btns.pack(side="bottom", fill="x", pady=(14, 0))
        self.close_btn = ttk.Button(btns, text="Close", command=self.on_close)
        self.close_btn.pack(side="right")
        self.repair_btn = ttk.Button(btns, text="Repair", command=lambda: self.start(True))
        self.main_btn = ttk.Button(btns, text="Install", command=lambda: self.start(False))
        self.main_btn.pack(side="right", padx=6)
        if (target / ".setup" / "state.json").exists():
            self.main_btn.config(text="Update")
            self.repair_btn.pack(side="right")
        root.after(100, self.pump)

    def change_dir(self) -> None:
        d = filedialog.askdirectory(initialdir=str(self.target.parent), mustexist=False)
        if d:
            p = Path(d)
            if p.name.lower() != "papermaker":
                p = p / "PaperMaker"
            if "program files" in str(p).lower():
                messagebox.showwarning(APP, "Please choose a folder outside 'Program Files' "
                                       "(it needs administrator rights).")
                return
            if len(str(p)) > MAX_DIR:
                messagebox.showwarning(APP, f"This folder path is too long for Windows "
                                       f"({len(str(p))} characters, at most {MAX_DIR}).\n"
                                       "Please choose a shorter one, e.g. D:\\PaperMaker.")
                return
            self.target = p
            self.dir_var.set(str(p))

    def start(self, repair: bool) -> None:
        if not self.base_url:
            messagebox.showerror(APP, "This setup has no download address built in.")
            return
        if len(str(self.target)) > MAX_DIR:
            messagebox.showerror(APP, f"The folder path is too long for Windows (at most {MAX_DIR} "
                                 "characters).\nClick 'Change...' and choose e.g. D:\\PaperMaker.")
            return
        while app_running():
            if not messagebox.askretrycancel(APP, f"{APP} is open. Please close it, then click Retry."):
                return
        try:
            self.target.mkdir(parents=True, exist_ok=True)
            probe = self.target / ".write_test"
            probe.write_text("x")
            probe.unlink()
        except OSError as e:
            messagebox.showerror(APP, f"Cannot write to {self.target}.\nChoose another folder.\n\n{e}")
            return
        for b in (self.main_btn, self.repair_btn, self.change_btn):
            b.config(state="disabled")
        self.close_btn.config(text="Cancel")
        self.cancel.clear()
        inst = Installer(self.base_url, self.target, lambda k, v: self.q.put((k, v)), self.cancel)

        def work():
            try:
                self.q.put(("done", inst.run(repair)))
            except Cancelled:
                self.q.put(("cancelled", None))
            except SetupError as e:
                inst.log(f"ERROR {e}")
                self.q.put(("error", str(e)))
            except Exception as e:
                inst.log(traceback.format_exc())
                self.q.put(("error", f"Unexpected problem: {e!r}"))
            finally:
                if inst.log_fp:
                    inst.log_fp.close()

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def pump(self) -> None:
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "status":
                    self.status.config(text=val)
                elif kind == "detail":
                    self.detail.config(text=val)
                elif kind == "progress":
                    if val is None:
                        self.bar.config(mode="indeterminate")
                        self.bar.start(15)
                    else:
                        self.bar.stop()
                        self.bar.config(mode="determinate", value=val)
                else:
                    self.finished(kind, val)
        except queue.Empty:
            pass
        self.root.after(100, self.pump)

    def finished(self, kind: str, val) -> None:
        self.worker = None
        self.bar.stop()
        self.bar.config(mode="determinate", value=1.0 if kind == "done" else 0)
        self.close_btn.config(text="Close")
        self.detail.config(text="")
        for b in (self.main_btn, self.repair_btn, self.change_btn):
            b.config(state="normal")
        if kind == "done":
            self.status.config(text="")
            self.msg.config(text=(f"{APP} is ready (version {val})." if val else
                                  f"{APP} is already up to date.") +
                            " Start it from the desktop icon.")
            self.main_btn.config(text="Open Paper Maker", command=self.launch)
            self.repair_btn.pack_forget()
        elif kind == "cancelled":
            self.status.config(text="Stopped. Run the setup again to continue where it stopped.")
        else:
            self.status.config(text="Setup did not finish.")
            log = self.target / ".setup" / "setup.log"
            messagebox.showerror(APP, f"{val}\n\nLog file: {log}")

    def launch(self) -> None:
        try:
            subprocess.Popen([str(self.target / EXE)], cwd=self.target)
        except OSError as e:
            messagebox.showerror(APP, str(e))
            return
        self.root.destroy()

    def on_close(self) -> None:
        if self.worker:
            if messagebox.askyesno(APP, "Stop the setup? It continues where it stopped next time."):
                self.cancel.set()
                self.close_btn.config(state="disabled")
                self.status.config(text="Stopping ...")
            return
        self.root.destroy()


# ---------------------------------------------------------------- uninstall
def uninstall(target: Path) -> None:
    root = tk.Tk()
    root.withdraw()
    if not messagebox.askyesno(APP, f"Remove {APP} from this PC?\n\n{target}"):
        return
    while app_running():
        if not messagebox.askretrycancel(APP, f"{APP} is open. Please close it, then click Retry."):
            return
    for lnk in shortcut_paths():
        for p in (lnk, lnk.with_suffix(".url")):
            p.unlink(missing_ok=True)
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY)
    except OSError:
        pass
    cache = Path(os.environ.get("LOCALAPPDATA", "")) / "PaperMaker" / "engine_cache"
    shutil.rmtree(cache, ignore_errors=True)
    if (target / ".setup" / "state.json").exists():  # only a folder this setup made
        for name in OWNED:
            p = target / name
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            elif p.exists():
                p.unlink(missing_ok=True)
        # this exe usually runs from the folder: delete it (and the then-empty
        # folder; plain rmdir never removes anything else) after it exits
        subprocess.Popen(f'cmd /c ping 127.0.0.1 -n 3 >nul & del /f /q "{target / SETUP_COPY}" '
                         f'& rmdir "{target}"', creationflags=NO_WINDOW, cwd=target.parent)
    messagebox.showinfo(APP, f"{APP} was removed.")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default=BASE_URL)
    ap.add_argument("--dir", type=Path)
    ap.add_argument("--uninstall", action="store_true")
    a = ap.parse_args()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    me = Path(sys.executable)
    here = me.parent if getattr(sys, "frozen", False) and me.name == SETUP_COPY else None
    target = a.dir or here or default_dir()
    if a.uninstall:
        uninstall(target)
        return
    SetupWindow(a.url, target).root.mainloop()


if __name__ == "__main__":
    main()
