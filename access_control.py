"""Access check for Paper Maker.

Only people whose email is on the owner's access list (column A of Sheet1 of the
team's Google Sheet) can open the tool. The user signs in with Google in the
browser, which proves the email; the app sends Google's proof (an ID token) to
the access service (``access_service/worker.js``), which reads the list with the
owner's token and answers only ``allowed: true/false``. Users never need, and
never get, access to the sheet.

The sign-in is remembered (a Google refresh token with the "email" permission
only) and the list is asked again at every start, so an email taken off the
sheet is refused the next time the tool is opened. Without an answer (no
internet, service down) the tool does not open.
"""

from __future__ import annotations

import base64
import hashlib
import http.client
import json
import queue
import secrets
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from tkinter import messagebox, ttk

# Access service (Cloudflare Worker in access_service/); answers POST {"id_token": ...}.
ACCESS_URL = "https://paper-maker-access.access-service.workers.dev/check"
TIMEOUT = 20  # seconds per network call
SIGN_IN_WAIT = 300  # seconds the app waits for the browser sign-in

# Desktop OAuth client of the Google Cloud project "Paper Maker". Google treats a
# desktop client's secret as not confidential; it ships inside the exe.
CLIENT_FILE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "oauth_client.json"
SCOPES = "openid email"
TOKEN_KEY = "google_refresh_token"  # in the app's config file
EMAIL_KEY = "signed_in_email"

SIGNED_IN_PAGE = ("<html><body style='font-family:Segoe UI,sans-serif;text-align:center;"
                  "padding-top:80px'><h2>Paper Maker: signed in.</h2>"
                  "<p>You can close this tab and go back to Paper Maker.</p></body></html>")


class AccessError(Exception):
    """The check could not be done (no internet, service or Google not answering)."""


class SignInNeeded(Exception):
    """No valid Google sign-in: the user has to sign in in the browser."""


class Cancelled(Exception):
    """The user stopped the sign-in."""


def _client() -> dict:
    try:
        return json.loads(CLIENT_FILE.read_text(encoding="utf-8"))["installed"]
    except (OSError, ValueError, KeyError) as exc:
        raise AccessError("The sign-in is not set up: oauth_client.json is missing. "
                          "Please tell the owner of this tool.") from exc


def _post(url: str, data: bytes, content_type: str) -> tuple[int, dict]:
    """POST and return (status, JSON body); AccessError if there is no usable answer."""
    # Cloudflare refuses Python's default User-Agent (error 1010)
    req = urllib.request.Request(url, data=data, headers={"Content-Type": content_type,
                                                          "User-Agent": "PaperMaker/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            status, raw = resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        status, raw = exc.code, exc.read()
    except (OSError, http.client.HTTPException) as exc:
        raise AccessError("Could not reach the access list. Check your internet "
                          "connection and try again.") from exc
    try:
        body = json.loads(raw.decode("utf-8"))
    except ValueError:
        body = None
    if not isinstance(body, dict):
        raise AccessError(f"The access list gave no valid answer (HTTP {status}). "
                          "Please tell the owner of this tool.")
    return status, body


def _google_tokens(fields: dict) -> dict:
    client = _client()
    fields = {**fields, "client_id": client["client_id"], "client_secret": client["client_secret"]}
    status, body = _post(client["token_uri"], urllib.parse.urlencode(fields).encode(),
                         "application/x-www-form-urlencoded")
    if status == 400 and body.get("error") == "invalid_grant":
        raise SignInNeeded()  # signed out, revoked or expired
    if status != 200 or "id_token" not in body:
        raise AccessError(f"Google sign-in failed ({body.get('error', status)}). Try again.")
    return body


def refresh_sign_in(refresh_token: str) -> str:
    """A fresh ID token from the remembered sign-in, without the browser."""
    return _google_tokens({"grant_type": "refresh_token", "refresh_token": refresh_token})["id_token"]


def browser_sign_in(stop: threading.Event) -> tuple[str, str | None]:
    """Sign in with Google in the browser; returns (id_token, refresh_token)."""
    client = _client()
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    state = secrets.token_urlsafe(16)
    answer: dict = {}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_a):
            pass

        def do_GET(self):
            query = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            if query.get("state", [""])[0] == state:
                answer.update({k: v[0] for k, v in query.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(SIGNED_IN_PAGE.encode("utf-8"))

    server = HTTPServer(("127.0.0.1", 0), Handler)
    server.timeout = 0.5
    redirect = f"http://127.0.0.1:{server.server_address[1]}/"
    webbrowser.open(client["auth_uri"] + "?" + urllib.parse.urlencode({
        "client_id": client["client_id"], "redirect_uri": redirect, "response_type": "code",
        "scope": SCOPES, "code_challenge": challenge, "code_challenge_method": "S256",
        "state": state, "access_type": "offline", "prompt": "select_account consent",
    }))
    deadline = time.monotonic() + SIGN_IN_WAIT
    try:
        while "code" not in answer and "error" not in answer:
            if stop.is_set():
                raise Cancelled()
            if time.monotonic() > deadline:
                raise AccessError("The sign-in took too long. Try again.")
            server.handle_request()
    finally:
        server.server_close()
    if "error" in answer:
        raise Cancelled()  # e.g. "access_denied": Cancel on Google's page
    tokens = _google_tokens({"grant_type": "authorization_code", "code": answer["code"],
                             "redirect_uri": redirect, "code_verifier": verifier})
    return tokens["id_token"], tokens.get("refresh_token")


def ask_service(id_token: str) -> tuple[str, bool]:
    """Ask the access service about the signed-in user: (email, allowed)."""
    if not ACCESS_URL:
        raise AccessError("The access list is not connected yet: ACCESS_URL in "
                          "access_control.py is empty.")
    status, body = _post(ACCESS_URL, json.dumps({"id_token": id_token}).encode(), "application/json")
    if status == 401:
        raise SignInNeeded()
    if status != 200 or not isinstance(body.get("email"), str):
        raise AccessError(f"The access list did not answer (HTTP {status}). "
                          "Please tell the owner of this tool.")
    return body["email"], body.get("allowed") is True


class AccessWindow:
    """Sign-in window shown before the main window; ``allowed`` is the result."""

    def __init__(self, app) -> None:
        self.app = app
        self.allowed = False
        self.checking = False
        self.results: "queue.Queue[tuple]" = queue.Queue()
        self.stop = threading.Event()
        self._after: str | None = None
        self._front_after: str | None = None
        self.win = tk.Toplevel(app.root)
        self.win.title("Paper Maker — Sign in")
        self.win.configure(bg=app.root["bg"])
        self.win.resizable(False, False)
        self._build()
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        self.win.update_idletasks()
        x = (self.win.winfo_screenwidth() - self.win.winfo_reqwidth()) // 2
        y = (self.win.winfo_screenheight() - self.win.winfo_reqheight()) // 3
        self.win.geometry(f"+{max(0, x)}+{max(0, y)}")
        self._to_front()

        if app._config.get(TOKEN_KEY):
            # checked again at every start: off the sheet = no access
            self._start(self._recheck_worker, "Checking access…")
        else:
            self._say("Sign in with your Google account to open the tool.")

    # ── layout ────────────────────────────────────────────────────────────
    def _build(self) -> None:
        a, w = self.app, self.win
        ttk.Label(w, text="📝  Paper Maker", style="Title.TLabel").pack(
            anchor="w", padx=24, pady=(20, 4))
        self.message_var = tk.StringVar()
        self.message = tk.Label(
            w, textvariable=self.message_var, bg=w["bg"], fg="#9ca3af",
            font=a.font_small, wraplength=360, justify="left", anchor="w", height=3)
        self.message.pack(fill="x", padx=24, pady=(4, 6))
        self.button = ttk.Button(w, text="Sign in with Google", style="Accent.TButton",
                                 command=self._on_button, width=44)
        self.button.pack(fill="x", padx=24, pady=(0, 20))

    def _to_front(self) -> None:
        # Windows will not let a window take the focus from another program;
        # being "topmost" for a moment still brings it in front.
        self.win.attributes("-topmost", True)
        self.win.lift()
        self.win.focus_force()
        if self._front_after is not None:
            self.win.after_cancel(self._front_after)
        self._front_after = self.win.after(300, self._not_topmost)

    def _not_topmost(self) -> None:
        self._front_after = None
        self.win.attributes("-topmost", False)

    def _close(self) -> None:
        for job in (self._after, self._front_after):
            if job is not None:
                self.win.after_cancel(job)
        self._after = self._front_after = None
        self.win.destroy()

    def _say(self, text: str, colour: str = "#9ca3af") -> None:
        self.message_var.set(text)
        self.message.configure(fg=colour)

    # ── check ─────────────────────────────────────────────────────────────
    def _on_button(self) -> None:
        if self.checking:  # the button reads "Cancel" while the browser is open
            self.stop.set()
            return
        self.stop.clear()
        self._start(self._sign_in_worker, "Finish the sign-in in your browser…", cancel=True)

    def _start(self, worker, text: str, cancel: bool = False) -> None:
        self.checking = True
        self._say(text)
        if cancel:
            self.button.configure(text="Cancel")
        else:
            self.button.state(["disabled"])
        threading.Thread(target=worker, daemon=True).start()
        self._after = self.win.after(60, self._poll)

    def _recheck_worker(self) -> None:
        self._run(lambda: (refresh_sign_in(self.app._config[TOKEN_KEY]), None))

    def _sign_in_worker(self) -> None:
        self._run(lambda: browser_sign_in(self.stop))

    def _run(self, get_tokens) -> None:
        try:
            id_token, refresh_token = get_tokens()
            email, allowed = ask_service(id_token)
            self.results.put(("answer", email, allowed, refresh_token))
        except SignInNeeded:
            self.results.put(("sign_in", None, False, None))
        except Cancelled:
            self.results.put(("cancelled", None, False, None))
        except AccessError as exc:
            self.results.put(("error", str(exc), False, None))

    def _poll(self) -> None:
        try:
            kind, detail, allowed, refresh_token = self.results.get_nowait()
        except queue.Empty:
            self._after = self.win.after(60, self._poll)
            return
        self._after = None
        self.checking = False
        self.button.state(["!disabled"])
        self.button.configure(text="Sign in with Google", command=self._on_button)
        self._to_front()  # back from the browser
        if kind == "answer" and refresh_token:
            self.app._config[TOKEN_KEY] = refresh_token
            self.app._config.pop("access_email", None)  # from the type-your-email version
            self.app._save_config()
        if kind == "answer" and allowed:
            if self.app._config.get(EMAIL_KEY) != detail:
                self.app._config[EMAIL_KEY] = detail
                self.app._save_config()
            self.allowed = True
            self._close()
        elif kind == "answer":
            self._say(f"{detail} does not have access. Ask the owner of this tool to add "
                      "your email, or sign in with another account.", "#ef4444")
            self.button.configure(text="Sign in with another account")
        elif kind == "sign_in":
            if self.app._config.pop(TOKEN_KEY, None):
                self.app._save_config()
            self._say("Your sign-in has expired. Sign in with Google again.", "#f59e0b")
        elif kind == "cancelled":
            self._say("Sign-in cancelled. Sign in with your Google account to open the tool.")
        else:
            self._say(detail, "#f59e0b")
            if self.app._config.get(TOKEN_KEY):
                self.button.configure(text="Try again")
                self.button.configure(command=self._on_retry)

    def _on_retry(self) -> None:
        self._start(self._recheck_worker, "Checking access…")

    def _on_close(self) -> None:
        self.stop.set()
        self._close()


def sign_out(app) -> None:
    """Forget the Google sign-in on this PC and show the sign-in window again.

    Deliberately no revoke at Google: revoking cancels EVERY grant the account gave
    the Paper Maker client, including the owner's sheet token the access service
    reads the list with, so the owner signing out locked everyone out (HTTP 503).
    """
    who = app._config.get(EMAIL_KEY) or "your account"
    if not messagebox.askyesno("Sign out", f"Sign out {who} from Paper Maker on this PC?\n\n"
                               "Paper Maker will ask for a Google sign-in again."):
        return
    app._config.pop(TOKEN_KEY, None)
    app._config.pop(EMAIL_KEY, None)
    app._save_config()
    if not ensure_access(app):
        app.root.destroy()


def ensure_access(app) -> bool:
    """Sign in over the hidden main window; True once the access list allows the user."""
    app.root.withdraw()
    gate = AccessWindow(app)
    app.root.wait_window(gate.win)
    if gate.allowed:
        app.root.deiconify()
    return gate.allowed
