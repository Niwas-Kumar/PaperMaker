"""One-time: give the access service the owner's read-only Google token.

Opens the browser for the owner's Google sign-in (permission: see Google Sheets,
read only), checks that column A of the access sheet can be read with it, and
stores the token as the Worker secret OWNER_REFRESH_TOKEN. The token is not
printed and not written to disk. Also stores GOOGLE_CLIENT_SECRET.

    python access_service\\set_owner_token.py      (after: npx wrangler login)

Run it again if the owner's token is ever revoked (the service then answers 503).
"""

import base64
import hashlib
import json
import secrets
import subprocess
import sys
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
client = json.loads((HERE.parent / "oauth_client.json").read_text(encoding="utf-8"))["installed"]
toml = (HERE / "wrangler.toml").read_text(encoding="utf-8")
SHEET_ID = toml.split('SHEET_ID = "')[1].split('"')[0]
SHEET_RANGE = toml.split('SHEET_RANGE = "')[1].split('"')[0]

verifier = secrets.token_urlsafe(64)
challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
state = secrets.token_urlsafe(16)
answer = {}


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
        self.wfile.write(b"<h2>Done. You can close this tab.</h2>")


server = HTTPServer(("127.0.0.1", 0), Handler)
redirect = f"http://127.0.0.1:{server.server_address[1]}/"
print("Sign in in the browser with the account that can open the access sheet ...")
webbrowser.open(client["auth_uri"] + "?" + urllib.parse.urlencode({
    "client_id": client["client_id"], "redirect_uri": redirect, "response_type": "code",
    "scope": "https://www.googleapis.com/auth/spreadsheets.readonly",
    "code_challenge": challenge, "code_challenge_method": "S256", "state": state,
    "access_type": "offline", "prompt": "consent select_account",
}))
while "code" not in answer and "error" not in answer:
    server.handle_request()
server.server_close()
if "error" in answer:
    sys.exit("Sign-in failed: " + answer["error"])

with urllib.request.urlopen(client["token_uri"], urllib.parse.urlencode({
    "grant_type": "authorization_code", "code": answer["code"], "redirect_uri": redirect,
    "code_verifier": verifier, "client_id": client["client_id"],
    "client_secret": client["client_secret"],
}).encode(), timeout=30) as resp:
    tokens = json.load(resp)
if "refresh_token" not in tokens:
    sys.exit("Google gave no refresh token; run again and allow the permission.")

req = urllib.request.Request(
    f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{urllib.parse.quote(SHEET_RANGE)}",
    headers={"Authorization": "Bearer " + tokens["access_token"]})
with urllib.request.urlopen(req, timeout=30) as resp:
    rows = json.load(resp).get("values", [])
count = sum(1 for row in rows for cell in row if "@" in str(cell))
print(f"Sheet read OK: {count} emails in {SHEET_RANGE}.")


def put_secret(name: str, value: str) -> None:
    done = subprocess.run(f"npx wrangler secret put {name}", cwd=HERE, input=value,
                          text=True, encoding="utf-8", errors="replace", shell=True,
                          capture_output=True)  # wrangler prints emoji
    if done.returncode:
        sys.exit(f"Storing {name} failed:\n{done.stdout}\n{done.stderr}")
    print(f"Stored secret {name}.")


put_secret("GOOGLE_CLIENT_SECRET", client["client_secret"])
put_secret("OWNER_REFRESH_TOKEN", tokens["refresh_token"])
