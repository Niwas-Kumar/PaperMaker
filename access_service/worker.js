/**
 * Paper Maker access check (Cloudflare Worker).
 *
 * POST /check  {"id_token": "<Google ID token from the desktop sign-in>"}
 *   -> 200 {"allowed": true|false, "email": "..."}
 *   -> 401 {"error": "..."}   token missing, expired or not issued to Paper Maker
 *   -> 503 {"error": "..."}   the sheet could not be read
 *
 * The allowed emails are read from SHEET_RANGE of SHEET_ID with the owner's
 * read-only Google token (secret OWNER_REFRESH_TOKEN). Only yes/no leaves the
 * sheet. Settings: wrangler.toml [vars]; secrets: GOOGLE_CLIENT_SECRET,
 * OWNER_REFRESH_TOKEN (wrangler secret put ...).
 */

const LIST_TTL_MS = 60 * 1000; // a removed email stops working within a minute

let cachedAccess = null; // {token, until}
let cachedList = null; // {emails: Set, until}

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
  });
}

// Who is signing in: Google checks the token's signature and expiry for us.
async function verifiedEmail(idToken, env) {
  const resp = await fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken));
  if (!resp.ok) return null;
  const info = await resp.json();
  const issuerOk = info.iss === 'https://accounts.google.com' || info.iss === 'accounts.google.com';
  if (!issuerOk || info.aud !== env.GOOGLE_CLIENT_ID) return null; // token made for another app
  if (String(info.email_verified) !== 'true' || !info.email) return null;
  return info.email.trim().toLowerCase();
}

async function ownerAccessToken(env) {
  if (cachedAccess && cachedAccess.until > Date.now()) return cachedAccess.token;
  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: env.OWNER_REFRESH_TOKEN,
      client_id: env.GOOGLE_CLIENT_ID,
      client_secret: env.GOOGLE_CLIENT_SECRET,
    }),
  });
  if (!resp.ok) throw new Error('owner token refresh failed: HTTP ' + resp.status + ' ' + (await resp.text()).slice(0, 200));
  const tok = await resp.json();
  cachedAccess = { token: tok.access_token, until: Date.now() + (tok.expires_in - 120) * 1000 };
  return cachedAccess.token;
}

async function allowedEmails(env) {
  if (cachedList && cachedList.until > Date.now()) return cachedList.emails;
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + env.SHEET_ID +
      '/values/' + encodeURIComponent(env.SHEET_RANGE) + '?majorDimension=ROWS';
  const resp = await fetch(url, { headers: { Authorization: 'Bearer ' + (await ownerAccessToken(env)) } });
  if (!resp.ok) throw new Error('sheet read failed: HTTP ' + resp.status + ' ' + (await resp.text()).slice(0, 200));
  const rows = (await resp.json()).values || [];
  const emails = new Set();
  for (const row of rows) {
    for (const cell of row) {
      const v = String(cell).trim().toLowerCase();
      if (v.includes('@')) emails.add(v);
    }
  }
  cachedList = { emails, until: Date.now() + LIST_TTL_MS };
  return emails;
}

// Home page and privacy policy for Google's sign-in screen (Google Auth Platform > Branding).
const PAGE_STYLE = '<meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
    'body{font-family:Segoe UI,Arial,sans-serif;max-width:640px;margin:48px auto;padding:0 16px;' +
    'line-height:1.6;color:#1f2937}h1{font-size:1.6em}</style>';
const PAGES = {
  '/': '<!doctype html><title>Paper Maker</title>' + PAGE_STYLE +
      '<h1>Paper Maker</h1><p>Paper Maker is an internal Windows tool of the PW content team for ' +
      'building question papers from Word and PDF files.</p><p>Only team members whose email the ' +
      'owner has put on the access list can use it. To get access, ask the owner of the tool.</p>' +
      '<p><a href="/privacy">Privacy policy</a></p>',
  '/privacy': '<!doctype html><title>Paper Maker privacy policy</title>' + PAGE_STYLE +
      '<h1>Paper Maker privacy policy</h1>' +
      '<p>When you sign in to Paper Maker with Google, the tool receives only your email address. ' +
      'It uses it for one thing: to check whether that email is on the access list, and to open ' +
      'the tool if it is.</p>' +
      '<p>Your email is not stored by the access service, not shared with anyone and not used for ' +
      'anything else. On your own computer Paper Maker remembers the sign-in so that you do not ' +
      'have to sign in at every start; you can remove it by revoking "Paper Maker" at ' +
      '<a href="https://myaccount.google.com/permissions">myaccount.google.com/permissions</a>.</p>' +
      '<p>The papers you open and create stay on your computer; they are never uploaded by the sign-in.</p>' +
      '<p>Questions: ask the owner of the tool.</p>',
};

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (request.method === 'GET' && PAGES[path]) {
      return new Response(PAGES[path], { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }
    if (request.method !== 'POST' || path !== '/check') return json(404, { error: 'not found' });

    let idToken = '';
    try {
      idToken = String((await request.json()).id_token || '');
    } catch (e) {
      return json(400, { error: 'body must be JSON with id_token' });
    }
    if (!idToken) return json(401, { error: 'not signed in' });

    const email = await verifiedEmail(idToken, env);
    if (!email) return json(401, { error: 'sign-in expired or not valid for Paper Maker' });

    try {
      const emails = await allowedEmails(env);
      return json(200, { allowed: emails.has(email), email });
    } catch (e) {
      console.log(String(e));
      return json(503, { error: 'the access list could not be read' });
    }
  },
};
