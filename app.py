#!/usr/bin/env python3
import base64
import hashlib
import hmac
import html
import imaplib
import json
import os
import re
import secrets
import sqlite3
import time
import urllib.parse
import urllib.request
from datetime import datetime
from email import message_from_bytes
from email.header import decode_header, make_header
from http import HTTPStatus
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get('RTWEB_DB', BASE_DIR / 'app.db'))
ADMIN_PASSWORD = os.environ.get('RTWEB_ADMIN_PASSWORD') or os.environ.get('RTWEB_PASSWORD', 'change-me')
SESSION_SECRET = os.environ.get('RTWEB_SESSION_SECRET', secrets.token_hex(32))
BIND_HOST = os.environ.get('RTWEB_HOST', '127.0.0.1')
BIND_PORT = int(os.environ.get('RTWEB_PORT', '8020'))
TOKEN_URL_TEMPLATE = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'
ME_URL = 'https://graph.microsoft.com/v1.0/me'
MAIL_FOLDERS = [('INBOX', '收件箱'), ('Junk', '垃圾邮箱')]
GRAPH_MAIL_FOLDERS = [('inbox', '收件箱'), ('junkemail', '垃圾邮箱')]

def default_oauth_scope() -> str:
    # Official Exchange OAuth docs list these full Outlook resource scopes for IMAP/SMTP.
    return 'offline_access https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send'


def normalize_tenant(value: str) -> str:
    value = (value or 'consumers').strip()
    allowed = {'consumers', 'common', 'organizations'}
    if value in allowed:
        return value
    if value and not value.startswith('http') and all(c.isalnum() or c in '-_.' for c in value):
        return value
    return 'consumers'


def mask_secret(value: str) -> str:
    if not value or len(value) < 12:
        return '***'
    return value[:4] + '…' + value[-4:]


def secret_digest(value: str) -> str:
    return hashlib.sha256(value.encode('utf-8')).hexdigest()


def check_password(given: str, expected: str) -> bool:
    return hmac.compare_digest((given or '').encode(), (expected or '').encode())


def sign_session(ts: str) -> str:
    return hmac.new(SESSION_SECRET.encode(), ts.encode(), hashlib.sha256).hexdigest()


def make_session_cookie() -> str:
    ts = str(int(time.time()))
    secure = '; Secure' if os.environ.get('RTWEB_COOKIE_SECURE', '1') != '0' else ''
    return f'rtweb={ts}.{sign_session(ts)}; HttpOnly{secure}; SameSite=Lax; Path=/; Max-Age=604800'


def parse_cookies(cookie_header: str):
    cookies = {}
    for part in (cookie_header or '').split(';'):
        if '=' in part:
            k, v = part.strip().split('=', 1)
            cookies[k] = v
    return cookies


def make_active_account_cookie(account_id: str) -> str:
    secure = '; Secure' if os.environ.get('RTWEB_COOKIE_SECURE', '1') != '0' else ''
    return f'active_account_id={urllib.parse.quote(str(account_id))}; HttpOnly{secure}; SameSite=Lax; Path=/; Max-Age=604800'


def get_active_account_id(cookie_header: str) -> str:
    raw = parse_cookies(cookie_header).get('active_account_id', '')
    return urllib.parse.unquote(raw) if raw.isdigit() else ''


def verify_session(cookie_header: str) -> bool:
    if not cookie_header:
        return False
    cookies = parse_cookies(cookie_header)
    raw = cookies.get('rtweb', '')
    if '.' not in raw:
        return False
    ts, sig = raw.split('.', 1)
    if not ts.isdigit():
        return False
    if int(time.time()) - int(ts) > 7 * 24 * 3600:
        return False
    return hmac.compare_digest(sig, sign_session(ts))


def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS decode_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                client_id TEXT NOT NULL,
                token_digest TEXT NOT NULL,
                token_mask TEXT NOT NULL,
                status TEXT NOT NULL,
                account TEXT,
                scope TEXT,
                expires_in INTEGER,
                error TEXT
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS saved_accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                email TEXT NOT NULL UNIQUE,
                password TEXT,
                password_mask TEXT,
                client_id TEXT NOT NULL,
                refresh_token TEXT NOT NULL,
                token_digest TEXT NOT NULL,
                token_mask TEXT NOT NULL,
                last_status TEXT,
                last_error TEXT
            )
        ''')
        cols = [r[1] for r in conn.execute('PRAGMA table_info(saved_accounts)')]
        if 'password' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN password TEXT')
        if 'password_mask' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN password_mask TEXT')
        conn.commit()


def insert_record(client_id, refresh_token, status, account=None, scope=None, expires_in=None, error=None):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO decode_records
            (created_at, client_id, token_digest, token_mask, status, account, scope, expires_in, error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (int(time.time()), client_id, secret_digest(refresh_token), mask_secret(refresh_token), status, account, scope, expires_in, error))
        conn.commit()


def recent_records(limit=20):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return list(conn.execute('SELECT * FROM decode_records ORDER BY id DESC LIMIT ?', (limit,)))


def parse_batch_accounts(text: str):
    rows = []
    for line in (text or '').splitlines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = None
        for sep in ('----', '|', ',', '\t'):
            candidate = [p.strip() for p in line.split(sep)]
            if len(candidate) >= 3:
                parts = candidate
                break
        if not parts:
            continue
        if len(parts) >= 4:
            email_addr, password, client_id, refresh_token = parts[0], parts[1], parts[2], parts[3]
        else:
            email_addr, password, client_id, refresh_token = parts[0], '', parts[1], parts[2]
        if '@' not in email_addr or not client_id or not refresh_token:
            continue
        rows.append({'email': email_addr, 'password': password, 'client_id': client_id, 'refresh_token': refresh_token})
    return rows


def save_account(email_addr: str, client_id: str, refresh_token: str, password: str = ''):
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO saved_accounts
            (created_at, updated_at, email, password, password_mask, client_id, refresh_token, token_digest, token_mask, last_status, last_error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL)
            ON CONFLICT(email) DO UPDATE SET
                updated_at=excluded.updated_at,
                password=CASE WHEN excluded.password != '' THEN excluded.password ELSE saved_accounts.password END,
                password_mask=CASE WHEN excluded.password_mask != '' THEN excluded.password_mask ELSE saved_accounts.password_mask END,
                client_id=excluded.client_id,
                refresh_token=excluded.refresh_token,
                token_digest=excluded.token_digest,
                token_mask=excluded.token_mask
        ''', (now, now, email_addr, password, mask_secret(password) if password else '', client_id, refresh_token, secret_digest(refresh_token), mask_secret(refresh_token)))
        conn.commit()


def update_saved_status(email_addr: str, status: str, error: str = ''):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('UPDATE saved_accounts SET last_status=?, last_error=?, updated_at=? WHERE email=?', (status, error[:1000], int(time.time()), email_addr))
        conn.commit()


def saved_accounts():
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return list(conn.execute('SELECT * FROM saved_accounts ORDER BY updated_at DESC'))


def get_saved_account(account_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return conn.execute('SELECT * FROM saved_accounts WHERE id=?', (account_id,)).fetchone()


def delete_saved_account(account_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('DELETE FROM saved_accounts WHERE id=?', (account_id,))
        conn.commit()


def exchange_refresh_token(client_id: str, refresh_token: str, scope: str = '', tenant: str = 'consumers'):
    fields = {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
    }
    scope = (scope or '').strip()
    if scope:
        fields['scope'] = scope
    data = urllib.parse.urlencode(fields).encode('utf-8')
    token_url = TOKEN_URL_TEMPLATE.format(tenant=normalize_tenant(tenant))
    req = urllib.request.Request(token_url, data=data, headers={'Content-Type': 'application/x-www-form-urlencoded'})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
            return True, payload
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        try:
            payload = json.loads(body)
        except Exception:
            payload = {'error': f'HTTP {e.code}', 'error_description': body[:500]}
        return False, payload
    except Exception as e:
        return False, {'error': type(e).__name__, 'error_description': str(e)}


def fetch_me(access_token: str):
    req = urllib.request.Request(ME_URL, headers={'Authorization': f'Bearer {access_token}'})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except Exception:
        return {}


def is_imap_not_connected_error(err: str) -> bool:
    return 'User is authenticated but not connected' in (err or '')


def graph_message_to_mail(msg: dict):
    sender = (((msg.get('from') or {}).get('emailAddress')) or {})
    name = sender.get('name') or ''
    addr = sender.get('address') or ''
    source = (name + (' <' + addr + '>' if addr else '')).strip() or addr
    return {
        'from': source,
        'subject': msg.get('subject') or '',
        'date': msg.get('receivedDateTime') or msg.get('sentDateTime') or '',
        'preview': msg.get('bodyPreview') or '',
    }


def fetch_graph_latest_emails(access_token: str, limit: int = 10):
    results = []
    for folder_id, folder_label in GRAPH_MAIL_FOLDERS:
        url = f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?' + urllib.parse.urlencode({
            '$top': str(limit),
            '$orderby': 'receivedDateTime desc',
            '$select': 'receivedDateTime,sentDateTime,subject,bodyPreview,from'
        })
        req = urllib.request.Request(url, headers={'Authorization': f'Bearer {access_token}', 'Accept': 'application/json'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
        for m in payload.get('value', []):
            mail = graph_message_to_mail(m)
            mail['folder'] = folder_label
            results.append(mail)
    return results[:limit * len(GRAPH_MAIL_FOLDERS)]


def build_xoauth2_payload(username: str, access_token: str) -> str:
    # imaplib.authenticate() base64-encodes the returned value itself.
    # Return the raw SASL XOAUTH2 string to avoid double-encoding.
    return f'user={username}\x01auth=Bearer {access_token}\x01\x01'


def decode_mime(value) -> str:
    if not value:
        return ''
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def html_to_text_preview(value: str, limit: int = 500) -> str:
    value = re.sub(r'<(br|/p|/div|/tr)\b[^>]*>', ' ', value or '', flags=re.I)
    value = re.sub(r'<[^>]+>', ' ', value)
    value = html.unescape(value)
    value = re.sub(r'\s+', ' ', value).strip()
    return value[:limit]


def message_preview(msg, limit: int = 500) -> str:
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get('Content-Disposition') or '').lower()
            if 'attachment' in disp:
                continue
            if ctype in ('text/plain', 'text/html'):
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    text = payload.decode(charset, errors='replace')
                    if ctype == 'text/html':
                        text = html_to_text_preview(text, limit)
                    parts.append(text)
                    if ctype == 'text/plain':
                        break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or 'utf-8'
            text = payload.decode(charset, errors='replace')
            if msg.get_content_type() == 'text/html':
                text = html_to_text_preview(text, limit)
            parts.append(text)
    return html_to_text_preview('\n'.join(parts), limit)


def extract_verification_summary(email_addr: str, mail: dict):
    text = ' '.join([mail.get('subject', ''), mail.get('preview', '')])
    patterns = [
        r'(?<!\d)(\d{4,8})(?!\d)',
        r'code[^0-9]{0,20}(\d{4,8})',
        r'验证码[^0-9]{0,20}(\d{4,8})',
    ]
    code = None
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            code = m.group(1)
            break
    if not code:
        return None
    source = mail.get('from', '') or mail.get('subject', '') or '未知来源'
    return {
        'email': email_addr,
        'source': source,
        'code': code,
        'date': mail.get('date', ''),
        'subject': mail.get('subject', ''),
        'folder': mail.get('folder', ''),
    }


def fetch_latest_codes_for_account(account, limit: int = 10):
    graph_scope = 'offline_access https://graph.microsoft.com/Mail.Read'
    gok, gpayload = exchange_refresh_token(account['client_id'], account['refresh_token'], scope=graph_scope, tenant='consumers')
    graph_error = ''
    if gok:
        try:
            mails = fetch_graph_latest_emails(gpayload.get('access_token'), limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='graph') for c in codes if c]
            if codes:
                update_saved_status(account['email'], 'ok')
                return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'graph', 'codes': codes}
        except Exception as e:
            graph_error = str(e)
    else:
        graph_error = gpayload.get('error_description') or gpayload.get('error') or ''

    ok, payload = exchange_refresh_token(account['client_id'], account['refresh_token'], scope=default_oauth_scope(), tenant='consumers')
    if not ok:
        err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
        update_saved_status(account['email'], 'error', err)
        return {'email': account['email'], 'status': 'error', 'error': err, 'codes': []}
    try:
        mails = fetch_latest_emails(account['email'], payload.get('access_token'), limit=limit)
        codes = [extract_verification_summary(account['email'], m) for m in mails]
        codes = [dict(c, source_api='imap') for c in codes if c]
        update_saved_status(account['email'], 'ok')
        return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'imap', 'graph_error': graph_error, 'codes': codes}
    except Exception as e:
        err = str(e)
        if is_imap_not_connected_error(err):
            err = 'IMAP 已认证但邮箱未连接：通常是该账号未开通/未初始化 Outlook 邮箱，或 Microsoft 账户没有可连接的 Exchange mailbox。Graph 兜底错误：' + graph_error
        update_saved_status(account['email'], 'error', err)
        return {'email': account['email'], 'status': 'error', 'error': err, 'codes': []}


def fetch_latest_emails(username: str, access_token: str, limit: int = 10):
    imap = imaplib.IMAP4_SSL('outlook.office365.com', 993)
    try:
        imap.authenticate('XOAUTH2', lambda _: build_xoauth2_payload(username, access_token))
        results = []
        for folder_name, folder_label in MAIL_FOLDERS:
            typ, _ = imap.select(folder_name, readonly=True)
            if typ != 'OK':
                continue
            typ, data = imap.search(None, 'ALL')
            if typ != 'OK' or not data or not data[0]:
                continue
            ids = data[0].split()[-limit:][::-1]
            for mid in ids:
                typ, msg_data = imap.fetch(mid, '(RFC822)')
                if typ != 'OK' or not msg_data:
                    continue
                raw = next((item[1] for item in msg_data if isinstance(item, tuple)), None)
                if not raw:
                    continue
                msg = message_from_bytes(raw)
                results.append({
                    'from': decode_mime(msg.get('From')),
                    'subject': decode_mime(msg.get('Subject')),
                    'date': decode_mime(msg.get('Date')),
                    'preview': message_preview(msg),
                    'folder': folder_label,
                })
        return results
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def page(title, body):
    return f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title><style>
:root{{color-scheme:dark}}body{{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#0f172a;color:#e5e7eb;font-size:13px}}.wrap{{max-width:1440px;margin:0 auto;padding:10px}}.card{{background:#111827;border:1px solid #263244;border-radius:12px;padding:12px;margin:0;box-shadow:0 8px 18px #0003;min-height:0}}input,textarea{{width:100%;box-sizing:border-box;background:#020617;color:#e5e7eb;border:1px solid #334155;border-radius:8px;padding:8px;margin:4px 0 8px;font:inherit}}textarea{{min-height:70px;font-family:ui-monospace,monospace}}button{{background:#2563eb;color:white;border:0;border-radius:8px;padding:8px 12px;font-weight:700;cursor:pointer}}button:hover{{background:#1d4ed8}}h2,h3{{margin:0 0 8px}}.muted{{color:#94a3b8}}.ok{{color:#86efac}}.bad{{color:#fca5a5}}code,pre{{background:#020617;border:1px solid #334155;border-radius:8px;padding:8px;display:block;overflow:auto}}table{{width:100%;border-collapse:collapse}}td,th{{border-bottom:1px solid #263244;padding:5px;text-align:left;font-size:12px;vertical-align:top}}a{{color:#93c5fd}}.top{{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:8px}}.grid{{display:grid;grid-template-columns:360px 360px 1fr;grid-template-rows:auto 1fr;gap:10px;height:calc(100vh - 62px)}}.span2{{grid-column:1 / span 2}}.codes{{grid-row:1 / span 2;grid-column:3;overflow:auto}}.scroll{{overflow:auto}}@media(max-width:900px){{.grid{{display:block;height:auto}}.card{{margin-bottom:10px}}}}
</style></head><body><div class="wrap"><div class="top"><h2>🎟️ Refresh Token 登录解码</h2><a href="/logout">退出</a></div>{body}</div></body></html>'''


def render_home_page(active_account_id: str = ''):
    accounts = saved_accounts()
    active_email = ''
    account_rows = ''.join(
        f'<tr class="{"ok" if str(a["id"]) == str(active_account_id) else ""}"><td>{a["id"]}</td><td>{html.escape(a["email"])}</td><td>{html.escape(a["password_mask"] or "")}</td><td>{html.escape(a["client_id"])}</td><td>{html.escape(a["token_mask"])}</td><td>{html.escape(a["last_status"] or "")}</td><td>{html.escape((a["last_error"] or "")[:80])}</td><td><a href="/select?id={a["id"]}">登录/选中</a><form method="post" action="/delete" style="display:inline" onsubmit="return confirm(\'删除这个账号？\')"><input type="hidden" name="id" value="{a["id"]}"><button type="submit" style="padding:3px 6px;margin-left:6px;background:#dc2626">删除</button></form></td></tr>'
        for a in accounts
    )
    for a in accounts:
        if str(a['id']) == str(active_account_id):
            active_email = a['email']
            break
    body = f'''
<div class="grid">
<div class="card"><h3>保存账号</h3><form method="post" action="/decode">
<label>Client ID</label><input name="client_id" placeholder="9e5f94bc-e8a4-4e73-b8be-63364c29d753" required>
<label>邮箱地址</label><input name="email" placeholder="name@hotmail.com / name@outlook.com" required>
<label>Refresh Token</label><textarea name="refresh_token" placeholder="粘贴 refresh token" required></textarea><button>保存账号</button></form></div>
<div class="card"><h3>批量导入</h3><p class="muted">格式：邮箱----密码----ClientID----RefreshToken</p><form method="post" action="/batch"><textarea name="batch" placeholder="a@hotmail.com----password----client_id----refresh_token"></textarea><button>批量保存</button></form></div>
<div class="card codes"><h3>最新验证码邮件</h3><p class="muted">当前邮箱：<b id="active-email">{html.escape(active_email or '未选择')}</b> · 自动刷新：<span id="poll-status">等待中</span></p><div id="codes">加载中...</div></div>
<div class="card span2 scroll"><h3>已保存账号</h3><table><tr><th>ID</th><th>邮箱</th><th>密码</th><th>Client ID</th><th>Token</th><th>状态</th><th>错误</th><th>操作</th></tr>{account_rows}</table></div>
</div>
<script>
async function pollCodes() {{
  const box = document.getElementById('codes');
  const status = document.getElementById('poll-status');
  try {{
    status.textContent = '获取中 ' + new Date().toLocaleTimeString();
    const res = await fetch('/api/codes', {{cache: 'no-store'}});
    const data = await res.json();
    let html = '<table><tr><th>登入邮箱</th><th>邮箱位置</th><th>来源</th><th>验证码</th><th>时间</th><th>主题</th></tr>';
    for (const item of data.results || []) {{
      if (item.status !== 'ok') {{
        html += `<tr><td>${{escapeHtml(item.email)}}</td><td colspan="5" class="bad">${{escapeHtml(item.error || '读取失败')}}</td></tr>`;
        continue;
      }}
      if (!item.codes.length) {{
        html += `<tr><td>${{escapeHtml(item.email)}}</td><td colspan="5" class="muted">暂无验证码邮件</td></tr>`;
      }}
      for (const c of item.codes) {{
        html += `<tr><td>${{escapeHtml(c.email)}}</td><td>${{escapeHtml(c.folder || '')}}</td><td>${{escapeHtml(c.source)}}</td><td><b>${{escapeHtml(c.code)}}</b></td><td>${{escapeHtml(c.date)}}</td><td>${{escapeHtml(c.subject)}}</td></tr>`;
      }}
    }}
    html += '</table>';
    box.innerHTML = html;
    status.textContent = '已更新 ' + new Date().toLocaleTimeString();
  }} catch (e) {{
    status.textContent = '失败 ' + e;
  }}
}}
function escapeHtml(s) {{ return String(s || '').replace(/[&<>"]/g, c => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}}[c])); }}
pollCodes(); setInterval(pollCodes, 30000);
</script>'''
    return page('Refresh Token 登录解码', body)


class Handler(BaseHTTPRequestHandler):
    server_version = 'RTWeb/1.0'

    def log_message(self, fmt, *args):
        safe = fmt % args
        safe = safe.replace('refresh_token', 'refresh_token')
        print(f'{self.address_string()} - {safe}')

    def read_form(self):
        length = int(self.headers.get('Content-Length', '0') or '0')
        raw = self.rfile.read(length).decode('utf-8')
        return {k: v[0] if v else '' for k, v in urllib.parse.parse_qs(raw).items()}

    def send_html(self, body, status=200, headers=None):
        data = body.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, path, headers=None):
        self.send_response(302)
        self.send_header('Location', path)
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.end_headers()

    def require_auth(self):
        return verify_session(self.headers.get('Cookie', ''))

    def do_GET(self):
        if self.path.startswith('/health'):
            self.send_response(200); self.end_headers(); self.wfile.write(b'ok'); return
        if self.path.startswith('/logout'):
            self.redirect('/login', {'Set-Cookie': 'rtweb=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0'}); return
        if self.path.startswith('/login'):
            self.send_html(page('登录', '<div class="card"><form method="post" action="/login"><label>管理密码</label><input type="password" name="password" autofocus><button>登录</button></form></div>'))
            return
        if not self.require_auth():
            self.redirect('/login'); return
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/api/codes':
            self.send_codes_json()
            return
        if parsed.path == '/select':
            account_id = urllib.parse.parse_qs(parsed.query).get('id', [''])[0]
            self.redirect('/', {'Set-Cookie': make_active_account_cookie(account_id)})
            return
        if parsed.path == '/read':
            account_id = urllib.parse.parse_qs(parsed.query).get('id', [''])[0]
            self.read_saved_account(account_id)
            return
        self.send_html(render_home_page(get_active_account_id(self.headers.get('Cookie', ''))))

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(data)

    def send_codes_json(self):
        active_id = get_active_account_id(self.headers.get('Cookie', ''))
        results = []
        if active_id:
            account = get_saved_account(active_id)
            if account:
                results = [fetch_latest_codes_for_account(account, limit=10)]
        self.send_json({'results': results, 'active_id': active_id, 'ts': int(time.time())})

    def render_mail_result(self, account, payload):
        shown = {k: v for k, v in payload.items() if k not in ('access_token', 'refresh_token', 'id_token')}
        shown['account'] = account
        result = '<p class="ok">成功换取 access token。</p><pre>' + html.escape(json.dumps(shown, ensure_ascii=False, indent=2)) + '</pre>'
        try:
            mails = fetch_latest_emails(account, payload.get('access_token'), limit=10)
            rows = ''.join(
                '<tr><td>' + html.escape(m.get('date','')) + '</td><td>' + html.escape(m.get('from','')) + '</td><td>' + html.escape(m.get('subject','')) + '</td><td>' + html.escape(m.get('preview','')) + '</td></tr>'
                for m in mails
            )
            result += '<h3>最新邮件</h3><table><tr><th>时间</th><th>发件人</th><th>主题</th><th>预览</th></tr>' + rows + '</table>'
        except Exception as e:
            result += '<p class="bad">邮件读取失败：' + html.escape(str(e)) + '</p>'
        return result

    def read_saved_account(self, account_id):
        account = get_saved_account(account_id)
        if not account:
            self.send_html(page('不存在', '<div class="card"><p class="bad">账号不存在。</p><a href="/">返回</a></div>'), 404)
            return
        ok, payload = exchange_refresh_token(account['client_id'], account['refresh_token'], scope=default_oauth_scope(), tenant='consumers')
        if ok:
            insert_record(account['client_id'], account['refresh_token'], 'ok', account=account['email'], scope=payload.get('scope'), expires_in=payload.get('expires_in'))
            update_saved_status(account['email'], 'ok')
            result = self.render_mail_result(account['email'], payload)
        else:
            err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
            insert_record(account['client_id'], account['refresh_token'], 'error', account=account['email'], error=err)
            update_saved_status(account['email'], 'error', err)
            result = '<p class="bad">换取失败。</p><pre>' + html.escape(json.dumps(payload, ensure_ascii=False, indent=2)) + '</pre>'
        self.send_html(page('结果', f'<div class="card">{result}<p><a href="/">返回</a></p></div>'))

    def do_POST(self):
        if self.path == '/login':
            form = self.read_form()
            if check_password(form.get('password', ''), ADMIN_PASSWORD):
                self.redirect('/', {'Set-Cookie': make_session_cookie()})
            else:
                self.send_html(page('登录失败', '<div class="card"><p class="bad">密码错误</p><a href="/login">返回</a></div>'), 403)
            return
        if not self.require_auth():
            self.redirect('/login'); return
        if self.path == '/delete':
            form = self.read_form()
            delete_saved_account(form.get('id', ''))
            self.redirect('/')
            return
        if self.path == '/batch':
            form = self.read_form()
            rows = parse_batch_accounts(form.get('batch', ''))
            for row in rows:
                save_account(row['email'], row['client_id'], row['refresh_token'], row.get('password', ''))
            self.redirect('/?saved=' + str(len(rows)))
            return
        if self.path == '/decode':
            form = self.read_form()
            client_id = form.get('client_id', '').strip()
            refresh_token = form.get('refresh_token', '').strip()
            email_addr = form.get('email', '').strip()
            scope = default_oauth_scope()
            tenant = 'consumers'
            save_account(email_addr, client_id, refresh_token)
            ok, payload = exchange_refresh_token(client_id, refresh_token, scope=scope, tenant=tenant)
            if ok:
                insert_record(client_id, refresh_token, 'ok', account=email_addr, scope=payload.get('scope'), expires_in=payload.get('expires_in'))
                update_saved_status(email_addr, 'ok')
            else:
                err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
                insert_record(client_id, refresh_token, 'error', account=email_addr, error=err)
                update_saved_status(email_addr, 'error', err)
            self.redirect('/')
            return
        self.send_error(404)


def main():
    init_db()
    httpd = ThreadingHTTPServer((BIND_HOST, BIND_PORT), Handler)
    print(f'RTWeb listening on http://{BIND_HOST}:{BIND_PORT}, db={DB_PATH}')
    httpd.serve_forever()

if __name__ == '__main__':
    main()
