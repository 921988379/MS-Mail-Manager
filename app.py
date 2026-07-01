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
import socket
import sqlite3
import time
import threading
import queue
import urllib.parse
import urllib.request
from datetime import datetime
from email import message_from_bytes
from email.header import decode_header, make_header
from http import HTTPStatus
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from cryptography.fernet import Fernet, InvalidToken

BASE_DIR = Path(__file__).resolve().parent
VERSION_FILE = BASE_DIR / 'VERSION'
APP_VERSION = os.environ.get('RTWEB_VERSION', VERSION_FILE.read_text().strip() if VERSION_FILE.exists() else '1.0.0')
RELEASE_REPO_URL = os.environ.get('RTWEB_RELEASE_REPO', 'https://github.com/921988379/MS-Mail-Manager')
UPDATE_REPO = os.environ.get('RTWEB_UPDATE_REPO', RELEASE_REPO_URL + '.git')
UPDATE_BRANCH = os.environ.get('RTWEB_UPDATE_BRANCH', 'main')
AUTO_UPDATE_ENABLED = os.environ.get('RTWEB_AUTO_UPDATE_ENABLED', '1') != '0'
UPDATE_COMMAND = os.environ.get('RTWEB_UPDATE_COMMAND', './scripts/update.sh')
DB_PATH = Path(os.environ.get('RTWEB_DB', BASE_DIR / 'app.db'))
ADMIN_PASSWORD = os.environ.get('RTWEB_ADMIN_PASSWORD') or os.environ.get('RTWEB_PASSWORD', 'change-me')
LOGIN_USERNAME = os.environ.get('RTWEB_LOGIN_USERNAME', 'rtweb')
LOGIN_PASSWORD = os.environ.get('RTWEB_LOGIN_PASSWORD', 'change-me')
SESSION_SECRET = os.environ.get('RTWEB_SESSION_SECRET', secrets.token_hex(32))
BIND_HOST = os.environ.get('RTWEB_HOST', '127.0.0.1')
BIND_PORT = int(os.environ.get('RTWEB_PORT', '8020'))
TOKEN_URL_TEMPLATE = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'
ME_URL = 'https://graph.microsoft.com/v1.0/me'
OUTLOOK_REST_BASE_URL = 'https://outlook.office.com/api/v2.0/me'
MAIL_FOLDERS = [('INBOX', '收件箱'), ('Junk', '垃圾邮箱')]
GRAPH_MAIL_FOLDERS = [('inbox', '收件箱'), ('junkemail', '垃圾邮箱')]
OUTLOOK_REST_MAIL_FOLDERS = [('inbox', '收件箱'), ('junkemail', '垃圾邮箱')]
API_RATE_LIMIT_PER_MINUTE = int(os.environ.get('RTWEB_API_RATE_LIMIT_PER_MINUTE', '60'))
API_RATE_LIMIT_PER_DAY = int(os.environ.get('RTWEB_API_RATE_LIMIT_PER_DAY', '2000'))
API_ALLOW_QUERY_KEY = os.environ.get('RTWEB_API_ALLOW_QUERY_KEY', '1') != '0'
DEFAULT_API_SCOPES = os.environ.get('RTWEB_DEFAULT_API_SCOPES', 'health,latest_code,accounts,projects')
DATA_KEY = os.environ.get('RTWEB_DATA_KEY', '').strip()
ENCRYPTED_PREFIX = 'enc:v1:'
LOGIN_FAIL_WINDOW_SECONDS = int(os.environ.get('RTWEB_LOGIN_FAIL_WINDOW_SECONDS', '600'))
LOGIN_FAIL_MAX = int(os.environ.get('RTWEB_LOGIN_FAIL_MAX', '8'))
TOKEN_READ_TIMEOUT_SECONDS = int(os.environ.get('RTWEB_TOKEN_READ_TIMEOUT_SECONDS', '8'))
GRAPH_READ_TIMEOUT_SECONDS = int(os.environ.get('RTWEB_GRAPH_READ_TIMEOUT_SECONDS', '6'))
MAIL_READ_TIMEOUT_SECONDS = int(os.environ.get('RTWEB_MAIL_READ_TIMEOUT_SECONDS', '5'))
API_CODES_TIMEOUT_SECONDS = int(os.environ.get('RTWEB_API_CODES_TIMEOUT_SECONDS', '18'))
LOGIN_FAILURES = {}
MAIL_FETCH_INFLIGHT = set()
MAIL_FETCH_INFLIGHT_LOCK = threading.Lock()
BATCH_JOBS = {}
BATCH_JOBS_LOCK = threading.Lock()
BATCH_JOB_TTL_SECONDS = int(os.environ.get('RTWEB_BATCH_JOB_TTL_SECONDS', '1800'))
BATCH_DONE_TTL_SECONDS = int(os.environ.get('RTWEB_BATCH_DONE_TTL_SECONDS', '600'))
BATCH_MAX_RUNNING = int(os.environ.get('RTWEB_BATCH_MAX_RUNNING', '2'))

def login_rate_limited(identity: str):
    now = int(time.time())
    hits = [t for t in LOGIN_FAILURES.get(identity, []) if now - t < LOGIN_FAIL_WINDOW_SECONDS]
    LOGIN_FAILURES[identity] = hits
    return len(hits) >= LOGIN_FAIL_MAX, len(hits)


def record_login_failure(identity: str):
    now = int(time.time())
    hits = [t for t in LOGIN_FAILURES.get(identity, []) if now - t < LOGIN_FAIL_WINDOW_SECONDS]
    hits.append(now)
    LOGIN_FAILURES[identity] = hits


def clear_login_failures(identity: str):
    LOGIN_FAILURES.pop(identity, None)

def call_with_timeout(func, timeout_seconds: int, *args, **kwargs):
    q = queue.Queue(maxsize=1)
    def runner():
        try:
            q.put(('ok', func(*args, **kwargs)))
        except Exception as exc:
            q.put(('error', exc))
    t = threading.Thread(target=runner, daemon=True)
    t.start()
    try:
        kind, value = q.get(timeout=max(1, int(timeout_seconds)))
    except queue.Empty:
        return False, TimeoutError(f'读取超时：超过 {timeout_seconds} 秒仍未返回，可能是 Microsoft Graph/IMAP 通道阻塞或账号风控。')
    if kind == 'error':
        return False, value
    return True, value

def timed_fetch_latest_codes_for_account(account, limit: int = 10):
    key = str(account.get('id') or account.get('email') or '')
    with MAIL_FETCH_INFLIGHT_LOCK:
        if key in MAIL_FETCH_INFLIGHT:
            return False, RuntimeError('上一次邮件读取仍在执行，请稍后再试；这通常说明 Microsoft Graph/IMAP 通道响应很慢。')
        MAIL_FETCH_INFLIGHT.add(key)
    q = queue.Queue(maxsize=1)
    def runner():
        try:
            q.put(('ok', fetch_latest_codes_for_account(account, limit=limit)))
        except Exception as exc:
            q.put(('error', exc))
        finally:
            with MAIL_FETCH_INFLIGHT_LOCK:
                MAIL_FETCH_INFLIGHT.discard(key)
    threading.Thread(target=runner, daemon=True).start()
    try:
        kind, value = q.get(timeout=max(1, int(API_CODES_TIMEOUT_SECONDS)))
    except queue.Empty:
        return False, TimeoutError(f'读取超时：超过 {API_CODES_TIMEOUT_SECONDS} 秒仍未返回，可能是 Microsoft Graph/IMAP 通道阻塞或账号风控。')
    if kind == 'error':
        return False, value
    return True, value


def data_fernet():
    if not DATA_KEY:
        return None
    key = base64.urlsafe_b64encode(hashlib.sha256(DATA_KEY.encode('utf-8')).digest())
    return Fernet(key)


def is_encrypted_value(value: str) -> bool:
    return isinstance(value, str) and value.startswith(ENCRYPTED_PREFIX)


def encrypt_secret_value(value: str) -> str:
    value = value or ''
    if not value or is_encrypted_value(value):
        return value
    f = data_fernet()
    if not f:
        return value
    return ENCRYPTED_PREFIX + f.encrypt(value.encode('utf-8')).decode('ascii')


def decrypt_secret_value(value: str) -> str:
    value = value or ''
    if not is_encrypted_value(value):
        return value
    f = data_fernet()
    if not f:
        return ''
    try:
        return f.decrypt(value[len(ENCRYPTED_PREFIX):].encode('ascii')).decode('utf-8')
    except (InvalidToken, ValueError, UnicodeDecodeError):
        return ''


def decrypt_account_row(row):
    if not row:
        return row
    data = dict(row)
    for field in ('password', 'aux_password', 'refresh_token'):
        data[field] = decrypt_secret_value(data.get(field) or '')
    return data


def encrypt_account_fields(password: str, aux_password: str, refresh_token: str):
    return encrypt_secret_value(password), encrypt_secret_value(aux_password), encrypt_secret_value(refresh_token)


def default_oauth_scope() -> str:
    # User only needs Microsoft Graph mail reading, not Outlook IMAP/SMTP scopes.
    return 'offline_access https://graph.microsoft.com/Mail.Read'


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


def make_csrf_cookie() -> str:
    token = secrets.token_urlsafe(24)
    sig = hmac.new(SESSION_SECRET.encode(), token.encode(), hashlib.sha256).hexdigest()
    secure = '; Secure' if os.environ.get('RTWEB_COOKIE_SECURE', '1') != '0' else ''
    # Deliberately not HttpOnly: a tiny same-origin script injects it into legacy forms.
    return f'rtweb_csrf={token}.{sig}{secure}; SameSite=Lax; Path=/; Max-Age=604800'


def verify_csrf_cookie_value(raw: str) -> bool:
    if not raw or '.' not in raw:
        return False
    token, sig = raw.split('.', 1)
    if not token or not sig:
        return False
    expected = hmac.new(SESSION_SECRET.encode(), token.encode(), hashlib.sha256).hexdigest()
    return hmac.compare_digest(sig, expected)


def verify_csrf_form(cookie_header: str, form_raw) -> bool:
    raw = parse_cookies(cookie_header).get('rtweb_csrf', '')
    if not verify_csrf_cookie_value(raw):
        return False
    submitted = (form_raw or {}).get('csrf_token', [''])[0]
    return hmac.compare_digest(submitted or '', raw)


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
                aux_email TEXT,
                aux_password TEXT,
                category TEXT,
                client_id TEXT NOT NULL,
                refresh_token TEXT NOT NULL,
                token_digest TEXT NOT NULL,
                token_mask TEXT NOT NULL,
                last_status TEXT,
                last_error TEXT
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS api_keys (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                name TEXT NOT NULL,
                key_digest TEXT NOT NULL UNIQUE,
                key_prefix TEXT NOT NULL,
                enabled INTEGER NOT NULL DEFAULT 1,
                last_used_at INTEGER,
                note TEXT,
                scopes TEXT NOT NULL DEFAULT 'health,latest_code,accounts,projects'
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS api_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                api_key_id INTEGER,
                endpoint TEXT NOT NULL,
                status TEXT NOT NULL,
                detail TEXT
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS project_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                category TEXT NOT NULL UNIQUE,
                sender_keywords TEXT,
                subject_keywords TEXT,
                body_keywords TEXT,
                max_results INTEGER NOT NULL DEFAULT 5,
                enabled INTEGER NOT NULL DEFAULT 1
            )
        ''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_api_logs_key_time ON api_logs(api_key_id, created_at)')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_api_logs_time ON api_logs(created_at)')
        conn.execute('DELETE FROM api_logs WHERE created_at < ?', (int(time.time()) - 90 * 86400,))
        api_key_cols = [r[1] for r in conn.execute('PRAGMA table_info(api_keys)')]
        if 'scopes' not in api_key_cols:
            conn.execute("ALTER TABLE api_keys ADD COLUMN scopes TEXT NOT NULL DEFAULT 'health,latest_code,accounts,projects'")
        cols = [r[1] for r in conn.execute('PRAGMA table_info(saved_accounts)')]
        if 'password' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN password TEXT')
        if 'password_mask' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN password_mask TEXT')
        if 'aux_email' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN aux_email TEXT')
        if 'aux_password' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN aux_password TEXT')
        if 'category' not in cols:
            conn.execute('ALTER TABLE saved_accounts ADD COLUMN category TEXT')
        conn.commit()
    try:
        os.chmod(DB_PATH, 0o600)
    except Exception:
        pass


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
        aux_email = ''
        aux_password = ''
        category = ''
        if len(parts) >= 4:
            email_addr, password, client_id, refresh_token = parts[0], parts[1], parts[2], parts[3]
            if len(parts) >= 6:
                aux_email, aux_password = parts[4], parts[5]
            if len(parts) >= 7:
                category = parts[6]
        else:
            email_addr, password, client_id, refresh_token = parts[0], '', parts[1], parts[2]
        if '@' not in email_addr or not client_id or not refresh_token:
            continue
        row = {'email': email_addr, 'password': password, 'client_id': client_id, 'refresh_token': refresh_token}
        if aux_email or aux_password:
            row.update({'aux_email': aux_email, 'aux_password': aux_password})
        if category:
            row['category'] = category
        rows.append(row)
    return rows


def save_account(email_addr: str, client_id: str, refresh_token: str, password: str = '', aux_email: str = '', aux_password: str = '', category: str = ''):
    category = normalize_category_name(category)
    if category:
        ensure_category(category)
    password_enc, aux_password_enc, refresh_token_enc = encrypt_account_fields(password, aux_password, refresh_token)
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO saved_accounts
            (created_at, updated_at, email, password, password_mask, aux_email, aux_password, category, client_id, refresh_token, token_digest, token_mask, last_status, last_error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL)
            ON CONFLICT(email) DO UPDATE SET
                updated_at=excluded.updated_at,
                password=CASE WHEN excluded.password != '' THEN excluded.password ELSE saved_accounts.password END,
                password_mask=CASE WHEN excluded.password_mask != '' THEN excluded.password_mask ELSE saved_accounts.password_mask END,
                aux_email=CASE WHEN excluded.aux_email != '' THEN excluded.aux_email ELSE saved_accounts.aux_email END,
                aux_password=CASE WHEN excluded.aux_password != '' THEN excluded.aux_password ELSE saved_accounts.aux_password END,
                category=CASE WHEN excluded.category != '' THEN excluded.category ELSE saved_accounts.category END,
                client_id=excluded.client_id,
                refresh_token=excluded.refresh_token,
                token_digest=excluded.token_digest,
                token_mask=excluded.token_mask
        ''', (now, now, email_addr, password_enc, mask_secret(password) if password else '', aux_email, aux_password_enc, category, client_id, refresh_token_enc, secret_digest(refresh_token), mask_secret(refresh_token)))
        conn.commit()


def update_saved_status(email_addr: str, status: str, error: str = ''):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('UPDATE saved_accounts SET last_status=?, last_error=?, updated_at=? WHERE email=?', (status, error[:1000], int(time.time()), email_addr))
        conn.commit()


OK_ACCOUNT_STATUSES = {'ok', 'token_ok', 'graph_ok', 'outlook_rest_ok', 'xoauth2_imap_ok', 'imap_password_ok'}
ERROR_ACCOUNT_STATUSES = {'error', 'token_failed', 'graph_failed', 'xoauth2_imap_failed', 'imap_password_failed', 'all_failed'}


def account_status_is_ok(value: str) -> bool:
    return (value or '') in OK_ACCOUNT_STATUSES


def account_status_is_error(value: str) -> bool:
    value = value or ''
    return value in ERROR_ACCOUNT_STATUSES or value.endswith('_failed')


def status_label(value: str) -> str:
    labels = {
        'ok': '可用',
        'error': '异常',
        'token_ok': '令牌可用',
        'token_failed': '令牌失效',
        'graph_ok': 'Graph 可读',
        'graph_failed': 'Graph 失败',
        'outlook_rest_ok': 'Outlook REST 可读',
        'outlook_rest_failed': 'Outlook REST 失败',
        'xoauth2_imap_ok': 'OAuth IMAP 可读',
        'xoauth2_imap_failed': 'OAuth IMAP 失败',
        'imap_password_ok': '密码 IMAP 可读',
        'imap_password_failed': '密码 IMAP 失败',
        'all_failed': '全部失败',
    }
    return labels.get(value or '', value or '')


ERROR_FILTERS = {
    'abuse': '账号风控 / AADSTS70000',
    'invalid_grant': '令牌失效 / invalid_grant',
    'consent': '授权或 scope 问题',
    'imap_disabled': 'IMAP 未开启或不可用',
    'login_failed': '密码或登录失败',
    'network': '网络或服务异常',
    'no_token': '未配置令牌',
}


def error_advice(error_text: str, status: str = ''):
    text = (error_text or '').strip()
    low = text.lower()
    if not text:
        if status in ('', None):
            return ('未检测', '还没有检测结果。可先点“综合检测”，确认令牌、Graph、IMAP 和密码通道是否可用。', 'unknown')
        return ('暂无错误', '当前没有保存最近错误。如状态异常但无错误详情，建议重新执行“综合检测”刷新诊断结果。', 'unknown')
    if 'aadsts70000' in low or 'service abuse mode' in low or 'abuse' in low:
        return ('账号风控 / 服务滥用限制', '暂停自动刷新和批量检测；网页登录微软账号完成验证或解锁；解锁后重新授权获取新的 Refresh Token。若无法解锁，建议标记为不可用或更换账号。', 'abuse')
    if 'invalid_grant' in low or 'refresh token' in low and ('expired' in low or 'revoked' in low or 'invalid' in low):
        return ('Refresh Token 失效', '重新登录授权获取新的 Refresh Token；如果账号同时触发风控，先网页登录解锁。不要反复用旧 token 重试。', 'invalid_grant')
    if 'aadsts65001' in low or 'consent' in low or 'permission' in low or 'scope' in low:
        return ('授权或 scope 不足', '检查 Client ID、授权 scope 和租户；重新授权并确认已授予 Mail.Read / offline_access 等必要权限。', 'consent')
    if 'imap is disabled' in low or 'imap disabled' in low or 'imap 已禁用' in text or '未连接' in text or 'not connected' in low or 'no mailbox' in low:
        return ('IMAP 未开启或邮箱未初始化', '网页登录 Outlook 初始化邮箱；确认账号允许 IMAP；如果 OAuth IMAP 不通，可尝试 Graph 或密码 IMAP 兜底。', 'imap_disabled')
    if 'authentication failed' in low or 'login failed' in low or 'invalid credentials' in low or 'password' in low and ('wrong' in low or 'incorrect' in low):
        return ('账号密码登录失败', '检查邮箱密码是否正确；确认没有要求安全验证；如果开启两步验证，需要使用应用密码或改用 OAuth 令牌。', 'login_failed')
    if 'timeout' in low or 'timed out' in low or 'connection reset' in low or 'temporarily unavailable' in low or 'http 5' in low:
        return ('网络或微软服务异常', '稍后重试；降低批量并发和频率；检查服务器出口 IP、代理和 Microsoft 服务是否临时异常。', 'network')
    if '未保存邮箱密码' in text or '未配置' in text or 'no token' in low:
        return ('缺少可用凭证', '补充 Refresh Token 或邮箱密码；如果只需要密码 IMAP，请确认密码已保存且 IMAP 可用。', 'no_token')
    return ('未知错误', '先执行“综合检测”获取更完整诊断；保留 trace_id / correlation_id 方便排查；如果是单个账号反复失败，建议网页登录确认账号状态。', 'unknown')


def error_type_matches(error_text: str, status: str, error_type: str):
    if not error_type:
        return True
    title, advice, code = error_advice(error_text, status)
    return code == error_type




def normalize_category_name(name: str) -> str:
    return (name or '').strip()


def ensure_category(name: str):
    name = normalize_category_name(name)
    if not name:
        return
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO categories (name, created_at, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET updated_at=excluded.updated_at
        ''', (name, now, now))
        conn.commit()


def sync_categories_from_accounts():
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("SELECT DISTINCT category FROM saved_accounts WHERE TRIM(COALESCE(category, '')) != ''").fetchall()
    for r in rows:
        ensure_category(r[0])


def saved_categories():
    try:
        init_db()
        sync_categories_from_accounts()
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute('''
                SELECT name FROM categories
                UNION
                SELECT DISTINCT category FROM saved_accounts WHERE TRIM(COALESCE(category, '')) != ''
                ORDER BY 1
            ''').fetchall()
            return [r[0] for r in rows]
    except sqlite3.OperationalError:
        return []


def update_account_category(account_id: str, category: str):
    category = normalize_category_name(category)
    if category:
        ensure_category(category)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('UPDATE saved_accounts SET category=?, updated_at=? WHERE id=?', (category, int(time.time()), account_id))
        conn.commit()


def update_accounts_category(account_ids, category: str):
    ids = [str(x).strip() for x in (account_ids or []) if str(x).strip().isdigit()]
    category = normalize_category_name(category)
    if not ids:
        return 0
    if category:
        ensure_category(category)
    placeholders = ','.join('?' for _ in ids)
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            f'UPDATE saved_accounts SET category=?, updated_at=? WHERE id IN ({placeholders})',
            [category, int(time.time()), *ids]
        )
        conn.commit()
        return cur.rowcount


def categories_with_counts():
    sync_categories_from_accounts()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = list(conn.execute('''
            SELECT c.id, c.name, c.created_at, c.updated_at, COUNT(a.id) AS account_count
            FROM categories c
            LEFT JOIN saved_accounts a ON a.category = c.name
            GROUP BY c.id, c.name, c.created_at, c.updated_at
            ORDER BY c.updated_at DESC, c.name ASC
        '''))
        uncategorized = conn.execute("SELECT COUNT(*) AS n FROM saved_accounts WHERE COALESCE(NULLIF(TRIM(category), ''), '未分类')='未分类'").fetchone()['n']
        return rows, uncategorized


def create_category(name: str):
    ensure_category(name)


def rename_category(category_id: str, new_name: str):
    new_name = normalize_category_name(new_name)
    if not category_id or not new_name:
        return False
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        old = conn.execute('SELECT name FROM categories WHERE id=?', (category_id,)).fetchone()
        if not old:
            return False
        old_name = old['name']
        conn.execute('''
            INSERT INTO categories (name, created_at, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET updated_at=excluded.updated_at
        ''', (new_name, now, now))
        conn.execute('UPDATE saved_accounts SET category=?, updated_at=? WHERE category=?', (new_name, now, old_name))
        conn.execute('DELETE FROM categories WHERE id=?', (category_id,))
        conn.commit()
        return True


def delete_category(category_id: str):
    if not category_id:
        return False
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute('SELECT name FROM categories WHERE id=?', (category_id,)).fetchone()
        if not row:
            return False
        conn.execute("UPDATE saved_accounts SET category='', updated_at=? WHERE category=?", (now, row['name']))
        conn.execute('DELETE FROM categories WHERE id=?', (category_id,))
        conn.commit()
        return True



def split_keywords(value: str):
    parts = re.split(r'[\n,，;；|]+', value or '')
    return [p.strip().lower() for p in parts if p.strip()]


def mail_matches_rule(mail: dict, rule) -> bool:
    if not rule:
        return True
    checks = [
        (split_keywords(rule['sender_keywords']), mail.get('from', '')),
        (split_keywords(rule['subject_keywords']), mail.get('subject', '')),
        (split_keywords(rule['body_keywords']), (mail.get('preview', '') + ' ' + mail.get('subject', ''))),
    ]
    for keywords, text in checks:
        if keywords and not any(k in (text or '').lower() for k in keywords):
            return False
    return True


def get_project_rule(category: str):
    category = normalize_category_name(category)
    if not category:
        return None
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return conn.execute('SELECT * FROM project_rules WHERE category=? AND enabled=1', (category,)).fetchone()


def upsert_project_rule(category: str, sender_keywords: str = '', subject_keywords: str = '', body_keywords: str = '', max_results: int = 5, enabled: bool = True):
    category = normalize_category_name(category)
    if not category:
        return False
    ensure_category(category)
    try:
        max_results = max(1, min(50, int(max_results)))
    except Exception:
        max_results = 5
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO project_rules (created_at, updated_at, category, sender_keywords, subject_keywords, body_keywords, max_results, enabled)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(category) DO UPDATE SET
                updated_at=excluded.updated_at,
                sender_keywords=excluded.sender_keywords,
                subject_keywords=excluded.subject_keywords,
                body_keywords=excluded.body_keywords,
                max_results=excluded.max_results,
                enabled=excluded.enabled
        ''', (now, now, category, sender_keywords.strip()[:1000], subject_keywords.strip()[:1000], body_keywords.strip()[:1000], max_results, 1 if enabled else 0))
        conn.commit()
    return True


def project_rules_map():
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute('SELECT * FROM project_rules').fetchall()
        return {r['category']: r for r in rows}


def fetch_filtered_codes_for_account(account, rule=None, limit: int = 10):
    result = fetch_latest_codes_for_account(account, limit=max(limit, 10))
    if result.get('status') != 'ok':
        return result
    codes = result.get('codes') or []
    if rule:
        filtered = []
        for c in codes:
            mail = {'from': c.get('source', ''), 'subject': c.get('subject', ''), 'preview': c.get('preview', '')}
            if mail_matches_rule(mail, rule):
                filtered.append(c)
        codes = filtered
        result['rule_applied'] = True
    result['codes'] = codes[:limit]
    return result

def migrate_sensitive_fields_to_encrypted():
    init_db()
    if not data_fernet():
        return 0
    changed = 0
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute('SELECT id,password,aux_password,refresh_token FROM saved_accounts').fetchall()
        for r in rows:
            password = r['password'] or ''
            aux_password = r['aux_password'] or ''
            refresh_token = r['refresh_token'] or ''
            password_enc = encrypt_secret_value(password)
            aux_password_enc = encrypt_secret_value(aux_password)
            refresh_token_enc = encrypt_secret_value(refresh_token)
            if (password_enc, aux_password_enc, refresh_token_enc) != (password, aux_password, refresh_token):
                conn.execute('UPDATE saved_accounts SET password=?, aux_password=?, refresh_token=? WHERE id=?', (password_enc, aux_password_enc, refresh_token_enc, r['id']))
                changed += 1
        conn.commit()
    return changed


def make_api_key() -> str:
    return 'rt_' + secrets.token_urlsafe(32)


def api_key_digest(value: str) -> str:
    return hashlib.sha256((value or '').encode('utf-8')).hexdigest()


def normalize_api_scopes(raw: str) -> str:
    allowed = {'health', 'latest_code', 'accounts', 'projects'}
    parts = [p.strip().lower().replace('-', '_') for p in re.split(r'[,;|\s]+', raw or '') if p.strip()]
    scopes = [p for p in parts if p in allowed]
    if not scopes:
        scopes = [p.strip() for p in DEFAULT_API_SCOPES.split(',') if p.strip() in allowed]
    if not scopes:
        scopes = ['health', 'latest_code']
    return ','.join(dict.fromkeys(scopes))


def api_key_has_scope(key_row, scope: str) -> bool:
    scopes = set((key_row['scopes'] or DEFAULT_API_SCOPES or '').replace(';', ',').replace('|', ',').split(','))
    scopes = {s.strip() for s in scopes if s.strip()}
    return scope in scopes


def create_api_key(name: str, note: str = '', scopes: str = ''):
    init_db()
    key = make_api_key()
    now = int(time.time())
    scope_text = normalize_api_scopes(scopes or DEFAULT_API_SCOPES)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            INSERT INTO api_keys (created_at, updated_at, name, key_digest, key_prefix, enabled, note, scopes)
            VALUES (?, ?, ?, ?, ?, 1, ?, ?)
        ''', (now, now, (name or '默认 API Key').strip()[:80], api_key_digest(key), key[:10], (note or '').strip()[:500], scope_text))
        conn.commit()
    return key


def list_api_keys():
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return list(conn.execute('SELECT * FROM api_keys ORDER BY updated_at DESC, id DESC'))


def set_api_key_enabled(key_id: str, enabled: bool):
    if not str(key_id).isdigit():
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute('UPDATE api_keys SET enabled=?, updated_at=? WHERE id=?', (1 if enabled else 0, int(time.time()), key_id))
        conn.commit()
        return cur.rowcount > 0


def delete_api_key(key_id: str):
    if not str(key_id).isdigit():
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute('DELETE FROM api_keys WHERE id=?', (key_id,))
        conn.commit()
        return cur.rowcount > 0


def verify_api_key_value(value: str):
    value = (value or '').strip()
    if not value:
        return None
    digest = api_key_digest(value)
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute('SELECT * FROM api_keys WHERE key_digest=? AND enabled=1', (digest,)).fetchone()
        if row:
            conn.execute('UPDATE api_keys SET last_used_at=?, updated_at=? WHERE id=?', (int(time.time()), int(time.time()), row['id']))
            conn.commit()
        return row


def log_api_call(api_key_id, endpoint: str, status: str, detail: str = ''):
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute('INSERT INTO api_logs (created_at, api_key_id, endpoint, status, detail) VALUES (?, ?, ?, ?, ?)', (int(time.time()), api_key_id, endpoint[:120], status[:40], detail[:500]))
            conn.commit()
    except Exception:
        pass


def recent_api_logs(limit=30):
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return list(conn.execute('''
            SELECT l.*, k.name AS key_name, k.key_prefix
            FROM api_logs l LEFT JOIN api_keys k ON k.id=l.api_key_id
            ORDER BY l.id DESC LIMIT ?
        ''', (limit,)))


def api_usage_stats():
    init_db()
    now = int(time.time())
    minute_ago = now - 60
    day_start = now - 86400
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute('''
            SELECT k.id,
                   SUM(CASE WHEN l.created_at >= ? THEN 1 ELSE 0 END) AS last_minute,
                   SUM(CASE WHEN l.created_at >= ? THEN 1 ELSE 0 END) AS today,
                   SUM(CASE WHEN l.created_at >= ? AND l.status IN ('ok','success') THEN 1 ELSE 0 END) AS today_ok,
                   SUM(CASE WHEN l.created_at >= ? AND l.status NOT IN ('ok','success') THEN 1 ELSE 0 END) AS today_fail
            FROM api_keys k
            LEFT JOIN api_logs l ON l.api_key_id = k.id
            GROUP BY k.id
        ''', (minute_ago, day_start, day_start, day_start)).fetchall()
    return {r['id']: r for r in rows}


def api_rate_limit_status(api_key_id):
    now = int(time.time())
    minute_ago = now - 60
    day_start = now - 86400
    with sqlite3.connect(DB_PATH) as conn:
        minute_count = conn.execute('SELECT COUNT(*) FROM api_logs WHERE api_key_id=? AND created_at>=?', (api_key_id, minute_ago)).fetchone()[0]
        day_count = conn.execute('SELECT COUNT(*) FROM api_logs WHERE api_key_id=? AND created_at>=?', (api_key_id, day_start)).fetchone()[0]
    if minute_count >= API_RATE_LIMIT_PER_MINUTE:
        return False, 'rate_limited_minute', minute_count, day_count
    if day_count >= API_RATE_LIMIT_PER_DAY:
        return False, 'rate_limited_day', minute_count, day_count
    return True, '', minute_count, day_count




def sh_quote(value: str) -> str:
    return "'" + str(value).replace("'", "'\\''") + "'"


def run_shell_command(command: str, timeout: int = 20):
    import subprocess
    try:
        proc = subprocess.run(command, shell=True, cwd=str(BASE_DIR), text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)
        return proc.returncode, proc.stdout[-4000:]
    except subprocess.TimeoutExpired as exc:
        return 124, (exc.stdout or '') + '\n命令超时'
    except Exception as exc:
        return 1, str(exc)


def git_output(args, timeout: int = 8):
    import subprocess
    try:
        proc = subprocess.run(['git'] + args, cwd=str(BASE_DIR), text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
        return proc.stdout.strip() if proc.returncode == 0 else ''
    except Exception:
        return ''


def version_info(fetch_remote: bool = False):
    if fetch_remote:
        run_shell_command('git fetch origin ' + sh_quote(UPDATE_BRANCH) + ' --tags', timeout=30)
    latest_commit = git_output(['rev-parse', '--short', 'origin/' + UPDATE_BRANCH]) or ''
    behind = git_output(['rev-list', '--count', 'HEAD..origin/' + UPDATE_BRANCH]) if latest_commit else ''
    latest_tag = git_output(['describe', '--tags', '--abbrev=0', 'origin/' + UPDATE_BRANCH]) or git_output(['tag', '--list', 'v*', '--sort=-v:refname']) .splitlines()[0:1]
    if isinstance(latest_tag, list):
        latest_tag = latest_tag[0] if latest_tag else ''
    current_tag = git_output(['describe', '--tags', '--exact-match', 'HEAD']) or ''
    tag_commit = git_output(['rev-list', '-n', '1', latest_tag]) if latest_tag else ''
    release_behind = ''
    if tag_commit:
        head_commit = git_output(['rev-parse', 'HEAD']) or ''
        release_behind = '0' if tag_commit == head_commit else (git_output(['rev-list', '--count', 'HEAD..' + latest_tag]) or '')
    return {
        'version': APP_VERSION,
        'branch': git_output(['branch', '--show-current']) or UPDATE_BRANCH,
        'commit': git_output(['rev-parse', '--short', 'HEAD']) or 'unknown',
        'remote': git_output(['remote', 'get-url', 'origin']) or UPDATE_REPO,
        'release_repo': RELEASE_REPO_URL,
        'update_branch': UPDATE_BRANCH,
        'latest_commit': latest_commit,
        'behind': behind,
        'latest_tag': latest_tag,
        'current_tag': current_tag,
        'release_behind': release_behind,
        'auto_update_enabled': AUTO_UPDATE_ENABLED,
    }

def dashboard_summary():
    init_db()
    now = int(time.time())
    day_start = now - 86400
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        total_accounts = conn.execute('SELECT COUNT(*) AS n FROM saved_accounts').fetchone()['n']
        ok_accounts = conn.execute("SELECT COUNT(*) AS n FROM saved_accounts WHERE last_status IN ('ok','token_ok','graph_ok','outlook_rest_ok','xoauth2_imap_ok','imap_password_ok')").fetchone()['n']
        error_accounts = conn.execute("SELECT COUNT(*) AS n FROM saved_accounts WHERE last_status IN ('error','token_failed','graph_failed','xoauth2_imap_failed','imap_password_failed','all_failed') OR last_status LIKE '%_failed'").fetchone()['n']
        categories = conn.execute("SELECT COUNT(*) AS n FROM categories").fetchone()['n']
        api_keys = conn.execute("SELECT COUNT(*) AS n FROM api_keys WHERE enabled=1").fetchone()['n']
        api_calls_today = conn.execute('SELECT COUNT(*) AS n FROM api_logs WHERE created_at>=?', (day_start,)).fetchone()['n']
        api_fail_today = conn.execute("SELECT COUNT(*) AS n FROM api_logs WHERE created_at>=? AND status NOT IN ('ok','success')", (day_start,)).fetchone()['n']
        recent = list(conn.execute('SELECT email, category, last_status, last_error, updated_at FROM saved_accounts ORDER BY updated_at DESC LIMIT 8'))
    return {'total_accounts': total_accounts, 'ok_accounts': ok_accounts, 'error_accounts': error_accounts, 'categories': categories, 'api_keys': api_keys, 'api_calls_today': api_calls_today, 'api_fail_today': api_fail_today, 'recent': recent}


def public_account_rows(category: str = '', limit: int = 50):
    limit = max(1, min(200, int(limit or 50)))
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        if category:
            if category == '未分类':
                return list(conn.execute("SELECT id,email,category,last_status,last_error,updated_at FROM saved_accounts WHERE COALESCE(NULLIF(TRIM(category), ''), '未分类')='未分类' ORDER BY updated_at DESC LIMIT ?", (limit,)))
            return list(conn.execute('SELECT id,email,category,last_status,last_error,updated_at FROM saved_accounts WHERE category=? ORDER BY updated_at DESC LIMIT ?', (category, limit)))
        return list(conn.execute('SELECT id,email,category,last_status,last_error,updated_at FROM saved_accounts ORDER BY updated_at DESC LIMIT ?', (limit,)))


def public_project_rows():
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return list(conn.execute('''
            SELECT c.name AS category,
                   COUNT(a.id) AS account_count,
                   r.sender_keywords,
                   r.subject_keywords,
                   r.body_keywords,
                   COALESCE(r.max_results, 5) AS max_results,
                   COALESCE(r.enabled, 1) AS enabled
            FROM categories c
            LEFT JOIN saved_accounts a ON a.category = c.name
            LEFT JOIN project_rules r ON r.category = c.name
            GROUP BY c.name
            ORDER BY c.name COLLATE NOCASE
        '''))

def render_version_page(message: str = ''):
    info = version_info(fetch_remote=False)
    repo_url = html.escape(info.get('release_repo') or RELEASE_REPO_URL)
    latest_tag = html.escape(info.get('latest_tag') or '未检测')
    release_raw = str(info.get('release_behind') or '')
    if release_raw == '0':
        release_text = '已是最新版'
        release_cls = 'ok'
    elif release_raw:
        release_text = '发现新版本'
        release_cls = 'bad'
    else:
        release_text = '点击检测更新'
        release_cls = 'muted'
    update_badge = '可用' if info.get('auto_update_enabled') else '已关闭'
    update_cls = 'ok' if info.get('auto_update_enabled') else 'bad'
    msg_html = ''
    if message:
        msg_html = '<div class="card"><h3>提示</h3><p>' + html.escape(message).replace('\n', '<br>') + '</p></div>'
    update_button = (
        '<button type="submit" class="primary">立即更新</button>'
        if info.get('auto_update_enabled')
        else '<button type="submit" disabled title="已设置 RTWEB_AUTO_UPDATE_ENABLED=0，手动更新关闭">立即更新</button>'
    )
    content = f"""
<section class="section-stack">
  <div class="toolbox-hero">
    <div class="toolbox-kicker">⬆️ Update</div>
    <h1 class="toolbox-title">版本更新</h1>
    <p class="toolbox-desc">检查项目是否有新版本，并可在这里手动执行更新。</p>
    <div class="toolbox-stats">
      <span class="stat-pill">当前版本：<b>{html.escape(info.get('version') or APP_VERSION)}</b></span>
      <span class="stat-pill">最新版本：<b>{latest_tag}</b></span>
      <span class="stat-pill">状态：<b class="{release_cls}">{release_text}</b></span>
      <span class="stat-pill">手动更新：<b class="{update_cls}">{update_badge}</b></span>
    </div>
  </div>

  {msg_html}

  <div class="quick-actions">
    <div class="card">
      <h3>更新操作</h3>
      <form method="post" action="/check-update" class="action-row" style="margin-bottom:10px">
        <button type="submit" class="primary">检测更新</button>
        <a class="mini-btn" href="{repo_url}" target="_blank" rel="noopener noreferrer">查看仓库</a>
      </form>
      <form method="post" action="/update" class="action-row">
        {update_button}
      </form>
      <p class="muted" style="margin-top:12px">手动更新默认可用；如需关闭，可设置 <code>RTWEB_AUTO_UPDATE_ENABLED=0</code>。更新前建议先备份数据库和环境配置。</p>
    </div>

    <div class="card">
      <h3>仓库地址</h3>
      <p><a href="{repo_url}" target="_blank" rel="noopener noreferrer">{repo_url}</a></p>
      <p class="muted">页面只展示必要版本信息；分支、提交号、远程地址等调试信息已隐藏。</p>
    </div>
  </div>
</section>"""
    return app_page('版本/更新', 'version', content)


def get_saved_account_by_email(email_addr: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return decrypt_account_row(conn.execute('SELECT * FROM saved_accounts WHERE lower(email)=lower(?)', ((email_addr or '').strip(),)).fetchone())


def newest_account_for_category(category: str):
    category = normalize_category_name(category)
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        if category == '未分类':
            return decrypt_account_row(conn.execute("SELECT * FROM saved_accounts WHERE COALESCE(NULLIF(TRIM(category), ''), '未分类')='未分类' ORDER BY updated_at DESC LIMIT 1").fetchone())
        return decrypt_account_row(conn.execute('SELECT * FROM saved_accounts WHERE category=? ORDER BY updated_at DESC LIMIT 1', (category,)).fetchone())

def export_accounts_text():
    lines = []
    for a in saved_accounts():
        fields = [
            a['email'] or '',
            a['password'] or '',
            a['client_id'] or '',
            a['refresh_token'] or '',
            a['aux_email'] or '',
            a['aux_password'] or '',
        ]
        lines.append('----'.join(str(x) for x in fields))
    return '\n'.join(lines) + ('\n' if lines else '')

def saved_accounts(category: str = ''):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        if category:
            if category == '未分类':
                return [decrypt_account_row(r) for r in conn.execute("SELECT * FROM saved_accounts WHERE COALESCE(NULLIF(TRIM(category), ''), '未分类')='未分类' ORDER BY updated_at DESC")]
            return [decrypt_account_row(r) for r in conn.execute("SELECT * FROM saved_accounts WHERE category=? ORDER BY updated_at DESC", (category,))]
        return [decrypt_account_row(r) for r in conn.execute('SELECT * FROM saved_accounts ORDER BY updated_at DESC')]


def paged_saved_accounts(category: str = '', page: int = 1, per_page: int = 50, status_filter: str = '', error_type: str = '', q: str = ''):
    init_db()
    page = max(1, int(page or 1))
    per_page = int(per_page or 50)
    if per_page not in (20, 50, 100, 200):
        per_page = 50
    offset = (page - 1) * per_page
    where_parts = []
    args = []
    if category:
        if category == '未分类':
            where_parts.append("COALESCE(NULLIF(TRIM(category), ''), '未分类')='未分类'")
        else:
            where_parts.append('category=?')
            args.append(category)
    if status_filter == 'ok':
        where_parts.append("last_status IN ('ok','token_ok','graph_ok','xoauth2_imap_ok','imap_password_ok')")
    elif status_filter == 'error':
        where_parts.append("(last_status IN ('error','token_failed','graph_failed','xoauth2_imap_failed','imap_password_failed','all_failed') OR last_status LIKE '%_failed')")
    elif status_filter == 'unchecked':
        where_parts.append("(last_status IS NULL OR TRIM(last_status)='')")
    elif status_filter:
        where_parts.append('last_status=?')
        args.append(status_filter)
    if q:
        where_parts.append('(email LIKE ? OR category LIKE ? OR last_error LIKE ?)')
        like = '%' + q + '%'
        args.extend([like, like, like])
    if error_type:
        # Error types are rule-based and easier to keep correct in Python; prefilter to error-ish rows first.
        where_parts.append("(last_error IS NOT NULL AND TRIM(last_error)!='')")
    where_sql = (' WHERE ' + ' AND '.join(where_parts)) if where_parts else ''
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        if error_type:
            all_rows = list(conn.execute('SELECT * FROM saved_accounts' + where_sql + ' ORDER BY updated_at DESC', args))
            matched = [r for r in all_rows if error_type_matches(row_get(r, 'last_error', ''), row_get(r, 'last_status', ''), error_type)]
            total = len(matched)
            rows = matched[offset:offset + per_page]
        else:
            total = conn.execute('SELECT COUNT(*) AS n FROM saved_accounts' + where_sql, args).fetchone()['n']
            rows = list(conn.execute('SELECT * FROM saved_accounts' + where_sql + ' ORDER BY updated_at DESC LIMIT ? OFFSET ?', args + [per_page, offset]))
    total_pages = max(1, (int(total) + per_page - 1) // per_page)
    if page > total_pages:
        return paged_saved_accounts(category, total_pages, per_page, status_filter, error_type, q)
    return {'rows': [decrypt_account_row(r) for r in rows], 'total': int(total), 'page': page, 'per_page': per_page, 'total_pages': total_pages}


def mailbox_page_url(category: str = '', page: int = 1, per_page: int = 50, status_filter: str = '', error_type: str = '', q: str = ''):
    params = {'page': str(max(1, int(page or 1))), 'per_page': str(int(per_page or 50))}
    if category:
        params['category'] = category
    if status_filter:
        params['status'] = status_filter
    if error_type:
        params['error_type'] = error_type
    if q:
        params['q'] = q
    return '/mailboxes?' + urllib.parse.urlencode(params)


def render_pagination(meta: dict, category: str = '', status_filter: str = '', error_type: str = '', q: str = ''):
    total = int(meta.get('total') or 0)
    page = int(meta.get('page') or 1)
    per_page = int(meta.get('per_page') or 50)
    total_pages = int(meta.get('total_pages') or 1)
    start = 0 if total == 0 else ((page - 1) * per_page + 1)
    end = min(total, page * per_page)
    per_page_options = ''.join('<option value="' + str(n) + '"' + (' selected' if n == per_page else '') + '>每页 ' + str(n) + ' 条</option>' for n in (20, 50, 100, 200))
    page_links = []
    last = 0
    for n in sorted(set([1, total_pages, page - 2, page - 1, page, page + 1, page + 2])):
        if n < 1 or n > total_pages:
            continue
        if last and n - last > 1:
            page_links.append('<span class="muted">…</span>')
        cls = 'mini-btn primary' if n == page else 'mini-btn'
        page_links.append('<a class="' + cls + '" href="' + html.escape(mailbox_page_url(category, n, per_page, status_filter, error_type, q)) + '">' + str(n) + '</a>')
        last = n
    prev_link = '<span class="mini-btn disabled">上一页</span>' if page <= 1 else '<a class="mini-btn" href="' + html.escape(mailbox_page_url(category, page - 1, per_page, status_filter, error_type, q)) + '">上一页</a>'
    next_link = '<span class="mini-btn disabled">下一页</span>' if page >= total_pages else '<a class="mini-btn" href="' + html.escape(mailbox_page_url(category, page + 1, per_page, status_filter, error_type, q)) + '">下一页</a>'
    hidden_category = '<input type="hidden" name="category" value="' + html.escape(category) + '">' if category else ''
    hidden_status = '<input type="hidden" name="status" value="' + html.escape(status_filter) + '">' if status_filter else ''
    hidden_error_type = '<input type="hidden" name="error_type" value="' + html.escape(error_type) + '">' if error_type else ''
    hidden_q = '<input type="hidden" name="q" value="' + html.escape(q) + '">' if q else ''
    return '''
<div class="pagination-bar">
  <span class="muted">显示 {start}-{end} / 共 {total} 个邮箱，第 {page}/{total_pages} 页</span>
  <span class="bulk-spacer"></span>
  <form method="get" action="/mailboxes" class="action-row" style="gap:6px;margin:0">
    {hidden_category}{hidden_status}{hidden_error_type}{hidden_q}
    <select name="per_page" onchange="this.form.submit()">{per_page_options}</select>
    <input name="page" value="{page}" style="width:70px" aria-label="页码">
    <button type="submit" class="mini-btn">跳转</button>
  </form>
  {prev_link}
  {page_links}
  {next_link}
</div>'''.format(start=start, end=end, total=total, page=page, total_pages=total_pages, hidden_category=hidden_category, hidden_status=hidden_status, hidden_error_type=hidden_error_type, hidden_q=hidden_q, per_page_options=per_page_options, prev_link=prev_link, page_links=''.join(page_links), next_link=next_link)


def get_saved_account(account_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return decrypt_account_row(conn.execute('SELECT * FROM saved_accounts WHERE id=?', (account_id,)).fetchone())


def selected_saved_accounts(account_ids):
    ids = [str(x).strip() for x in (account_ids or []) if str(x).strip().isdigit()]
    if not ids:
        return []
    placeholders = ','.join('?' for _ in ids)
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = list(conn.execute(f'SELECT * FROM saved_accounts WHERE id IN ({placeholders}) ORDER BY updated_at DESC', ids))
    return [decrypt_account_row(r) for r in rows]


def update_account_refresh_token(account_id: str, refresh_token: str):
    refresh_token_enc = encrypt_secret_value(refresh_token)
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            UPDATE saved_accounts
            SET refresh_token=?, token_digest=?, token_mask=?, updated_at=?, last_status='ok', last_error=NULL
            WHERE id=?
        ''', (refresh_token_enc, secret_digest(refresh_token), mask_secret(refresh_token), now, account_id))
        conn.commit()


def refresh_saved_account_token(account):
    scope = default_oauth_scope()
    ok, payload = exchange_refresh_token_compatible(account['client_id'], account['refresh_token'], scope=scope, tenant='consumers')
    if ok:
        new_refresh = payload.get('refresh_token') or account['refresh_token']
        if new_refresh != account['refresh_token']:
            update_account_refresh_token(str(account['id']), new_refresh)
        insert_record(account['client_id'], new_refresh, 'ok', account=account['email'], scope=payload.get('scope'), expires_in=payload.get('expires_in'))
        update_saved_status(account['email'], 'token_ok')
        return True, payload, new_refresh
    err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
    insert_record(account['client_id'], account['refresh_token'], 'error', account=account['email'], error=err)
    update_saved_status(account['email'], 'token_failed', err)
    return False, payload, account['refresh_token']


def check_saved_account_token(account):
    ok, payload, refresh_token = refresh_saved_account_token(account)
    if ok:
        return {
            'email': account['email'],
            'status': '可用',
            'ok': True,
            'expires_in': payload.get('expires_in'),
            'scope': payload.get('scope', ''),
            'scope_mode': payload.get('_scope_mode', ''),
            'rotated': refresh_token != account['refresh_token'],
            'error': ''
        }
    err = (payload.get('error_description') or payload.get('error') or '令牌不可用')[:500]
    return {
        'email': account['email'],
        'status': '失效/错误',
        'ok': False,
        'expires_in': '',
        'scope': '',
        'rotated': False,
        'error': err
    }


def delete_saved_account(account_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('DELETE FROM saved_accounts WHERE id=?', (account_id,))
        conn.commit()


def delete_saved_accounts(account_ids):
    ids = [str(x).strip() for x in (account_ids or []) if str(x).strip().isdigit()]
    if not ids:
        return 0
    placeholders = ','.join('?' for _ in ids)
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(f'DELETE FROM saved_accounts WHERE id IN ({placeholders})', ids)
        conn.commit()
        return int(cur.rowcount or 0)


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
        with urllib.request.urlopen(req, timeout=TOKEN_READ_TIMEOUT_SECONDS) as resp:
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


def should_retry_without_scope(payload: dict) -> bool:
    desc = (payload.get('error_description') or '') if isinstance(payload, dict) else ''
    codes = payload.get('error_codes') or [] if isinstance(payload, dict) else []
    return payload.get('error') == 'invalid_grant' and (70000 in codes or 'AADSTS70000' in desc or 'scope' in desc.lower())


def exchange_refresh_token_compatible(client_id: str, refresh_token: str, scope: str = '', tenant: str = 'consumers'):
    ok, payload = exchange_refresh_token(client_id, refresh_token, scope=scope, tenant=tenant)
    if ok:
        payload.setdefault('_scope_mode', 'requested' if scope else 'original')
        return ok, payload
    if scope and should_retry_without_scope(payload):
        ok2, payload2 = exchange_refresh_token(client_id, refresh_token, scope='', tenant=tenant)
        if ok2:
            payload2['_scope_mode'] = 'original_scope_fallback'
            payload2['_fallback_from_error'] = payload.get('error_description') or payload.get('error') or ''
            return ok2, payload2
    return ok, payload


def fetch_me(access_token: str):
    req = urllib.request.Request(ME_URL, headers={'Authorization': f'Bearer {access_token}'})
    try:
        with urllib.request.urlopen(req, timeout=GRAPH_READ_TIMEOUT_SECONDS) as resp:
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


def outlook_rest_message_to_mail(msg: dict):
    sender = (((msg.get('From') or msg.get('from') or {}).get('EmailAddress') or (msg.get('from') or {}).get('emailAddress')) or {})
    name = sender.get('Name') or sender.get('name') or ''
    addr = sender.get('Address') or sender.get('address') or ''
    source = (name + (' <' + addr + '>' if addr else '')).strip() or addr
    body = msg.get('BodyPreview') or msg.get('bodyPreview') or ''
    if not body:
        body_obj = msg.get('Body') or msg.get('body') or {}
        body = body_obj.get('Content') or body_obj.get('content') or ''
        if (body_obj.get('ContentType') or body_obj.get('contentType') or '').lower() == 'html':
            body = html_to_text_preview(body)
    return {
        'from': source,
        'subject': msg.get('Subject') or msg.get('subject') or '',
        'date': msg.get('ReceivedDateTime') or msg.get('receivedDateTime') or msg.get('SentDateTime') or msg.get('sentDateTime') or '',
        'preview': body or '',
    }


def imap4_ssl_ipv4(host: str, port: int = 993, timeout: int = None):
    # Some hosts resolve Microsoft IMAP to IPv6 addresses even when the server has
    # no usable IPv6 route, which surfaces as [Errno 101] Network is unreachable.
    # Force AF_INET for IMAP while keeping the original hostname for TLS/SNI.
    original_getaddrinfo = socket.getaddrinfo
    def getaddrinfo_ipv4(name, svc, family=0, type=0, proto=0, flags=0):
        rows = original_getaddrinfo(name, svc, socket.AF_INET, type, proto, flags)
        if rows:
            return rows
        return original_getaddrinfo(name, svc, family, type, proto, flags)
    socket.getaddrinfo = getaddrinfo_ipv4
    try:
        return imaplib.IMAP4_SSL(host, port, timeout=timeout)
    finally:
        socket.getaddrinfo = original_getaddrinfo

def fetch_graph_latest_emails(access_token: str, limit: int = 10):
    results = []
    for folder_id, folder_label in GRAPH_MAIL_FOLDERS:
        url = f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?' + urllib.parse.urlencode({
            '$top': str(limit),
            '$orderby': 'receivedDateTime desc',
            '$select': 'receivedDateTime,sentDateTime,subject,bodyPreview,from'
        })
        req = urllib.request.Request(url, headers={'Authorization': f'Bearer {access_token}', 'Accept': 'application/json'})
        with urllib.request.urlopen(req, timeout=GRAPH_READ_TIMEOUT_SECONDS) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
        for m in payload.get('value', []):
            mail = graph_message_to_mail(m)
            mail['folder'] = folder_label
            results.append(mail)
    return results[:limit * len(GRAPH_MAIL_FOLDERS)]


def fetch_outlook_rest_latest_emails(access_token: str, limit: int = 10):
    results = []
    for folder_id, folder_label in OUTLOOK_REST_MAIL_FOLDERS:
        url = f'{OUTLOOK_REST_BASE_URL}/mailfolders/{folder_id}/messages?' + urllib.parse.urlencode({
            '$top': str(limit),
            '$orderby': 'ReceivedDateTime desc',
            '$select': 'ReceivedDateTime,SentDateTime,Subject,BodyPreview,From'
        })
        req = urllib.request.Request(url, headers={'Authorization': f'Bearer {access_token}', 'Accept': 'application/json'})
        with urllib.request.urlopen(req, timeout=GRAPH_READ_TIMEOUT_SECONDS) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
        for m in payload.get('value', []):
            mail = outlook_rest_message_to_mail(m)
            mail['folder'] = folder_label
            results.append(mail)
    return results[:limit * len(OUTLOOK_REST_MAIL_FOLDERS)]


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
        'preview': (mail.get('preview', '') or '')[:500],
        'folder': mail.get('folder', ''),
    }


def summarize_latest_mails(mails, limit: int = 5):
    rows = []
    for m in (mails or [])[:max(1, int(limit or 5))]:
        rows.append({
            'source': m.get('from', ''),
            'subject': m.get('subject', ''),
            'date': m.get('date', ''),
            'folder': m.get('folder', ''),
            'preview': (m.get('preview', '') or '')[:500],
        })
    return rows


def fetch_latest_codes_for_account(account, limit: int = 10):
    graph_scope = default_oauth_scope()
    gok, gpayload = exchange_refresh_token_compatible(account['client_id'], account['refresh_token'], scope=graph_scope, tenant='consumers')
    graph_error = ''
    if gok:
        try:
            mails = fetch_graph_latest_emails(gpayload.get('access_token'), limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='graph') for c in codes if c]
            if codes or mails:
                update_saved_status(account['email'], 'graph_ok')
                return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'graph', 'channel_status': {'graph': 'ok', 'token': 'ok'}, 'mail_count': len(mails), 'latest_mails': summarize_latest_mails(mails), 'codes': codes}
        except Exception as e:
            graph_error = str(e)
    else:
        graph_error = gpayload.get('error_description') or gpayload.get('error') or ''

    if gok:
        try:
            mails = fetch_outlook_rest_latest_emails(gpayload.get('access_token'), limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='outlook_rest') for c in codes if c]
            if codes or mails:
                update_saved_status(account['email'], 'outlook_rest_ok')
                return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'outlook_rest', 'graph_error': graph_error, 'channel_status': {'graph': 'failed' if graph_error else 'empty', 'token': 'ok', 'outlook_rest': 'ok'}, 'mail_count': len(mails), 'latest_mails': summarize_latest_mails(mails), 'codes': codes}
        except Exception as e:
            graph_error = (graph_error + ' | Outlook REST: ' + str(e)).strip(' |')

    if gok:
        try:
            mails = fetch_latest_emails(account['email'], gpayload.get('access_token'), limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='imap') for c in codes if c]
            update_saved_status(account['email'], 'xoauth2_imap_ok')
            return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'imap', 'graph_error': graph_error, 'channel_status': {'graph': 'failed' if graph_error else 'empty', 'token': 'ok', 'outlook_rest': 'failed' if graph_error else 'empty', 'xoauth2_imap': 'ok'}, 'mail_count': len(mails), 'latest_mails': summarize_latest_mails(mails), 'codes': codes}
        except Exception as e:
            err = str(e)
            if is_imap_not_connected_error(err):
                err = 'IMAP 已认证但邮箱未连接：通常是该账号未开通/未初始化 Outlook 邮箱，或 Microsoft 账户没有可连接的 Exchange mailbox。Graph/Outlook REST 兜底错误：' + graph_error
            graph_error = (graph_error + ' | XOAUTH2 IMAP: ' + err).strip(' |')
    else:
        graph_error = (graph_error + ' | token: ' + graph_error).strip(' |')
    if account.get('password'):
        try:
            mails = fetch_password_latest_emails(account['email'], account['password'], limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='imap_password') for c in codes if c]
            update_saved_status(account['email'], 'imap_password_ok')
            return {'email': account['email'], 'status': 'ok', 'error': '', 'source_api': 'imap_password', 'graph_error': graph_error, 'channel_status': {'graph': 'failed' if graph_error else 'empty', 'token': 'failed' if graph_error else 'unknown', 'imap_password': 'ok'}, 'mail_count': len(mails), 'latest_mails': summarize_latest_mails(mails), 'codes': codes}
        except Exception as e:
            err = '账号密码 IMAP 失败：' + str(e)
            update_saved_status(account['email'], 'all_failed', (graph_error + ' | ' + err)[:1000])
            return {'email': account['email'], 'status': 'error', 'error': (graph_error + ' | ' + err)[:1000], 'source_api': '', 'channel_status': {'graph': 'failed' if graph_error else 'unknown', 'token': 'failed', 'imap_password': 'failed'}, 'mail_count': 0, 'latest_mails': [], 'codes': []}
    err = (graph_error or '令牌不可用，且未保存邮箱密码，无法使用账号密码 IMAP 兜底。')[:1000]
    update_saved_status(account['email'], 'all_failed', err)
    return {'email': account['email'], 'status': 'error', 'error': err, 'source_api': '', 'channel_status': {'graph': 'failed' if graph_error else 'unknown', 'token': 'failed', 'imap_password': 'not_configured'}, 'mail_count': 0, 'latest_mails': [], 'codes': []}


def fetch_latest_emails(username: str, access_token: str, limit: int = 10):
    imap = imap4_ssl_ipv4('outlook.office365.com', 993, timeout=MAIL_READ_TIMEOUT_SECONDS)
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


def fetch_password_latest_emails(username: str, password: str, limit: int = 10):
    if not password:
        raise ValueError('未保存邮箱密码，无法使用账号密码 IMAP 读取。')
    imap = imap4_ssl_ipv4('imap-mail.outlook.com', 993, timeout=MAIL_READ_TIMEOUT_SECONDS)
    try:
        imap.login(username, password)
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
                results.append({'from': decode_mime(msg.get('From')), 'subject': decode_mime(msg.get('Subject')), 'date': decode_mime(msg.get('Date')), 'preview': message_preview(msg), 'folder': folder_label})
        return results
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def inspect_saved_account(account, limit: int = 10):
    result = {'email': account.get('email', ''), 'password_login': {'configured': bool(account.get('password')), 'ok': False, 'error': '', 'mail_count': 0, 'codes': []}, 'token': {'configured': bool(account.get('client_id') and account.get('refresh_token')), 'ok': False, 'status': '未配置', 'expires_in': '', 'scope': '', 'rotated': False, 'error': ''}, 'graph_mail': {'ok': False, 'error': '', 'mail_count': 0, 'codes': []}, 'outlook_rest_mail': {'ok': False, 'error': '', 'mail_count': 0, 'codes': []}, 'best_codes': [], 'best_source': ''}
    if result['token']['configured']:
        token_status = check_saved_account_token(account)
        result['token'].update({'ok': bool(token_status.get('ok')), 'status': token_status.get('status') or '', 'expires_in': token_status.get('expires_in') or '', 'scope': token_status.get('scope') or '', 'rotated': bool(token_status.get('rotated')), 'error': token_status.get('error') or ''})
        ok, payload = exchange_refresh_token_compatible(account['client_id'], account['refresh_token'], scope=default_oauth_scope(), tenant='consumers')
        if ok:
            try:
                mails = fetch_graph_latest_emails(payload.get('access_token'), limit=limit)
                codes = [extract_verification_summary(account['email'], m) for m in mails]
                codes = [dict(c, source_api='graph') for c in codes if c]
                result['graph_mail'].update({'ok': True, 'mail_count': len(mails), 'codes': codes})
                if codes:
                    result['best_codes'] = codes
                    result['best_source'] = 'graph'
            except Exception as e:
                result['graph_mail']['error'] = str(e)[:500]
            try:
                mails = fetch_outlook_rest_latest_emails(payload.get('access_token'), limit=limit)
                codes = [extract_verification_summary(account['email'], m) for m in mails]
                codes = [dict(c, source_api='outlook_rest') for c in codes if c]
                result['outlook_rest_mail'].update({'ok': True, 'mail_count': len(mails), 'codes': codes})
                if codes and not result['best_codes']:
                    result['best_codes'] = codes
                    result['best_source'] = 'outlook_rest'
            except Exception as e:
                result['outlook_rest_mail']['error'] = str(e)[:500]
        else:
            err = (payload.get('error_description') or payload.get('error') or '无法换取 Access Token')[:500]
            result['graph_mail']['error'] = err
            result['outlook_rest_mail']['error'] = err
    if account.get('password'):
        try:
            mails = fetch_password_latest_emails(account['email'], account['password'], limit=limit)
            codes = [extract_verification_summary(account['email'], m) for m in mails]
            codes = [dict(c, source_api='imap_password') for c in codes if c]
            result['password_login'].update({'ok': True, 'mail_count': len(mails), 'codes': codes})
            if codes and not result['best_codes']:
                result['best_codes'] = codes
                result['best_source'] = 'imap_password'
        except Exception as e:
            result['password_login']['error'] = str(e)[:500]
    if result['token']['ok'] or result['graph_mail']['ok'] or result['outlook_rest_mail']['ok'] or result['password_login']['ok']:
        update_saved_status(account['email'], 'graph_ok' if result['graph_mail']['ok'] else ('outlook_rest_ok' if result['outlook_rest_mail']['ok'] else ('imap_password_ok' if result['password_login']['ok'] else 'token_ok')))
    else:
        err = result['token']['error'] or result['graph_mail']['error'] or result['outlook_rest_mail']['error'] or result['password_login']['error'] or '综合检测失败'
        update_saved_status(account['email'], 'all_failed', err[:1000])
    return result


def render_batch_inspect_result(results):
    ok_count = sum(1 for r in results if r['token']['ok'] or r['graph_mail']['ok'] or r.get('outlook_rest_mail', {}).get('ok') or r['password_login']['ok'])
    graph_ok = sum(1 for r in results if r['graph_mail']['ok'])
    outlook_ok = sum(1 for r in results if r.get('outlook_rest_mail', {}).get('ok'))
    token_ok = sum(1 for r in results if r['token']['ok'])
    pwd_ok = sum(1 for r in results if r['password_login']['ok'])
    code_count = sum(len(r.get('best_codes') or []) for r in results)
    rows = ''.join(
        '<tr><td>' + html.escape(r.get('email','')) + '</td>'
        '<td>' + ('<span class="ok">可用</span>' if r['token']['ok'] else '<span class="bad">失败</span>' if r['token']['configured'] else '<span class="muted">未配置</span>') + '</td>'
        '<td>' + ('<span class="ok">可读</span>' if r['graph_mail']['ok'] else '<span class="bad">失败</span>' if r['token']['configured'] else '<span class="muted">未配置</span>') + '</td>'
        '<td>' + ('<span class="ok">可读</span>' if r.get('outlook_rest_mail', {}).get('ok') else '<span class="bad">失败</span>' if r['token']['configured'] else '<span class="muted">未配置</span>') + '</td>'
        '<td>' + ('<span class="ok">可读</span>' if r['password_login']['ok'] else '<span class="bad">失败</span>' if r['password_login']['configured'] else '<span class="muted">未配置</span>') + '</td>'
        '<td>' + html.escape(r.get('best_source') or '') + '</td>'
        '<td>' + html.escape(str(len(r.get('best_codes') or []))) + '</td>'
        '<td>' + html.escape((r['token'].get('error') or r['graph_mail'].get('error') or r.get('outlook_rest_mail', {}).get('error') or r['password_login'].get('error') or '')[:160]) + '</td></tr>'
        for r in results
    ) or '<tr><td colspan="8" class="muted">没有选择账号。</td></tr>'
    body = '<div class="card"><h2>批量综合检测结果</h2>'
    body += '<div class="stat-grid"><div class="stat"><b>' + str(len(results)) + '</b><span>检测账号</span></div><div class="stat"><b>' + str(ok_count) + '</b><span>至少一个通道可用</span></div><div class="stat"><b>' + str(token_ok) + '</b><span>令牌可用</span></div><div class="stat"><b>' + str(graph_ok) + '</b><span>Graph 可读</span></div><div class="stat"><b>' + str(outlook_ok) + '</b><span>Outlook REST 可读</span></div><div class="stat"><b>' + str(pwd_ok) + '</b><span>密码 IMAP 可读</span></div><div class="stat"><b>' + str(code_count) + '</b><span>验证码摘要</span></div></div>'
    body += '<p class="muted">综合检测会按 Refresh Token / Graph / Outlook REST / OAuth IMAP / 账号密码 IMAP 聚合判断；结果不会显示密码、Access Token 或 Refresh Token 明文。</p>'
    body += '<table><tr><th>邮箱</th><th>令牌</th><th>Graph</th><th>Outlook REST</th><th>密码 IMAP</th><th>最佳来源</th><th>验证码数</th><th>错误摘要</th></tr>' + rows + '</table></div>'
    return page('批量综合检测结果', body)


def render_account_inspect_result(result: dict):
    def badge(ok, configured=True):
        if not configured:
            return '<span class="muted">未配置</span>'
        return '<span class="ok">成功</span>' if ok else '<span class="bad">失败</span>'
    code_rows = ''.join('<tr><td>' + html.escape(c.get('source_api','')) + '</td><td>' + html.escape(c.get('folder','')) + '</td><td><b>' + html.escape(c.get('code','')) + '</b></td><td>' + html.escape(c.get('source','')) + '</td><td>' + html.escape(c.get('subject','')) + '</td><td>' + html.escape(c.get('date','')) + '</td></tr>' for c in (result.get('best_codes') or [])[:10]) or '<tr><td colspan="6" class="muted">暂无验证码摘要。</td></tr>'
    token = result['token']; graph = result['graph_mail']; outlook = result.get('outlook_rest_mail', {'ok': False, 'error': '', 'mail_count': 0}); pwd = result['password_login']
    html_body = '<div class="card"><h2>账号综合检测：' + html.escape(result.get('email','')) + '</h2>'
    html_body += '<table><tr><th>项目</th><th>状态</th><th>信息</th></tr>'
    html_body += '<tr><td>Refresh Token / Access Token</td><td>' + badge(token.get('ok'), token.get('configured')) + '</td><td>状态：' + html.escape(str(token.get('status') or '')) + ' · 有效期：' + html.escape(str(token.get('expires_in') or '')) + ' · 轮换：' + ('是' if token.get('rotated') else '否') + '<br><span class="bad">' + html.escape(token.get('error') or '') + '</span></td></tr>'
    html_body += '<tr><td>Graph 邮件读取</td><td>' + badge(graph.get('ok'), token.get('configured')) + '</td><td>邮件数：' + html.escape(str(graph.get('mail_count') or 0)) + '<br><span class="bad">' + html.escape(graph.get('error') or '') + '</span></td></tr>'
    html_body += '<tr><td>Outlook REST 邮件读取</td><td>' + badge(outlook.get('ok'), token.get('configured')) + '</td><td>邮件数：' + html.escape(str(outlook.get('mail_count') or 0)) + '<br><span class="bad">' + html.escape(outlook.get('error') or '') + '</span></td></tr>'
    html_body += '<tr><td>账号密码 IMAP</td><td>' + badge(pwd.get('ok'), pwd.get('configured')) + '</td><td>邮件数：' + html.escape(str(pwd.get('mail_count') or 0)) + '<br><span class="bad">' + html.escape(pwd.get('error') or '') + '</span></td></tr>'
    html_body += '</table><h3>最新验证码 / 邮件摘要</h3><p class="muted">优先展示 Graph 结果；Graph 不可用时尝试 Outlook REST，再兜底账号密码 IMAP。不会显示密码、Access Token 或 Refresh Token 明文。</p><table><tr><th>来源</th><th>文件夹</th><th>验证码</th><th>发件人</th><th>主题</th><th>时间</th></tr>' + code_rows + '</table></div>'
    return page('账号综合检测', html_body)




def cleanup_batch_jobs():
    now = time.time()
    with BATCH_JOBS_LOCK:
        stale = []
        for job_id, job in BATCH_JOBS.items():
            age = now - job.get('created_at', now)
            idle = now - job.get('updated_at', job.get('created_at', now))
            if job.get('status') in ('done', 'error') and idle > BATCH_DONE_TTL_SECONDS:
                stale.append(job_id)
            elif age > BATCH_JOB_TTL_SECONDS:
                stale.append(job_id)
        for job_id in stale:
            BATCH_JOBS.pop(job_id, None)


def running_batch_job_count():
    cleanup_batch_jobs()
    with BATCH_JOBS_LOCK:
        return sum(1 for job in BATCH_JOBS.values() if job.get('status') == 'running')


def create_batch_job(title: str, kind: str, rows, worker):
    cleanup_batch_jobs()
    if running_batch_job_count() >= BATCH_MAX_RUNNING:
        return None
    job_id = secrets.token_urlsafe(18)
    emails = [row_get(a, 'email', '') for a in rows]
    account_ids = [str(row_get(a, 'id', '')) for a in rows]
    job = {
        'id': job_id,
        'title': title,
        'kind': kind,
        'total': len(rows),
        'done': 0,
        'ok': 0,
        'failed': 0,
        'current': '',
        'current_index': 0,
        'status': 'running',
        'error': '',
        'emails': emails,
        'account_ids': account_ids,
        'status_updates': [],
        'results': [],
        'html': '',
        'created_at': time.time(),
        'updated_at': time.time(),
    }
    with BATCH_JOBS_LOCK:
        BATCH_JOBS[job_id] = job

    def runner():
        try:
            results = []
            if not rows:
                with BATCH_JOBS_LOCK:
                    job.update({'status': 'done', 'html': '<div class="card"><h2>' + html.escape(title) + '</h2><p class="muted">没有可处理的邮箱。</p></div>', 'updated_at': time.time()})
                return
            for idx, account in enumerate(rows, 1):
                email_addr = row_get(account, 'email', '')
                with BATCH_JOBS_LOCK:
                    job.update({'current': email_addr, 'current_index': idx, 'updated_at': time.time()})
                item = worker(account)
                results.append(item)
                item_ok = False
                status_value = ''
                status_error = ''
                if kind == 'inspect':
                    item_ok = bool(item.get('token', {}).get('ok') or item.get('graph_mail', {}).get('ok') or item.get('password_login', {}).get('ok'))
                    if item.get('graph_mail', {}).get('ok'):
                        status_value = 'graph_ok'
                    elif item.get('password_login', {}).get('ok'):
                        status_value = 'imap_password_ok'
                    elif item.get('token', {}).get('ok'):
                        status_value = 'token_ok'
                    else:
                        status_value = 'all_failed'
                        status_error = item.get('token', {}).get('error') or item.get('graph_mail', {}).get('error') or item.get('password_login', {}).get('error') or ''
                elif kind in ('check',):
                    item_ok = bool(item.get('ok'))
                    status_value = 'token_ok' if item_ok else 'token_failed'
                    status_error = item.get('error') or ''
                elif isinstance(item, (tuple, list)) and len(item) >= 2:
                    item_ok = bool(item[1])
                    status_value = 'token_ok' if item_ok else 'token_failed'
                    payload = item[2] if len(item) >= 3 and isinstance(item[2], dict) else {}
                    status_error = payload.get('error_description') or payload.get('error') or ''
                status_update = {
                    'id': str(row_get(account, 'id', '')),
                    'email': email_addr,
                    'status': status_value,
                    'label': status_label(status_value),
                    'ok': account_status_is_ok(status_value),
                    'error': status_error[:500],
                }
                with BATCH_JOBS_LOCK:
                    job['results'] = results
                    job['status_updates'].append(status_update)
                    job['done'] = idx
                    job['ok'] += 1 if item_ok else 0
                    job['failed'] += 0 if item_ok else 1
                    job['updated_at'] = time.time()
            if kind == 'token':
                result_html = Handler.render_batch_token_result_static(results)
            elif kind == 'refresh':
                result_html = Handler.render_batch_refresh_result_static(results, title)
            elif kind == 'check':
                result_html = Handler.render_batch_check_result_static(results)
            elif kind == 'inspect':
                result_html = render_batch_inspect_result(results)
            else:
                result_html = page(title, '<div class="card"><h2>' + html.escape(title) + '</h2><p class="bad">未知任务类型。</p></div>')
            with BATCH_JOBS_LOCK:
                job.update({'status': 'done', 'current': '', 'html': extract_card_html(result_html), 'updated_at': time.time()})
        except Exception as exc:
            with BATCH_JOBS_LOCK:
                job.update({'status': 'error', 'error': str(exc)[:1000], 'current': '', 'updated_at': time.time()})

    threading.Thread(target=runner, daemon=True).start()
    return job_id


def get_batch_job(job_id: str):
    cleanup_batch_jobs()
    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        if not job:
            return None
        return {k: v for k, v in job.items() if k != 'results'}


def extract_card_html(page_html: str) -> str:
    m = re.search(r'<div class="card">(.*)</div>', page_html, re.S)
    return m.group(1) if m else page_html

UI_POLISH_CSS = "\nbody{background:radial-gradient(circle at 18% -8%,#1d4ed833 0,transparent 28rem),radial-gradient(circle at 86% 4%,#7c3aed22 0,transparent 24rem),#070b14;color:#e2e8f0;font-size:14px}.wrap{max-width:1500px;padding:14px}.top{background:#0b1220cc;border:1px solid #263244;border-radius:18px;padding:12px 14px;box-shadow:0 12px 30px #0004;backdrop-filter:blur(10px)}.top h2{font-size:17px;letter-spacing:-.02em}.top a{background:#111827;border:1px solid #334155;border-radius:999px;padding:7px 11px;text-decoration:none;color:#cbd5e1}.top a:hover{border-color:#60a5fa;color:white}.app-layout{grid-template-columns:238px 1fr;gap:14px}.sidebar{background:linear-gradient(180deg,#0b1220f5,#070b14f5);border-color:#263244cc;border-radius:20px;padding:14px}.side-title{font-size:16px;margin:4px 8px 14px}.side-link{background:transparent;border-radius:14px;padding:12px;color:#b6c2d3}.side-link:hover,.side-link.active{background:linear-gradient(135deg,#1d4ed833,#7c3aed22);border-color:#60a5fa66;box-shadow:inset 0 0 0 1px #ffffff08}.app-page-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;margin:0 0 14px;padding:18px;border:1px solid #263244;border-radius:20px;background:linear-gradient(135deg,#111827,#0f172a 60%,#172554);box-shadow:0 14px 34px #0004;position:relative;overflow:hidden}.app-page-head:after{content:'';position:absolute;right:-50px;top:-70px;width:210px;height:210px;border-radius:999px;background:#38bdf833;filter:blur(30px)}.app-page-title{position:relative;font-size:24px;font-weight:900;letter-spacing:-.04em;color:#f8fafc}.app-page-subtitle{position:relative;color:#94a3b8;margin-top:6px;line-height:1.6}.app-page-badge{position:relative;border:1px solid #60a5fa55;background:#1d4ed833;color:#dbeafe;border-radius:999px;padding:7px 11px;font-size:12px;font-weight:800;white-space:nowrap}.card{background:linear-gradient(180deg,#101827,#0b1220);border:1px solid #263244cc;border-radius:18px;padding:16px;box-shadow:0 14px 34px #0004}.card h3{font-size:17px;color:#f8fafc;letter-spacing:-.02em;margin-bottom:10px}.card h4{margin:14px 0 8px;color:#dbeafe}.section-grid{grid-template-columns:minmax(330px,420px) 1fr;gap:14px}.section-stack{gap:14px}.muted{color:#94a3b8;line-height:1.65}.ok{color:#86efac}.bad{color:#fca5a5}.notice{border:1px solid #334155;border-radius:14px;padding:12px 13px;margin:10px 0;background:#0f172a;line-height:1.65}.notice.ok{background:linear-gradient(135deg,#052e1699,#0f172a);border-color:#22c55e66;color:#dcfce7}.notice.bad{background:linear-gradient(135deg,#450a0a99,#0f172a);border-color:#ef444466;color:#fee2e2}.notice.info{background:linear-gradient(135deg,#17255499,#0f172a);border-color:#60a5fa66;color:#dbeafe}input,textarea,select{background:#050a14;border-color:#334155;border-radius:12px;padding:10px 11px;transition:border-color .15s,box-shadow .15s,background .15s}input:hover,textarea:hover,select:hover{border-color:#475569}input:focus,textarea:focus,select:focus{outline:0;border-color:#60a5fa;box-shadow:0 0 0 4px #2563eb26;background:#07101f}button,.button-link,.mini-btn{border-radius:11px!important;font-weight:800;box-shadow:0 8px 18px #0002;transition:transform .12s,filter .12s,border-color .12s}button:hover,.button-link:hover,.mini-btn:hover{transform:translateY(-1px);filter:brightness(1.08)}.mini-btn{display:inline-flex!important;align-items:center;justify-content:center;text-decoration:none!important;color:white!important;background:#334155!important;border:1px solid #475569!important}.mini-btn.primary{background:#2563eb!important;border-color:#60a5fa!important}.mini-btn.success{background:#0f766e!important;border-color:#2dd4bf!important}.mini-btn.warning{background:#ea580c!important;border-color:#fdba74!important}.mini-btn.danger{background:#dc2626!important;border-color:#fca5a5!important}.action-row,.toolbar{gap:9px}table{border-collapse:separate;border-spacing:0;width:100%}th{position:sticky;top:0;background:#0b1220;color:#cbd5e1;font-size:12px;font-weight:900;text-transform:none;z-index:1}td,th{border-bottom:1px solid #1f2a3a;padding:8px 7px}tr:hover td{background:#11182799}tr.ok td{background:#052e1644}.scroll{overflow:auto;border-radius:14px}.chip{padding:7px 11px;border-color:#334155;background:#0b1220}.chip.active,.chip:hover{background:linear-gradient(135deg,#2563eb,#7c3aed);border-color:#93c5fd}.category-bar{margin:12px 0 16px}.stat-pill{background:#0b1220cc;border-color:#334155;padding:11px 13px}.tool-card{border-radius:18px}.modal{border-radius:20px}.bulk-bar{background:linear-gradient(135deg,#0f172a,#111827);border-color:#334155;color:#e2e8f0}.selected-preview{background:#0f172a;border-color:#334155;color:#e2e8f0}.code-inline{display:inline-block;padding:2px 7px;border-radius:999px;background:#172554;color:#bfdbfe;border:1px solid #1d4ed8;font-family:ui-monospace,monospace}.mailbox-email{font-weight:800;color:#f8fafc}.mailbox-meta{display:flex;gap:6px;flex-wrap:wrap;margin-top:5px}.status-badge{display:inline-flex;align-items:center;border-radius:999px;padding:4px 8px;font-size:12px;font-weight:900;border:1px solid #334155;background:#0f172a;color:#cbd5e1}.status-badge.ok{border-color:#22c55e66;background:#052e16;color:#bbf7d0}.status-badge.bad{border-color:#ef444466;background:#450a0a;color:#fecaca}.mailbox-detail summary{cursor:pointer;color:#93c5fd;font-weight:800}.detail-grid{display:grid;grid-template-columns:90px 1fr;gap:6px 10px;margin-top:8px;min-width:280px}.detail-label{color:#94a3b8}.detail-value{word-break:break-all}.ops-wrap{display:flex;gap:6px;flex-wrap:wrap;min-width:360px}.table-note{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:8px}.pagination-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:8px 0 10px;padding:8px;border:1px solid #263244;border-radius:12px;background:#0b1220}.pagination-bar select,.pagination-bar input{width:auto;margin:0}.pagination-bar .disabled{opacity:.45;pointer-events:none}.filter-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:10px 0 14px;padding:10px;border:1px solid #263244;border-radius:14px;background:#0b1220}.filter-bar input,.filter-bar select{width:auto;margin:0}.mail-search{position:sticky;top:0;background:linear-gradient(180deg,#101827,#101827ee);padding-bottom:10px;z-index:2}.mail-picker-list{display:grid;gap:8px}.mail-picker-item{display:grid;grid-template-columns:1fr auto;gap:6px 10px;text-decoration:none;color:#e2e8f0;border:1px solid #263244;border-radius:14px;padding:10px 11px;background:#0b1220}.mail-picker-item:hover,.mail-picker-item.active{border-color:#60a5fa;background:linear-gradient(135deg,#172554aa,#0b1220)}.mail-picker-email{font-weight:900;word-break:break-all}.mail-picker-meta{font-size:12px;color:#94a3b8}.mail-picker-action{align-self:center;border-radius:999px;background:#1d4ed8;color:#dbeafe;padding:5px 9px;font-size:12px;font-weight:900}.mail-result-top{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px}.mail-result-area{display:grid;gap:12px}.mail-summary-card{border:1px solid #334155;border-radius:18px;padding:14px;background:#0f172a}.mail-summary-card.ok{border-color:#22c55e66;background:linear-gradient(135deg,#052e1699,#0f172a)}.mail-summary-card.bad{border-color:#ef444466;background:linear-gradient(135deg,#450a0a99,#0f172a)}.metric-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.metric-chip{border:1px solid #334155;background:#02061780;border-radius:999px;padding:6px 9px;color:#cbd5e1;font-size:12px}.code-grid,.mail-card-grid{display:grid;gap:10px}.code-card,.mail-card{border:1px solid #263244;border-radius:16px;padding:12px;background:#0b1220}.code-card-head{display:flex;justify-content:space-between;align-items:center;gap:10px}.code-pill{font-family:ui-monospace,monospace;font-size:22px;font-weight:900;color:#fef3c7;background:#713f12;border:1px solid #f59e0b66;border-radius:14px;padding:7px 11px;letter-spacing:.04em}.mail-subject{font-weight:900;color:#f8fafc;margin-bottom:5px;word-break:break-word}.mail-preview{color:#cbd5e1;line-height:1.65;margin-top:7px}.empty-state{border:1px dashed #334155;border-radius:18px;padding:22px;text-align:center;background:#0b1220;color:#94a3b8}.progress-shell{display:grid;gap:12px}.progress-track{height:16px;border-radius:999px;background:#020617;border:1px solid #334155;overflow:hidden}.progress-fill{height:100%;width:0;background:linear-gradient(90deg,#2563eb,#22c55e);transition:width .25s}.progress-meta{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}.progress-meta span{border:1px solid #334155;background:#02061780;border-radius:12px;padding:9px;color:#cbd5e1}.progress-current{font-size:15px;word-break:break-all}.progress-email-list{max-height:130px;overflow:auto;border:1px solid #263244;border-radius:12px;padding:8px;background:#0b1220;color:#94a3b8}.progress-email-list div.active{color:#fef3c7;font-weight:900}@media(max-width:900px){.wrap{padding:10px}.app-page-head{display:block;padding:15px}.app-page-badge{display:inline-flex;margin-top:10px}.side-nav{grid-template-columns:1fr 1fr}.section-grid{grid-template-columns:1fr}.card{padding:13px}td,th{padding:7px 6px}}\n"

def page(title, body, show_nav=True, body_class=''):
    nav = '<div class="top"><h2>🧰 一点微软工具箱</h2><a href="/logout">退出</a></div>' if show_nav else ''
    body_class_attr = f' class="{html.escape(body_class)}"' if body_class else ''
    return f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title><style>
:root{{color-scheme:dark}}body{{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#0f172a;color:#e5e7eb;font-size:13px}}.wrap{{max-width:1440px;margin:0 auto;padding:10px}}.card{{background:#111827;border:1px solid #263244;border-radius:12px;padding:12px;margin:0;box-shadow:0 8px 18px #0003;min-height:0}}input,textarea{{width:100%;box-sizing:border-box;background:#020617;color:#e5e7eb;border:1px solid #334155;border-radius:8px;padding:8px;margin:4px 0 8px;font:inherit}}textarea{{min-height:70px;font-family:ui-monospace,monospace}}button{{background:#2563eb;color:white;border:0;border-radius:8px;padding:8px 12px;font-weight:700;cursor:pointer}}button:hover{{background:#1d4ed8}}h2,h3{{margin:0 0 8px}}.muted{{color:#94a3b8}}.ok{{color:#86efac}}.bad{{color:#fca5a5}}code,pre{{background:#020617;border:1px solid #334155;border-radius:8px;padding:8px;display:block;overflow:auto}}table{{width:100%;border-collapse:collapse}}td,th{{border-bottom:1px solid #263244;padding:5px;text-align:left;font-size:12px;vertical-align:top}}a{{color:#93c5fd}}.top{{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:8px}}.grid{{display:grid;grid-template-columns:360px 360px 1fr;grid-template-rows:auto 1fr;gap:10px;height:calc(100vh - 62px)}}.span2{{grid-column:1 / span 2}}.codes{{grid-row:1 / span 2;grid-column:3;overflow:auto}}.scroll{{overflow:auto}}
.login-body{{min-height:100vh;background:#020617;display:flex;align-items:center;justify-content:center;padding:28px;box-sizing:border-box;overflow-x:hidden}}.login-body:before{{content:'';position:fixed;inset:-20%;background:radial-gradient(circle at 18% 18%,#2563eb55 0 16rem,transparent 33rem),radial-gradient(circle at 88% 18%,#7c3aed44 0 14rem,transparent 30rem),radial-gradient(circle at 60% 90%,#0891b244 0 16rem,transparent 34rem);filter:blur(10px);pointer-events:none}}.login-body:after{{content:'';position:fixed;inset:0;background-image:linear-gradient(#ffffff08 1px,transparent 1px),linear-gradient(90deg,#ffffff08 1px,transparent 1px);background-size:46px 46px;mask-image:radial-gradient(circle at center,#000 0 38%,transparent 75%);pointer-events:none}}.login-shell{{position:relative;width:min(1040px,100%);display:grid;grid-template-columns:1.06fr .94fr;border:1px solid #334155aa;border-radius:34px;overflow:hidden;background:linear-gradient(145deg,#0f172af2,#020617f6);box-shadow:0 30px 110px #000b,0 0 0 1px #ffffff08}}.login-hero{{position:relative;padding:42px;min-height:560px;background:linear-gradient(155deg,#111827 0%,#0f172a 42%,#172554 100%);overflow:hidden}}.login-hero:before{{content:'';position:absolute;right:-90px;bottom:-90px;width:270px;height:270px;border-radius:999px;background:#38bdf855;filter:blur(40px)}}.login-hero:after{{content:'';position:absolute;left:36px;right:36px;bottom:34px;height:1px;background:linear-gradient(90deg,transparent,#60a5fa88,transparent)}}.login-panel{{padding:42px;background:linear-gradient(180deg,#0b1220f7,#020617fa);display:flex;flex-direction:column;justify-content:center}}.login-brand{{position:relative;display:inline-flex;gap:10px;align-items:center;color:#dbeafe;background:#1d4ed84d;border:1px solid #60a5fa66;border-radius:999px;padding:9px 13px;font-weight:800;box-shadow:0 10px 30px #1d4ed833}}.login-title{{position:relative;font-size:42px;line-height:1.05;margin:28px 0 14px;letter-spacing:-.055em}}.login-subtitle{{position:relative;color:#cbd5e1;font-size:15px;line-height:1.75;max-width:35rem}}.login-points{{position:relative;margin:32px 0 0;padding:0;list-style:none;display:grid;gap:13px}}.login-points li{{display:flex;gap:11px;color:#dbeafe;background:#ffffff08;border:1px solid #ffffff10;border-radius:15px;padding:11px 12px}}.login-points li:before{{content:'✓';display:grid;place-items:center;flex:0 0 20px;height:20px;border-radius:999px;background:#16a34a;color:white;font-size:12px;font-weight:900}}.login-badge-row{{position:relative;display:flex;gap:10px;flex-wrap:wrap;margin-top:28px}}.login-badge{{border:1px solid #334155;background:#02061799;color:#bfdbfe;border-radius:999px;padding:7px 10px;font-size:12px}}.login-panel h2{{font-size:28px;margin:0 0 8px;letter-spacing:-.03em}}.login-panel .muted{{font-size:14px;line-height:1.65}}.login-form{{margin-top:12px}}.field{{margin-top:16px}}.field-head{{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}}.login-panel label{{color:#cbd5e1;font-weight:800}}.field-hint{{font-size:12px;color:#64748b}}.password-row{{position:relative}}.login-panel input{{height:50px;border-radius:15px;padding:0 44px 0 14px;background:#020617cc;border:1px solid #334155;color:#e5e7eb;transition:border-color .15s,box-shadow .15s,background .15s}}.login-panel input:hover{{border-color:#475569;background:#030b1acc}}.login-panel input:focus{{outline:0;border-color:#60a5fa;box-shadow:0 0 0 4px #2563eb33}}.toggle-pass{{position:absolute;right:8px;top:50%;transform:translateY(-50%);width:34px!important;height:34px!important;margin:0!important;padding:0!important;border-radius:11px!important;background:#1e293b!important;color:#cbd5e1!important;font-size:13px!important}}.toggle-pass:hover{{background:#334155!important}}.login-submit{{width:100%;height:50px;border-radius:15px;margin-top:22px;background:linear-gradient(135deg,#2563eb,#7c3aed 55%,#0891b2);font-size:15px;box-shadow:0 14px 30px #2563eb33}}.login-submit:hover{{filter:brightness(1.08);background:linear-gradient(135deg,#1d4ed8,#6d28d9 55%,#0e7490)}}.login-note{{font-size:12px;color:#94a3b8;line-height:1.7;margin-top:16px;padding:12px 13px;border:1px solid #33415588;background:#02061780;border-radius:15px}}.login-error{{border:1px solid #ef444466;background:linear-gradient(135deg,#7f1d1dcc,#450a0acc);color:#fecaca;border-radius:15px;padding:11px 13px;margin:14px 0 0}}.security-line{{display:flex;align-items:center;gap:8px;margin-top:14px;color:#86efac;font-size:12px}}.security-dot{{width:8px;height:8px;border-radius:999px;background:#22c55e;box-shadow:0 0 16px #22c55e}}.app-layout{{display:grid;grid-template-columns:230px 1fr;gap:12px;min-height:calc(100vh - 62px)}}.sidebar{{background:#0b1220;border:1px solid #263244;border-radius:16px;padding:12px;box-shadow:0 10px 24px #0003;position:sticky;top:10px;height:calc(100vh - 82px)}}.side-title{{font-weight:900;color:#dbeafe;margin:4px 6px 12px;font-size:15px}}.side-nav{{display:grid;gap:8px}}.side-link{{display:flex;align-items:center;gap:10px;text-decoration:none;color:#cbd5e1;border:1px solid transparent;border-radius:12px;padding:11px 12px;background:#11182788}}.side-link:hover,.side-link.active{{background:#1d4ed833;border-color:#3b82f688;color:#fff}}.side-icon{{width:22px;text-align:center}}.content-panel{{display:none}}.content-panel.active{{display:block}}.section-grid{{display:grid;grid-template-columns:380px 1fr;gap:12px}}.section-stack{{display:grid;gap:12px}}.action-row{{display:flex;gap:8px;flex-wrap:wrap;align-items:center}}.category-bar{{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 14px}}.chip{{display:inline-flex;align-items:center;border:1px solid #334155;border-radius:999px;padding:6px 10px;background:#0f172a;color:#cbd5e1;text-decoration:none;font-size:13px}}.chip.active,.chip:hover{{background:#1d4ed8;color:#fff;border-color:#60a5fa}}.toolbar{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}}.button-link{{display:inline-flex;align-items:center;justify-content:center;border-radius:10px;padding:9px 13px;background:#7c3aed;color:white!important;text-decoration:none;font-weight:800;border:0;line-height:1.2}}.button-link:hover{{filter:brightness(1.08)}}.modal-backdrop{{position:fixed;inset:0;background:#020617cc;backdrop-filter:blur(8px);display:none;align-items:center;justify-content:center;z-index:50;padding:18px}}.modal-backdrop.active{{display:flex}}.modal{{width:min(620px,100%);max-height:92vh;overflow:auto;background:#111827;border:1px solid #334155;border-radius:18px;box-shadow:0 28px 90px #000b;padding:18px}}.modal-head{{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px}}.modal-head h3{{margin:0}}.modal-close{{width:auto!important;height:auto!important;background:#334155!important;padding:6px 10px!important;margin:0!important}}.bulk-bar{{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:10px 0 14px;padding:10px;border:1px solid #dbeafe;border-radius:12px;background:#eff6ff}}.bulk-bar.compact{{padding:8px 10px;background:#f8fafc;border-color:#e2e8f0}}.mini-btn{{width:auto!important;height:auto!important;padding:6px 10px!important;margin:0!important;background:#334155!important}}.mini-btn.primary{{background:#7c3aed!important}}.bulk-spacer{{flex:1 1 auto}}.bulk-category-input{{max-width:100%;margin-bottom:12px}}.account-check{{width:auto!important}}.selected-preview{{padding:10px 12px;margin:8px 0 14px;border-radius:10px;background:#f8fafc;border:1px solid #e2e8f0;color:#0f172a}}.bulk-action-panel{{border:1px solid #263244;border-radius:14px;padding:12px;margin-top:12px;background:#0f172a}}.bulk-action-panel h4{{margin:0 0 10px;color:#dbeafe}}.toolbox-hero{{position:relative;overflow:hidden;border-radius:18px;padding:24px;background:linear-gradient(135deg,#0f172a,#172554 48%,#312e81);border:1px solid #334155;box-shadow:0 18px 45px #0005}}.toolbox-hero:before{{content:'';position:absolute;right:-80px;top:-80px;width:260px;height:260px;border-radius:999px;background:#38bdf855;filter:blur(38px)}}.toolbox-kicker{{position:relative;display:inline-flex;gap:8px;align-items:center;border:1px solid #60a5fa66;background:#1d4ed84d;color:#dbeafe;border-radius:999px;padding:7px 11px;font-size:12px;font-weight:900}}.toolbox-title{{position:relative;font-size:34px;line-height:1.1;margin:16px 0 8px;letter-spacing:-.04em}}.toolbox-desc{{position:relative;color:#cbd5e1;max-width:760px;line-height:1.75;margin:0}}.toolbox-stats{{position:relative;display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}}.stat-pill{{border:1px solid #334155;background:#02061780;color:#bfdbfe;border-radius:14px;padding:10px 12px}}.tool-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}}.tool-card{{display:block;text-decoration:none;color:#e5e7eb;background:linear-gradient(180deg,#111827,#0b1220);border:1px solid #263244;border-radius:16px;padding:16px;min-height:132px;box-shadow:0 10px 24px #0003;transition:transform .16s,border-color .16s,background .16s}}.tool-card:hover{{transform:translateY(-2px);border-color:#60a5fa;background:linear-gradient(180deg,#172554,#0b1220)}}.tool-icon{{font-size:28px;margin-bottom:12px}}.tool-card h3{{font-size:17px;margin-bottom:8px}}.tool-card p{{margin:0;color:#94a3b8;line-height:1.55}}.quick-actions{{display:grid;grid-template-columns:1.15fr .85fr;gap:12px}}.quick-list{{display:grid;gap:9px}}.quick-item{{display:flex;justify-content:space-between;gap:12px;align-items:center;border:1px solid #263244;background:#0f172a;border-radius:12px;padding:10px 12px}}.quick-item a{{text-decoration:none;font-weight:800}}.danger-note{{border:1px solid #7f1d1d;background:#450a0a80;color:#fecaca;border-radius:12px;padding:12px;line-height:1.65}}@media(max-width:1100px){{.tool-grid{{grid-template-columns:repeat(2,minmax(0,1fr))}}.quick-actions{{grid-template-columns:1fr}}}}@media(max-width:900px){{.app-layout{{grid-template-columns:1fr}}.sidebar{{position:static;height:auto}}.side-nav{{grid-template-columns:1fr 1fr 1fr}}.section-grid{{grid-template-columns:1fr}}.grid{{display:block;height:auto}}.card{{margin-bottom:10px}}.login-body{{padding:16px}}.login-shell{{grid-template-columns:1fr;border-radius:24px}}.login-hero{{display:none}}.login-panel{{padding:26px}}.login-panel h2{{font-size:24px}}}}
{UI_POLISH_CSS}
</style></head><body{body_class_attr}>{'<div class="wrap">' + nav + body + '</div>' if show_nav else body}<script>
(function() {{
  function cookie(name) {{
    return document.cookie.split(';').map(function(v){{return v.trim();}}).filter(function(v){{return v.indexOf(name + '=') === 0;}}).map(function(v){{return decodeURIComponent(v.slice(name.length + 1));}})[0] || '';
  }}
  var token = cookie('rtweb_csrf');
  if (!token) return;
  document.querySelectorAll('form[method="post"],form[method="POST"]').forEach(function(form) {{
    if (form.querySelector('input[name="csrf_token"]')) return;
    var input = document.createElement('input');
    input.type = 'hidden'; input.name = 'csrf_token'; input.value = token;
    form.appendChild(input);
  }});
}})();
</script></body></html>'''


def render_login_page(error: str = ''):
    err = f'<div class="login-error">{html.escape(error)}</div>' if error else ''
    return page('登录 - 一点微软工具箱', f'''
<main class="login-shell">
  <section class="login-hero">
    <div class="login-brand">🧰 一点微软工具箱</div>
    <h1 class="login-title">一点微软工具箱</h1>
    <p class="login-subtitle">部署在你自己服务器上的微软账号工具箱，用于管理 Outlook Token、刷新 Access Token、读取 Hotmail / Outlook 验证码邮件。</p>
    <ul class="login-points">
      <li>用户账号密码验证，拦住非授权访问</li>
      <li>应用管理密码二次确认，保护敏感 token 数据</li>
      <li>全站 HTTPS，Cookie 使用 HttpOnly / Secure / SameSite</li>
    </ul>
    <div class="login-badge-row">
      <span class="login-badge">Self-hosted</span>
      <span class="login-badge">SQLite Local</span>
      <span class="login-badge">HTTPS Only</span>
    </div>
  </section>
  <section class="login-panel">
    <h2>登录一点微软工具箱</h2>
    <p class="muted">请依次完成两层验证。三项都正确后才能进入后台。</p>
    {err}
    <form class="login-form" method="post" action="/login" autocomplete="off">
      <div class="field">
        <div class="field-head"><label for="username">用户账号</label><span class="field-hint">第一层</span></div>
        <input id="username" type="text" name="username" placeholder="请输入用户账号" autofocus required>
      </div>
      <div class="field">
        <div class="field-head"><label for="login_password">用户密码</label><span class="field-hint">第一层</span></div>
        <div class="password-row"><input id="login_password" type="password" name="login_password" placeholder="请输入用户密码" required><button class="toggle-pass" type="button" data-target="login_password">显示</button></div>
      </div>
      <div class="field">
        <div class="field-head"><label for="admin_password">应用管理密码</label><span class="field-hint">第二层</span></div>
        <div class="password-row"><input id="admin_password" type="password" name="admin_password" placeholder="请输入应用管理密码" required><button class="toggle-pass" type="button" data-target="admin_password">显示</button></div>
      </div>
      <button class="login-submit" type="submit">安全进入</button>
    </form>
    <div class="security-line"><span class="security-dot"></span><span>当前通过 HTTPS 加密访问</span></div>
    <p class="login-note">建议只在可信设备使用。离开时点击退出，避免他人访问你的 token 数据。</p>
  </section>
</main>
<script>
document.querySelectorAll('.toggle-pass').forEach(btn => {{
  btn.addEventListener('click', () => {{
    const input = document.getElementById(btn.dataset.target);
    const show = input.type === 'password';
    input.type = show ? 'text' : 'password';
    btn.textContent = show ? '隐藏' : '显示';
  }});
}});
</script>''', show_nav=False, body_class='login-body')

def sidebar(active: str) -> str:
    items = [
        ('dashboard', '/', '🏠', '控制台'),
        ('tokens', '/tokens', '🔄', '刷新令牌'),
        ('mailboxes', '/mailboxes', '📮', '邮箱管理'),
        ('mails', '/mails', '📥', '获取邮件'),
        ('api_key', '/api-key', '🔐', 'API 密钥'),
        ('api', '/api-manage', '🔌', '查询 API'),
        ('projects', '/project-manage', '📁', '项目管理'),
        ('categories', '/categories', '🏷️', '分类管理'),
        ('help', '/help', '📘', '使用指南'),
        ('version', '/version', '⬆️', '版本/更新'),
    ]
    links = ''.join(
        f'<a href="{href}" class="side-link {"active" if key == active else ""}"><span class="side-icon">{icon}</span><span>{label}</span></a>'
        for key, href, icon, label in items
    )
    return f'<aside class="sidebar"><div class="side-title">一点微软工具箱</div><nav class="side-nav">{links}</nav></aside>'


def app_page(title: str, active: str, content: str):
    result_modal = '''
<div id="result-modal" class="modal-backdrop" onclick="closeResultModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3 id="result-modal-title">操作结果</h3><button type="button" class="modal-close" onclick="closeResultModal()">关闭</button></div>
    <div id="result-modal-body"><p class="muted">处理中...</p></div>
  </div>
</div>
<script>
function showResultModal(title, html) {
  const modal = document.getElementById('result-modal');
  const titleEl = document.getElementById('result-modal-title');
  const body = document.getElementById('result-modal-body');
  if (!modal || !titleEl || !body) return;
  titleEl.textContent = title || '操作结果';
  body.innerHTML = html || '<p class="muted">无返回内容</p>';
  modal.classList.add('active');
}
function closeResultModal() {
  const modal = document.getElementById('result-modal');
  if (modal) modal.classList.remove('active');
}
function closeResultModalOnBackdrop(e) {
  if (e.target && e.target.id === 'result-modal') closeResultModal();
}
function extractResultHtml(raw) {
  try {
    const doc = new DOMParser().parseFromString(raw, 'text/html');
    const card = doc.querySelector('.card');
    if (card) return card.innerHTML;
    const body = doc.body;
    return body ? body.innerHTML : raw;
  } catch (e) {
    return '<pre>' + escapeHtml(raw) + '</pre>';
  }
}
function escapeHtml(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}
async function copyToClipboard(idOrText) {
  const el = document.getElementById(idOrText);
  const text = el ? (el.value || el.textContent || '') : idOrText;
  try {
    await navigator.clipboard.writeText(text);
    showResultModal('复制成功', '<p class="ok">内容已复制到剪贴板。</p>');
  } catch (e) {
    showResultModal('手动复制', '<p class="muted">浏览器不允许自动复制，请手动复制：</p><textarea readonly style="min-height:90px">' + escapeHtml(text) + '</textarea>');
  }
}
function toggleSecret(id, btn) {
  const el = document.getElementById(id);
  if (!el) return;
  const hidden = el.style.display === 'none';
  el.style.display = hidden ? '' : 'none';
  if (btn) btn.textContent = hidden ? '隐藏' : '显示';
}
async function runModalRequest(url, options, title) {
  showProgressModal(title, {total: 0, done: 0, ok: 0, failed: 0, current: '', emails: []});
  try {
    const res = await fetch(url, Object.assign({credentials: 'same-origin', cache: 'no-store'}, options || {}));
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('application/json')) {
      const data = await res.json();
      if (data && data.job_id) { pollBatchJob(data.job_id, title); return; }
      if (data && data.error === 'too_many_batch_jobs') {
        showResultModal(title, '<p class="bad">已有批量任务正在运行，请等当前任务完成后再开始新的批量操作。</p>');
        return;
      }
      showResultModal(title, '<pre>' + escapeHtml(JSON.stringify(data, null, 2)) + '</pre>');
      return;
    }
    const text = await res.text();
    showResultModal(title, extractResultHtml(text));
  } catch (e) {
    showResultModal(title, '<p class="bad">请求失败：' + escapeHtml(e.message || String(e)) + '</p>');
  }
}
function statusBadgeHtml(item) {
  const cls = item.ok ? 'ok' : (item.status ? 'bad' : '');
  return '<span class="status-badge ' + cls + '">' + escapeHtml(item.label || item.status || '未检测') + '</span>';
}
function applyBatchStatusUpdates(data) {
  (data.status_updates || []).forEach(item => {
    if (!item.id) return;
    const cell = document.querySelector('[data-status-cell="' + CSS.escape(String(item.id)) + '"]');
    if (cell) cell.innerHTML = statusBadgeHtml(item);
  });
}
function showProgressModal(title, data) {
  applyBatchStatusUpdates(data || {});
  const total = Number(data.total || 0), done = Number(data.done || 0);
  const pct = total ? Math.round(done * 100 / total) : 0;
  const current = data.current ? escapeHtml(data.current) : (data.status === 'done' ? '已完成' : '准备开始...');
  let emails = '';
  (data.emails || []).slice(0, 200).forEach((mail, idx) => {
    const active = (idx + 1) === Number(data.current_index || 0) ? ' class="active"' : '';
    emails += '<div' + active + '>' + (idx + 1) + '. ' + escapeHtml(mail) + '</div>';
  });
  showResultModal(title, `<div class="progress-shell">
    <div class="progress-current">正在处理：<b>${current}</b></div>
    <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
    <div class="progress-meta"><span>进度 <b>${done}/${total}</b></span><span>完成 <b>${pct}%</b></span><span class="ok">成功 <b>${Number(data.ok || 0)}</b></span><span class="bad">失败 <b>${Number(data.failed || 0)}</b></span></div>
    <div class="muted">批量任务执行中，请不要关闭这个弹窗。完成后会自动显示结果。</div>
    ${emails ? '<div class="progress-email-list">' + emails + '</div>' : ''}
  </div>`);
}
async function pollBatchJob(jobId, title) {
  try {
    const res = await fetch('/api/batch_job?id=' + encodeURIComponent(jobId), {credentials:'same-origin', cache:'no-store'});
    const data = await res.json();
    if (!data || data.error) { showResultModal(title, '<p class="bad">进度查询失败：' + escapeHtml((data && data.error) || 'unknown') + '</p>'); return; }
    applyBatchStatusUpdates(data);
    if (data.status === 'done') { showResultModal(title, data.html || '<p class="ok">已完成。</p>'); return; }
    if (data.status === 'error') { showResultModal(title, '<p class="bad">任务失败：' + escapeHtml(data.error || '') + '</p>'); return; }
    showProgressModal(title, data);
    setTimeout(() => pollBatchJob(jobId, title), 1500);
  } catch (e) {
    showResultModal(title, '<p class="bad">进度查询失败：' + escapeHtml(e.message || String(e)) + '</p>');
  }
}
document.addEventListener('click', e => {
  const link = e.target.closest('a[href^="/token?"]');
  if (!link) return;
  e.preventDefault();
  runModalRequest(link.getAttribute('href'), {}, '获取令牌');
});
document.addEventListener('submit', e => {
  const form = e.target;
  if (!form || !form.matches('form')) return;
  if (form.getAttribute('onsubmit')) return;
  const action = form.getAttribute('action') || '';
  const confirmText = form.getAttribute('data-confirm') || '';
  if (confirmText && !confirm(confirmText)) return;
  const titles = {
    '/check': '检测令牌状态',
    '/refresh': '刷新令牌',
    '/check_all': '批量检测令牌状态',
    '/refresh_all': '批量刷新令牌',
    '/bulk_category': '批量设置分类',
    '/bulk_delete': '批量删除邮箱',
    '/token_selected': '获取已选邮箱令牌',
    '/refresh_selected': '刷新已选邮箱令牌',
    '/check_selected': '检查已选邮箱状态',
    '/inspect_selected': '批量综合检测',
    '/inspect_account': '账号综合检测'
  };
  if (!Object.prototype.hasOwnProperty.call(titles, action)) return;
  e.preventDefault();
  runModalRequest(action, {method: 'POST', headers: {'Content-Type': 'application/x-www-form-urlencoded'}, body: new URLSearchParams(new FormData(form)).toString()}, titles[action]);
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') closeResultModal();
});
</script>'''
    subtitles = {
        'dashboard': '总览账号、API、项目和最近状态。',
        'version': '查看版本、仓库和安全更新状态。',
        'tokens': '刷新 Refresh Token / Access Token，临时使用默认不保存。',
        'mailboxes': '导入、分类、检测和管理 Outlook / Hotmail 邮箱。',
        'mails': '选择邮箱后手动读取验证码和最新邮件摘要。',
        'api_key': '创建和管理外部查询 API 的访问密钥。',
        'api': '查看外部接口、调用方式和安全说明。',
        'projects': '维护项目分类和邮件筛选规则。',
        'categories': '管理邮箱分类标签。',
        'help': '查看使用流程、状态说明和常见问题。',
    }
    page_head = '<div class="app-page-head"><div><div class="app-page-title">' + html.escape(title) + '</div><div class="app-page-subtitle">' + html.escape(subtitles.get(active, '一点微软工具箱后台。')) + '</div></div><div class="app-page-badge">Private Toolbox</div></div>'
    return page(title, f'<div class="app-layout">{sidebar(active)}<main>{page_head}{content}</main></div>{result_modal}')


def row_get(row, key, default=''):
    try:
        value = row[key]
    except (KeyError, IndexError):
        return default
    return default if value is None else value


def account_operations(a):
    account_id = html.escape(str(a['id']))
    category = html.escape(row_get(a, 'category', ''))
    return (
        f'<a class="mini-btn" href="/select?id={account_id}">选中</a> '
        f'<a class="mini-btn primary" href="/token?id={account_id}">获取令牌</a> '
        f'<form method="post" action="/check" style="display:inline"><input type="hidden" name="id" value="{account_id}"><button type="submit" class="mini-btn primary">检测</button></form> '
        f'<form method="post" action="/refresh" style="display:inline"><input type="hidden" name="id" value="{account_id}"><button type="submit" class="mini-btn success">刷新</button></form> '
        f'<form method="post" action="/inspect_account" style="display:inline"><input type="hidden" name="id" value="{account_id}"><button type="submit" class="mini-btn warning">综合检测</button></form> '
        f'<button type="button" onclick="openCategoryModal(&quot;{account_id}&quot;, &quot;{category}&quot;)" class="mini-btn">改分类</button> '
        f'<form method="post" action="/delete" style="display:inline" onsubmit="return confirm(&quot;删除这个账号？&quot;)"><input type="hidden" name="id" value="{account_id}"><button type="submit" class="mini-btn danger">删除</button></form>'
    )

def account_status_badge(value: str) -> str:
    label = html.escape(status_label(value or ''))
    cls = 'ok' if account_status_is_ok(value or '') else ('bad' if account_status_is_error(value or '') else '')
    return f'<span class="status-badge {cls}">{label or "未检测"}</span>'


def account_details(a) -> str:
    raw_err = (a.get('last_error') or '').strip()
    err = html.escape(raw_err or '-')
    advice_title, advice_text, advice_code = error_advice(raw_err, a.get('last_status') or '')
    advice_html = ''
    if raw_err or account_status_is_error(a.get('last_status') or ''):
        advice_html = (
            '<div class="detail-label">错误类型</div><div class="detail-value"><span class="status-badge bad">' + html.escape(advice_title) + '</span></div>'
            '<div class="detail-label">处理建议</div><div class="detail-value">' + html.escape(advice_text) + '</div>'
        )
    return (
        '<details class="mailbox-detail"><summary>详情</summary><div class="detail-grid">'
        '<div class="detail-label">密码</div><div class="detail-value">' + html.escape(a.get('password_mask') or '-') + '</div>'
        '<div class="detail-label">Client ID</div><div class="detail-value"><code>' + html.escape(a.get('client_id') or '-') + '</code></div>'
        '<div class="detail-label">Token</div><div class="detail-value"><code>' + html.escape(a.get('token_mask') or '-') + '</code></div>'
        '<div class="detail-label">最近错误</div><div class="detail-value bad">' + err + '</div>'
        + advice_html +
        '</div></details>'
    )


def render_mailboxes_page(active_account_id: str = '', category_filter: str = '', page_num: int = 1, per_page: int = 50, status_filter: str = '', error_type: str = '', q: str = ''):
    page_meta = paged_saved_accounts(category_filter, page_num, per_page, status_filter, error_type, q)
    accounts = page_meta['rows']
    categories = saved_categories()
    category_links = ''.join(
        f'<a class="chip {"active" if c == category_filter else ""}" href="{html.escape(mailbox_page_url(c, 1, page_meta["per_page"], status_filter, error_type, q))}">{html.escape(c)}</a>'
        for c in categories
    )
    pagination_html = render_pagination(page_meta, category_filter, status_filter, error_type, q)
    category_options = ''.join(f'<option value="{html.escape(c)}"></option>' for c in categories)
    account_rows = ''.join(
        '<tr class="' + ('ok' if str(a["id"]) == str(active_account_id) else '') + '">'
        + '<td><input type="checkbox" name="ids" value="' + str(a["id"]) + '" class="account-check"></td>'
        + '<td>' + str(a["id"]) + '</td>'
        + '<td><div class="mailbox-email">' + html.escape(a["email"]) + '</div><div class="mailbox-meta"><span class="chip">' + html.escape(row_get(a, "category", "未分类") or "未分类") + '</span></div></td>'
        + '<td data-status-cell="' + str(a["id"]) + '">' + account_status_badge(a["last_status"] or "") + '</td>'
        + '<td><div class="ops-wrap">' + account_operations(a) + '</div></td>'
        + '<td>' + account_details(a) + '</td>'
        + '</tr>'
        for a in accounts
    )
    content = f'''
<section class="section-stack">
  <div class="card">
    <h3>邮箱管理</h3>
    <p class="muted">通过弹窗导入单个或批量 Outlook 邮箱账号，并按分类管理。</p>
    <div class="category-bar"><a class="chip {'' if category_filter else 'active'}" href="{html.escape(mailbox_page_url('', 1, page_meta['per_page'], status_filter, error_type, q))}">全部</a>{category_links}</div>
    <form method="get" action="/mailboxes" class="filter-bar">
      {('<input type="hidden" name="category" value="' + html.escape(category_filter) + '">') if category_filter else ''}
      <input name="q" value="{html.escape(q)}" placeholder="搜索邮箱 / 分类 / 错误" style="min-width:220px">
      <select name="status">
        <option value="" {"selected" if not status_filter else ""}>全部状态</option>
        <option value="ok" {"selected" if status_filter == "ok" else ""}>可用</option>
        <option value="error" {"selected" if status_filter == "error" else ""}>异常</option>
        <option value="unchecked" {"selected" if status_filter == "unchecked" else ""}>未检测</option>
        <option value="token_failed" {"selected" if status_filter == "token_failed" else ""}>令牌失效</option>
        <option value="all_failed" {"selected" if status_filter == "all_failed" else ""}>全部失败</option>
      </select>
      <select name="error_type">
        <option value="" {"selected" if not error_type else ""}>全部错误类型</option>
        <option value="abuse" {"selected" if error_type == "abuse" else ""}>账号风控 / AADSTS70000</option>
        <option value="invalid_grant" {"selected" if error_type == "invalid_grant" else ""}>令牌失效</option>
        <option value="consent" {"selected" if error_type == "consent" else ""}>授权或 scope</option>
        <option value="imap_disabled" {"selected" if error_type == "imap_disabled" else ""}>IMAP 不可用</option>
        <option value="login_failed" {"selected" if error_type == "login_failed" else ""}>密码登录失败</option>
        <option value="network" {"selected" if error_type == "network" else ""}>网络/服务异常</option>
        <option value="no_token" {"selected" if error_type == "no_token" else ""}>缺少凭证</option>
      </select>
      <input type="hidden" name="per_page" value="{page_meta['per_page']}">
      <button type="submit" class="mini-btn primary">筛选</button>
      <a class="mini-btn" href="/mailboxes?per_page={page_meta['per_page']}">清空</a>
    </form>
    <div class="toolbar">
      <button type="button" onclick="openModal('single-import-modal')">单个导入</button>
      <button type="button" style="background:#0f766e" onclick="openModal('batch-import-modal')">批量导入</button>
      <a class="button-link" href="/export.txt">批量导出 .txt</a>
    </div>
  </div>
  <div class="card scroll"><h3>已保存邮箱{(' · ' + html.escape(category_filter)) if category_filter else ''}</h3>
      <div class="bulk-bar compact">
        <button type="button" class="mini-btn" onclick="setAllAccountChecks(true)">全选</button>
        <button type="button" class="mini-btn" onclick="invertAccountChecks()">反选</button>
        <button type="button" class="mini-btn" onclick="setAllAccountChecks(false)">取消</button>
        <span class="muted">已选 <b id="selected-count">0</b> 个</span>
        <span class="bulk-spacer"></span>
        <button type="button" class="mini-btn primary" onclick="openBulkActionModal()">批量操作</button>
      </div>
      <div class="table-note"><span class="muted">列表已分页：常用操作外露，Client ID / Token / 错误收进“详情”。</span><span class="muted">当前页 {len(accounts)} 个 / 共 {page_meta['total']} 个邮箱</span></div>
      {pagination_html}
      <table><tr><th><input type="checkbox" onclick="setAllAccountChecks(this.checked)"></th><th>ID</th><th>邮箱</th><th>状态</th><th>操作</th><th>详情</th></tr>{account_rows}</table>
      {pagination_html}
  </div>
</section>

<div id="bulk-action-modal" class="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3>已勾选邮箱批量操作</h3><button type="button" class="modal-close" onclick="closeModal('bulk-action-modal')">关闭</button></div>
    <p class="muted">下面所有操作只会作用于邮箱列表中已勾选的邮箱，不会影响未勾选账号。单个账号可通过列表行里的 <code>action="/inspect_account"</code> 综合检测，结果会显示 Graph 可读、OAuth IMAP 可读、密码 IMAP 可读 等状态。</p>
    <div class="selected-preview">当前已选 <b id="bulk-modal-selected-count">0</b> 个邮箱</div>

    <div class="bulk-action-panel">
      <h4>分类选择</h4>
      <form method="post" action="/bulk_category" id="bulk-category-form">
        <div class="bulk-selected-ids"></div>
        <label>目标分类</label>
        <input name="category" list="bulk-category-options" placeholder="选择或输入目标分类；留空=未分类" class="bulk-category-input">
        <datalist id="bulk-category-options">{category_options}</datalist>
        <button type="submit" style="background:#7c3aed">批量设置分类</button>
      </form>
    </div>

    <div class="bulk-action-panel">
      <h4>令牌操作</h4>
      <div class="action-row">
        <form method="post" action="/token_selected" class="bulk-selected-form"><div class="bulk-selected-ids"></div><button type="submit" style="background:#2563eb">获取令牌</button></form>
        <form method="post" action="/refresh_selected" class="bulk-selected-form"><div class="bulk-selected-ids"></div><button type="submit" style="background:#0f766e">刷新令牌</button></form>
        <form method="post" action="/check_selected" class="bulk-selected-form"><div class="bulk-selected-ids"></div><button type="submit" style="background:#475569">状态检查</button></form>
        <form method="post" action="/inspect_selected" class="bulk-selected-form"><div class="bulk-selected-ids"></div><button type="submit" style="background:#ea580c">综合检测</button></form>
      </div>
      <p class="muted">令牌获取/刷新/状态检查会使用 Microsoft Graph Mail.Read 权限调用 Microsoft 接口，账号较多时请稍等。</p>
    </div>

    <div class="bulk-action-panel">
      <h4>危险操作</h4>
      <form method="post" action="/bulk_delete" class="bulk-selected-form" data-confirm="确定删除已勾选邮箱？此操作不可恢复，建议先确认已经备份。">
        <div class="bulk-selected-ids"></div>
        <button type="submit" style="background:#dc2626">批量删除</button>
      </form>
      <p class="muted">只删除当前已勾选的邮箱记录；删除后不会再出现在邮箱管理、项目取码和 API 查询中。</p>
    </div>
  </div>
</div>

<div id="single-import-modal" class="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3>单个导入</h3><button type="button" class="modal-close" onclick="closeModal('single-import-modal')">关闭</button></div>
    <p class="muted">保存账号后会立即用 Refresh Token 获取 Access Token 并校验可用性。</p>
    <form method="post" action="/decode">
      <label>Client ID</label><input name="client_id" placeholder="9e5f94bc-e8a4-4e73-b8be-63364c29d753" required>
      <label>邮箱地址</label><input name="email" placeholder="name@hotmail.com / name@outlook.com" required>
      <label>Refresh Token</label><textarea name="refresh_token" placeholder="粘贴 refresh token" required></textarea>
      <label>分类</label><input name="category" placeholder="例如：项目A / 热邮箱 / 待检测">
      <button>保存并获取 Access Token</button>
    </form>
  </div>
</div>

<div id="batch-import-modal" class="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3>批量导入</h3><button type="button" class="modal-close" onclick="closeModal('batch-import-modal')">关闭</button></div>
    <p class="muted">格式：邮箱----密码----应用ID----令牌----辅助邮箱----辅助密码----分类；辅助邮箱/密码/分类可留空，也支持 |、逗号、Tab 分隔。</p>
    <form method="post" action="/batch">
      <label>统一分类</label>
      <input name="category" list="category-options" placeholder="选择或输入分类；每行自带分类时优先使用每行分类">
      <datalist id="category-options">{category_options}</datalist>
      <textarea name="batch" placeholder="a@hotmail.com----password----client_id----refresh_token----aux@example.com----aux_password----项目A"></textarea>
      <button>批量保存</button>
    </form>
  </div>
</div>

<div id="category-edit-modal" class="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3>修改分类</h3><button type="button" class="modal-close" onclick="closeModal('category-edit-modal')">关闭</button></div>
    <p class="muted">给当前邮箱设置分类；留空则移到未分类。</p>
    <form method="post" action="/category">
      <input type="hidden" id="category-account-id" name="id" value="">
      <label>分类</label>
      <input id="category-account-value" name="category" list="category-options" placeholder="选择或输入分类；留空=未分类">
      <button type="submit" style="background:#7c3aed">保存分类</button>
    </form>
  </div>
</div>
<script>
function openModal(id) {{ document.getElementById(id).classList.add('active'); }}
function closeModal(id) {{ document.getElementById(id).classList.remove('active'); }}
function openCategoryModal(id, category) {{
  document.getElementById('category-account-id').value = id || '';
  document.getElementById('category-account-value').value = category || '';
  openModal('category-edit-modal');
}}
function closeModalOnBackdrop(e) {{ if (e.target.classList.contains('modal-backdrop')) e.target.classList.remove('active'); }}
function accountChecks() {{ return Array.from(document.querySelectorAll('.account-check')); }}
function selectedAccountChecks() {{ return accountChecks().filter(c => c.checked); }}
function updateSelectedCount() {{
  const count = selectedAccountChecks().length;
  const el = document.getElementById('selected-count');
  const modalEl = document.getElementById('bulk-modal-selected-count');
  if (el) el.textContent = count;
  if (modalEl) modalEl.textContent = count;
}}
function setAllAccountChecks(checked) {{ accountChecks().forEach(c => c.checked = checked); updateSelectedCount(); }}
function invertAccountChecks() {{ accountChecks().forEach(c => c.checked = !c.checked); updateSelectedCount(); }}
function fillBulkSelectedIds() {{
  const boxes = Array.from(document.querySelectorAll('.bulk-selected-ids'));
  boxes.forEach(box => {{
    box.innerHTML = '';
    selectedAccountChecks().forEach(c => {{
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = 'ids';
      input.value = c.value;
      box.appendChild(input);
    }});
  }});
}}
function openBulkActionModal() {{ updateSelectedCount(); fillBulkSelectedIds(); openModal('bulk-action-modal'); }}
document.addEventListener('change', e => {{ if (e.target.classList && e.target.classList.contains('account-check')) updateSelectedCount(); }});
document.addEventListener('submit', e => {{
  if (e.target && (e.target.id === 'bulk-category-form' || e.target.classList.contains('bulk-selected-form'))) {{
    fillBulkSelectedIds();
    if (selectedAccountChecks().length === 0) {{ e.preventDefault(); alert('请先勾选要操作的邮箱。'); }}
  }}
}});
document.addEventListener('keydown', e => {{ if (e.key === 'Escape') document.querySelectorAll('.modal-backdrop.active').forEach(m => m.classList.remove('active')); }});
</script>'''
    return app_page('邮箱管理', 'mailboxes', content)


def render_tokens_page():
    accounts = saved_accounts()
    row_parts = []
    for a in accounts:
        account_id = str(a['id'])
        row_parts.append(
            '<tr><td>' + account_id + '</td>'
            '<td><b>' + html.escape(a['email']) + '</b><br><span class="muted">' + html.escape(a.get('category') or '未分类') + '</span></td>'
            '<td><code>' + html.escape(mask_secret(a['client_id'])) + '</code></td>'
            '<td>' + html.escape(a['token_mask']) + '</td>'
            '<td data-status-cell="' + account_id + '">' + account_status_badge(a['last_status'] or '') + '</td>'
            '<td>' + html.escape((a['last_error'] or '')[:100]) + '</td>'
            '<td><div class="action-row" style="gap:6px">'
            '<a class="mini-btn" href="/token?id=' + account_id + '">获取</a>'
            '<form method="post" action="/check" style="display:inline" onsubmit="runModalRequest(\'/check\', {method:\'POST\', body:new FormData(this)}, \'检测令牌\'); return false">'
            '<input type="hidden" name="id" value="' + account_id + '"><button type="submit" class="mini-btn" style="background:#2563eb!important">检测</button></form>'
            '<form method="post" action="/refresh" style="display:inline" onsubmit="runModalRequest(\'/refresh\', {method:\'POST\', body:new FormData(this)}, \'刷新令牌\'); return false">'
            '<input type="hidden" name="id" value="' + account_id + '"><button type="submit" class="mini-btn" style="background:#0f766e!important">刷新</button></form>'
            '</div></td></tr>'
        )
    token_rows = ''.join(row_parts) or '<tr><td colspan="7" class="muted">还没有保存账号。可以先在上方临时刷新，或到“邮箱管理”导入。</td></tr>'
    content = f'''
<section class="section-stack">
  <div class="toolbox-hero"><div class="toolbox-kicker">Refresh Token</div><h1 class="toolbox-title">刷新令牌工具</h1><p class="toolbox-desc">输入 Client ID 与 Refresh Token，换取 Access Token，并在微软返回新 Refresh Token 时自动识别轮换。默认不保存临时输入，避免敏感数据误入库。</p></div>
  <div class="section-grid">
    <div class="card"><h3>单次刷新 / 获取 Token</h3>
      <form method="post" action="/token_tool" onsubmit="runModalRequest('/token_tool', {{method:'POST', body:new FormData(this)}}, '刷新令牌结果'); return false">
        <label>Client ID</label><input name="client_id" placeholder="Microsoft OAuth Client ID" required>
        <label>Refresh Token</label><textarea name="refresh_token" placeholder="粘贴 Refresh Token。提交后仅用于本次请求，默认不保存。" required style="min-height:120px"></textarea>
        <details><summary>高级参数 / 保存选项</summary>
          <label>Tenant</label><select name="tenant"><option value="consumers" selected>consumers - 个人 Outlook/Hotmail</option><option value="common">common</option><option value="organizations">organizations</option></select>
          <label>Scope</label><textarea name="scope" style="min-height:70px">{html.escape(default_oauth_scope())}</textarea>
          <label>邮箱（可选，仅勾选保存时需要）</label><input name="email" placeholder="example@hotmail.com">
          <label>分类（可选）</label><input name="category" placeholder="项目名 / 分类名">
          <label><input type="checkbox" name="save_account" value="1"> 成功后保存/更新到邮箱管理（Refresh Token 会加密入库）</label>
        </details>
        <div class="action-row"><button type="submit">获取 Access Token / 刷新</button><button type="button" style="background:#64748b" onclick="this.closest('form').reset()">清空</button></div>
      </form>
    </div>
    <div class="card"><h3>使用说明</h3>
      <ul class="muted" style="line-height:1.9;margin-top:8px">
        <li>Access Token 有效期通常较短，不建议长期保存。</li>
        <li>微软可能在刷新时返回新的 Refresh Token；如果保存账号，系统会保存最新值。</li>
        <li>返回 <code>invalid_grant</code> 通常表示 Refresh Token 已失效、被撤销或账号触发风控。</li>
        <li>数据库中的 Refresh Token / 密码字段已加密，页面与 API 默认不暴露明文。</li>
      </ul>
      <div class="action-row"><button type="button" style="background:#7c3aed" onclick="openModal('token-bulk-modal')">批量操作</button><a class="mini-btn" href="/mailboxes">导入邮箱</a><a class="mini-btn" href="/api-manage">查看 API</a></div>
    </div>
  </div>
  <div class="card scroll"><h3>已保存账号令牌列表</h3><p class="muted">单账号检测/刷新会以弹窗显示结果；“获取”会打开一次性 token 结果页。</p><table><tr><th>ID</th><th>邮箱</th><th>Client ID</th><th>Token</th><th>状态</th><th>错误</th><th>操作</th></tr>{token_rows}</table></div>
</section>
<div id="token-bulk-modal" class="modal-backdrop" onclick="closeModalOnBackdrop(event)">
  <div class="modal">
    <div class="modal-head"><h3>令牌批量操作</h3><button type="button" class="modal-close" onclick="closeModal('token-bulk-modal')">关闭</button></div>
    <p class="muted">这里会对全部已保存账号执行操作。批量刷新可能耗时较长，建议账号很多时分批执行。</p>
    <div class="action-row">
      <form method="post" action="/check_all" onsubmit="runModalRequest('/check_all', {{method:'POST', body:new FormData(this)}}, '批量检测结果'); return false"><button style="background:#2563eb">批量检测令牌状态</button></form>
      <form method="post" action="/refresh_all" onsubmit="if(!confirm('将刷新所有已保存账号的令牌，继续？')) return false; runModalRequest('/refresh_all', {{method:'POST', body:new FormData(this)}}, '批量刷新结果'); return false"><button style="background:#0f766e">批量刷新全部令牌</button></form>
    </div>
  </div>
</div>'''
    return app_page('令牌管理', 'tokens', content)

def render_categories_page():
    rows, uncategorized = categories_with_counts()
    row_html = ''.join(
        f'<tr><td>{r["id"]}</td><td>{html.escape(r["name"])}</td><td>{r["account_count"]}</td><td>'
        f'<form method="post" action="/category_rename" style="display:inline"><input type="hidden" name="id" value="{r["id"]}"><input name="name" value="{html.escape(r["name"])}" style="width:160px;padding:5px 8px"><button type="submit" style="padding:4px 8px;background:#7c3aed">改名</button></form> '
        f'<form method="post" action="/category_delete" style="display:inline" onsubmit="return confirm(&quot;删除分类后，该分类下邮箱会变成未分类，继续？&quot;)"><input type="hidden" name="id" value="{r["id"]}"><button type="submit" style="padding:4px 8px;background:#dc2626">删除</button></form> '
        f'<a href="/mailboxes?category={urllib.parse.quote(r["name"])}">查看邮箱</a>'
        f'</td></tr>'
        for r in rows
    )
    content = f'''
<section class="section-stack">
  <div class="card"><h3>分类管理</h3><p class="muted">新增、改名、删除邮箱分类。删除分类不会删除邮箱，只会把该分类下邮箱移到未分类。</p>
    <form method="post" action="/category_create" class="action-row">
      <input name="name" placeholder="新分类名称，例如：项目A / 热邮箱 / 待检测" required style="max-width:360px">
      <button type="submit">新增分类</button>
    </form>
  </div>
  <div class="card scroll"><h3>分类列表</h3><p class="muted">未分类账号：{uncategorized} 个 · <a href="/mailboxes?category={urllib.parse.quote('未分类')}">查看未分类</a></p>
    <table><tr><th>ID</th><th>分类名称</th><th>邮箱数量</th><th>操作</th></tr>{row_html}</table>
  </div>
</section>'''
    return app_page('分类管理', 'categories', content)


def render_mails_page(active_account_id: str = ''):
    accounts = saved_accounts()
    active_email = ''
    for a in accounts:
        if str(a['id']) == str(active_account_id):
            active_email = a['email']
            break
    account_options = ''.join(
        '<a class="mail-picker-item ' + ('active' if str(a["id"]) == str(active_account_id) else '') + '" data-mailbox-row data-search="' + html.escape((str(a["id"]) + ' ' + a["email"] + ' ' + (a.get("category") or "未分类")).lower(), quote=True) + '" href="/select?id=' + str(a["id"]) + '&next=/mails">'
        + '<span><span class="mail-picker-email">' + html.escape(a["email"]) + '</span><span class="mail-picker-meta">#' + str(a["id"]) + ' · ' + html.escape(a.get("category") or "未分类") + ' · ' + html.escape(status_label(a.get("last_status") or "")) + '</span></span>'
        + '<span class="mail-picker-action">读取</span></a>'
        for a in accounts
    )
    content = '<section class="section-grid">'
    content += '<div class="card scroll"><div class="mail-search"><h3>选择邮箱</h3><p class="muted">先选择一个邮箱，再读取验证码和最新邮件摘要。</p><input id="mailbox-filter" oninput="filterMailboxes()" placeholder="搜索邮箱 / ID / 分类"><p class="muted">显示 <b id="mailbox-visible-count">' + str(len(accounts)) + '</b> / ' + str(len(accounts)) + ' 个邮箱</p></div><div class="mail-picker-list">' + account_options + '</div></div>'
    content += '<div class="card"><div class="mail-result-top"><div><h3>验证码邮件</h3><p class="muted">当前邮箱：<b id="active-email">' + html.escape(active_email or '未选择') + '</b> · 状态：<span id="poll-status">等待手动获取</span></p></div><button type="button" class="mini-btn primary" onclick="pollCodes()">手动获取邮件</button></div><div id="codes" class="mail-result-area"><div class="empty-state">不会自动刷新；请选择邮箱后点击“手动获取邮件”。</div></div></div></section>'
    content += '\n<script>\nconst SOURCE_LABELS = {graph: \'Graph 成功\', outlook_rest: \'Outlook REST 成功\', imap: \'OAuth IMAP 成功\', imap_password: \'密码 IMAP 成功\'};\nasync function pollCodes() {\n  const box = document.getElementById(\'codes\');\n  const status = document.getElementById(\'poll-status\');\n  try {\n    status.textContent = \'获取中 \' + new Date().toLocaleTimeString();\n    const res = await fetch(\'/api/codes\', {cache: \'no-store\'});\n    const data = await res.json();\n    let out = \'\';\n    if (!(data.results || []).length) out = \'<div class="empty-state">请先在左侧选择一个邮箱。</div>\';\n    for (const item of data.results || []) {\n      const sourceLabel = SOURCE_LABELS[item.source_api] || item.source_api || \'未知通道\';\n      if (item.status !== \'ok\') {\n        out += `<div class="mail-summary-card bad"><h3>读取失败</h3><p><b>${escapeHtml(item.email || \'\')}</b></p><p>${escapeHtml(item.error || \'读取失败\')}</p><p class="muted">建议到“邮箱管理”里点“综合检测”，或换一个账号测试。</p></div>`;\n        continue;\n      }\n      const codeCount = (item.codes || []).length;\n      const mailCount = Number(item.mail_count || 0);\n      out += `<div class="mail-summary-card ok"><h3>${escapeHtml(sourceLabel)}</h3><p><b>${escapeHtml(item.email || \'\')}</b></p><div class="metric-row"><span class="metric-chip">读取邮件 ${mailCount} 封</span><span class="metric-chip">验证码 ${codeCount} 条</span><span class="metric-chip">手动获取</span></div></div>`;\n      if (codeCount) {\n        out += \'<h4>验证码</h4><div class="code-grid">\';\n        for (const c of item.codes || []) out += `<div class="code-card"><div class="code-card-head"><span class="code-pill">${escapeHtml(c.code || \'\')}</span><button type="button" class="mini-btn primary" onclick="copyText(this.dataset.copy)" data-copy="${escapeAttr(c.code || \'\')}">复制验证码</button></div><p class="muted">邮箱位置：${escapeHtml(c.folder || \'\')} · 来源：${escapeHtml(c.source || \'\')} · 时间：${escapeHtml(c.date || \'\')}</p><div class="mail-subject">${escapeHtml(c.subject || \'\')}</div></div>`;\n        out += \'</div>\';\n      } else {\n        out += \'<div class="empty-state">通道读取成功，但最近邮件里没有识别到验证码。下面显示最新邮件摘要。</div>\';\n      }\n      if ((item.latest_mails || []).length) {\n        out += \'<h4>最新邮件摘要</h4><div class="mail-card-grid">\';\n        for (const m of item.latest_mails || []) out += `<div class="mail-card"><div class="mail-subject">${escapeHtml(m.subject || \'(无主题)\')}</div><p class="muted">邮箱位置：${escapeHtml(m.folder || \'\')} · 发件人：${escapeHtml(m.source || \'\')} · 时间：${escapeHtml(m.date || \'\')}</p><div class="mail-preview">${escapeHtml(m.preview || \'\').slice(0,260)}</div></div>`;\n        out += \'</div>\';\n      }\n    }\n    box.innerHTML = out; status.textContent = \'已更新 \' + new Date().toLocaleTimeString();\n  } catch (e) { status.textContent = \'失败 \' + e; box.innerHTML = \'<div class="mail-summary-card bad"><h3>请求失败</h3><p>\' + escapeHtml(String(e)) + \'</p></div>\'; }\n}\nfunction filterMailboxes() {\n  const q = (document.getElementById(\'mailbox-filter\').value || \'\').toLowerCase().trim();\n  let visible = 0;\n  document.querySelectorAll(\'[data-mailbox-row]\').forEach(row => {\n    const hit = !q || (row.dataset.search || \'\').includes(q);\n    row.style.display = hit ? \'\' : \'none\';\n    if (hit) visible++;\n  });\n  const el = document.getElementById(\'mailbox-visible-count\');\n  if (el) el.textContent = visible;\n}\nfunction escapeHtml(s) { return String(s || \'\').replace(/[&<>"]/g, c => ({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\'}[c])); }\nfunction escapeAttr(s) { return String(s || \'\').replace(/[\\\\\'"<>]/g, \'\'); }\nfunction copyText(s) { navigator.clipboard && navigator.clipboard.writeText(s || \'\'); }\ndocument.getElementById(\'poll-status\').textContent = \'等待手动获取\';\n</script>'
    return app_page('邮件获取', 'mails', content)


def render_public_home_page():
    body = '''
<style>
.public-home{min-height:100vh;background:#f6f8ff;color:#182033;overflow:hidden}.public-nav{position:sticky;top:0;z-index:10;backdrop-filter:blur(16px);background:#fffffff0;border-bottom:1px solid #e7eaf3}.public-nav-inner{max-width:1180px;margin:0 auto;padding:14px 22px;display:flex;align-items:center;justify-content:space-between;gap:18px}.public-logo{display:flex;align-items:center;gap:10px;font-weight:900;font-size:20px;color:#2563eb;text-decoration:none}.public-links{display:flex;gap:24px;align-items:center}.public-links a{color:#475569;text-decoration:none;font-weight:700}.public-links a:hover{color:#2563eb}.public-actions{display:flex;gap:10px}.public-btn{display:inline-flex;align-items:center;justify-content:center;border-radius:999px;padding:10px 18px;text-decoration:none;font-weight:900;border:1px solid #dbe3f0;color:#1e293b;background:#fff}.public-btn.primary{background:linear-gradient(135deg,#2563eb,#7c3aed);color:#fff;border-color:transparent;box-shadow:0 12px 30px #2563eb33}.public-hero{position:relative;max-width:1180px;margin:0 auto;padding:88px 22px 58px;display:grid;grid-template-columns:1.05fr .95fr;gap:42px;align-items:center}.public-hero:before{content:'';position:absolute;right:-12%;top:8%;width:520px;height:520px;border-radius:999px;background:linear-gradient(135deg,#bfdbfe,#ddd6fe);filter:blur(22px);opacity:.75}.public-kicker{display:inline-flex;gap:8px;align-items:center;background:#e0ecff;color:#1d4ed8;border:1px solid #bfdbfe;border-radius:999px;padding:8px 13px;font-weight:900}.public-title{font-size:58px;line-height:1.02;letter-spacing:-.06em;margin:22px 0 18px;color:#0f172a}.public-title span{background:linear-gradient(135deg,#2563eb,#7c3aed);-webkit-background-clip:text;background-clip:text;color:transparent}.public-desc{font-size:18px;line-height:1.85;color:#475569;max-width:640px}.public-cta{display:flex;gap:14px;flex-wrap:wrap;margin-top:30px}.public-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:34px}.metric{background:#fff;border:1px solid #e7eaf3;border-radius:18px;padding:18px;box-shadow:0 10px 28px #1e293b12}.metric b{font-size:24px;color:#2563eb}.mock-panel{position:relative;background:#0f172a;color:#e5e7eb;border-radius:30px;padding:18px;border:1px solid #1e293b;box-shadow:0 26px 80px #0f172a33}.mock-top{display:flex;gap:7px;margin-bottom:16px}.mock-dot{width:10px;height:10px;border-radius:999px;background:#64748b}.mock-card{background:#111827;border:1px solid #263244;border-radius:18px;padding:16px;margin-top:12px}.mock-row{display:flex;justify-content:space-between;gap:12px;border-bottom:1px solid #263244;padding:11px 0}.mock-row:last-child{border-bottom:0}.features,.plans,.about{max-width:1180px;margin:0 auto;padding:54px 22px}.section-title{text-align:center;font-size:36px;letter-spacing:-.04em;margin:0 0 10px;color:#0f172a}.section-sub{text-align:center;color:#64748b;font-size:16px;margin:0 0 34px}.feature-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}.feature-card,.plan-card{background:#fff;border:1px solid #e7eaf3;border-radius:24px;padding:24px;box-shadow:0 16px 40px #1e293b10}.feature-icon{font-size:34px;margin-bottom:16px}.feature-card h3,.plan-card h3{color:#0f172a;font-size:20px;margin:0 0 10px}.feature-card p,.plan-card p{color:#64748b;line-height:1.75;margin:0}.plan-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}.plan-card.featured{border-color:#93c5fd;box-shadow:0 20px 60px #2563eb22;transform:translateY(-6px)}.price{font-size:34px;font-weight:900;color:#2563eb;margin:16px 0}.plan-card ul{margin:16px 0 0;padding:0;list-style:none;display:grid;gap:10px;color:#475569}.plan-card li:before{content:'✓';color:#16a34a;font-weight:900;margin-right:8px}.public-footer{margin-top:34px;padding:28px 22px;background:#0f172a;color:#cbd5e1;text-align:center}.public-footer a{color:#93c5fd}.about-box{background:linear-gradient(135deg,#fff,#eef4ff);border:1px solid #dbeafe;border-radius:28px;padding:32px;color:#475569;line-height:1.9;box-shadow:0 18px 50px #1e293b10}@media(max-width:900px){.public-hero{grid-template-columns:1fr;padding-top:48px}.mock-panel{display:none}.public-title{font-size:42px}.feature-grid,.plan-grid,.public-metrics{grid-template-columns:1fr}.public-links{display:none}.public-nav-inner{padding:12px 16px}}
</style>
<div class="public-home"><nav class="public-nav"><div class="public-nav-inner"><a class="public-logo" href="/">🧰 一点微软工具箱</a><div class="public-links"><a href="#features">功能特色</a><a href="#plans">套餐/能力</a><a href="#about">关于</a></div><div class="public-actions"><a class="public-btn" href="/login">登录</a><a class="public-btn primary" href="/login">进入控制台</a></div></div></nav><section class="public-hero"><div><div class="public-kicker">Self-hosted Outlook / Graph API</div><h1 class="public-title">专业级 <span>一点微软工具箱</span></h1><p class="public-desc">一站式管理 Outlook / Hotmail 邮箱凭证、Refresh Token、Access Token 刷新和验证码邮件读取。风格对齐工具箱产品页，但数据仍保留在你的服务器本地。</p><div class="public-cta"><a class="public-btn primary" href="/login">立即进入后台</a><a class="public-btn" href="#features">查看功能</a></div><div class="public-metrics"><div class="metric"><b>Graph</b><br>优先读取邮件</div><div class="metric"><b>批量</b><br>导入/检测/刷新</div><div class="metric"><b>本地</b><br>SQLite 私有存储</div></div></div><div class="mock-panel"><div class="mock-top"><span class="mock-dot"></span><span class="mock-dot"></span><span class="mock-dot"></span></div><div class="mock-card"><h3>控制台预览</h3><div class="mock-row"><span>刷新令牌</span><b class="ok">可用</b></div><div class="mock-row"><span>邮箱管理</span><b>批量导入</b></div><div class="mock-row"><span>获取邮件</span><b>验证码提取</b></div><div class="mock-row"><span>查询 API</span><b>私有接口</b></div></div></div></section><section id="features" class="features"><h2 class="section-title">核心功能</h2><p class="section-sub">把常用微软邮箱操作集中到一个后台</p><div class="feature-grid"><div class="feature-card"><div class="feature-icon">🔄</div><h3>自动刷新令牌</h3><p>检测 Refresh Token 状态，获取 Access Token，支持批量刷新并保存轮换后的新令牌。</p></div><div class="feature-card"><div class="feature-icon">📮</div><h3>邮箱账号托管</h3><p>支持单个/批量导入，分类管理，批量导出，适合多账号项目维护。</p></div><div class="feature-card"><div class="feature-icon">📥</div><h3>邮件实时获取</h3><p>通过 Microsoft Graph 优先读取收件箱和垃圾箱，提取验证码摘要。</p></div><div class="feature-card"><div class="feature-icon">🏷️</div><h3>项目化管理</h3><p>按项目、用途或状态分类邮箱，后续可继续扩展项目规则和过滤器。</p></div><div class="feature-card"><div class="feature-icon">🔌</div><h3>接口服务</h3><p>保留私有 API 能力，可继续扩展 API Key、PT 码或项目查询接口。</p></div><div class="feature-card"><div class="feature-icon">🛡️</div><h3>安全可靠</h3><p>双层登录、HTTPS、HttpOnly Cookie、本地数据库，不公开敏感 token。</p></div></div></section><section id="plans" class="plans"><h2 class="section-title">私有版能力</h2><p class="section-sub">不是照抄第三方商业站，而是做成你自己的私有工具箱</p><div class="plan-grid"><div class="plan-card"><h3>基础管理</h3><div class="price">已支持</div><ul><li>邮箱导入导出</li><li>分类管理</li><li>安全登录</li></ul></div><div class="plan-card featured"><h3>令牌/邮件</h3><div class="price">已支持</div><ul><li>Refresh Token 检测</li><li>Access Token 获取</li><li>验证码邮件读取</li></ul></div><div class="plan-card"><h3>接口扩展</h3><div class="price">可继续做</div><ul><li>API Key</li><li>项目规则</li><li>外部查询 API</li></ul></div></div></section><section id="about" class="about"><div class="about-box"><h2>关于这个工具箱</h2><p>这是部署在 token.seoyh.net 的一点微软工具箱。当前版本主要面向你自己使用：进入控制台后可管理账号、刷新令牌、读取邮件。后续如果你想做成完整商业版，我可以继续补注册、套餐、API Key、项目管理和外部接口鉴权。</p></div></section><footer class="public-footer">© 一点微软工具箱 · <a href="https://www.seoyh.net/" target="_blank" rel="noopener noreferrer">一点优化</a> · <a href="/login">登录控制台</a></footer></div>'''
    return page('一点微软工具箱 - 专业邮箱 API 管理平台', body, show_nav=False)



def render_api_key_page(new_key: str = ''):
    keys = list_api_keys()
    logs = recent_api_logs(20)
    stats = api_usage_stats()
    new_key_html = ''
    if new_key:
        health_cmd = 'curl -H "X-API-Key: ' + new_key + '" https://token.seoyh.net/api/v1/health'
        new_key_html = (
            '<div class="card"><h3 class="ok">新 API Key 已生成</h3>'
            '<p class="muted">请立刻复制保存，之后页面不会再显示完整密钥。推荐通过请求头 <code>X-API-Key</code> 使用。</p>'
            '<textarea id="new-api-key" readonly style="min-height:58px">' + html.escape(new_key) + '</textarea>'
            '<textarea id="new-api-health-curl" readonly style="min-height:72px">' + html.escape(health_cmd) + '</textarea>'
            '<div class="action-row"><button type="button" onclick="copyToClipboard(\'new-api-key\')">一键复制密钥</button>'
            '<button type="button" style="background:#0f766e" onclick="copyToClipboard(\'new-api-health-curl\')">复制健康检查示例</button></div></div>'
        )
    row_parts = []
    for k in keys:
        st = stats.get(k['id'])
        last_minute = int(st['last_minute'] or 0) if st else 0
        today = int(st['today'] or 0) if st else 0
        today_ok = int(st['today_ok'] or 0) if st else 0
        today_fail = int(st['today_fail'] or 0) if st else 0
        row_parts.append('<tr><td>' + str(k['id']) + '</td><td>' + html.escape(k['name']) + '</td><td><code>' + html.escape(k['key_prefix']) + '…</code></td><td>' + ('启用' if k['enabled'] else '停用') + '</td><td><code>' + html.escape(k['scopes'] or '') + '</code></td><td>' + str(last_minute) + '/' + str(API_RATE_LIMIT_PER_MINUTE) + '</td><td>' + str(today) + '/' + str(API_RATE_LIMIT_PER_DAY) + '<br><span class="ok">成功 ' + str(today_ok) + '</span> · <span class="bad">失败 ' + str(today_fail) + '</span></td><td>' + html.escape(datetime.fromtimestamp(k['last_used_at']).strftime('%Y-%m-%d %H:%M') if k['last_used_at'] else '-') + '</td><td>' + html.escape(k['note'] or '') + '</td><td><form method="post" action="/api_key_toggle" style="display:inline"><input type="hidden" name="id" value="' + str(k['id']) + '"><input type="hidden" name="enabled" value="' + ('0' if k['enabled'] else '1') + '"><button type="submit" class="mini-btn">' + ('停用' if k['enabled'] else '启用') + '</button></form> <form method="post" action="/api_key_delete" style="display:inline" onsubmit="return confirm(&quot;删除这个 API Key？&quot;)"><input type="hidden" name="id" value="' + str(k['id']) + '"><button type="submit" class="mini-btn" style="background:#dc2626!important">删除</button></form></td></tr>')
    rows = ''.join(row_parts) or '<tr><td colspan="10" class="muted">还没有 API Key。</td></tr>'
    log_rows = ''.join(
        '<tr><td>' + html.escape(datetime.fromtimestamp(l['created_at']).strftime('%m-%d %H:%M:%S')) + '</td><td>' + html.escape(l['key_name'] or '-') + '</td><td>' + html.escape(l['endpoint']) + '</td><td>' + html.escape(l['status']) + '</td><td>' + html.escape(l['detail'] or '') + '</td></tr>'
        for l in logs
    ) or '<tr><td colspan="5" class="muted">暂无调用日志。</td></tr>'
    content = f'''
<section class="section-stack">
  <div class="toolbox-hero"><div class="toolbox-kicker">🔐 API Key</div><h1 class="toolbox-title">API 密钥</h1><p class="toolbox-desc">为外部系统生成独立 API Key。接口只返回验证码/状态等必要信息，不返回 Refresh Token 或 Access Token 原文。</p></div>
  {new_key_html}
  <div class="card"><h3>创建 API Key</h3><form method="post" action="/api_key_create"><div class="action-row"><input name="name" placeholder="名称，例如：项目A" style="max-width:260px"><input name="note" placeholder="备注，可选" style="max-width:360px"></div><p class="muted">权限范围：</p><div class="action-row"><label><input type="checkbox" name="scopes" value="health" checked> health</label><label><input type="checkbox" name="scopes" value="latest_code" checked> latest_code</label><label><input type="checkbox" name="scopes" value="projects" checked> projects</label><label><input type="checkbox" name="scopes" value="accounts"> accounts</label><button type="submit">生成密钥</button></div><p class="muted">建议只给外部系统必要权限；accounts 会暴露邮箱地址元数据，默认不勾选。</p></form></div>
  <div class="card scroll"><h3>密钥列表</h3><p class="muted">当前限流：每分钟 {API_RATE_LIMIT_PER_MINUTE} 次 / 每天 {API_RATE_LIMIT_PER_DAY} 次。达到限制会返回 429。Query key 开关：{'启用兼容' if API_ALLOW_QUERY_KEY else '已禁用'}。</p><table><tr><th>ID</th><th>名称</th><th>前缀</th><th>状态</th><th>权限</th><th>近1分钟</th><th>今日调用</th><th>最后使用</th><th>备注</th><th>操作</th></tr>{rows}</table></div>
  <div class="card scroll"><h3>最近调用日志</h3><table><tr><th>时间</th><th>密钥</th><th>接口</th><th>状态</th><th>详情</th></tr>{log_rows}</table></div>
</section>'''
    return app_page('API 密钥', 'api_key', content)


def render_project_manage_page():
    rows, uncategorized = categories_with_counts()
    rules = project_rules_map()
    project_rows = []
    for r in rows:
        rule = rules.get(r['name'])
        sender = html.escape((rule['sender_keywords'] if rule else '') or '', quote=True)
        subject = html.escape((rule['subject_keywords'] if rule else '') or '', quote=True)
        body = html.escape((rule['body_keywords'] if rule else '') or '', quote=True)
        max_results = html.escape(str(rule['max_results'] if rule else 5), quote=True)
        checked = ' checked' if (not rule or rule['enabled']) else ''
        project_rows.append('<tr><td>' + str(r['id']) + '</td><td><b>' + html.escape(r['name']) + '</b><br><span class="muted">邮箱 ' + str(r['account_count']) + ' 个</span></td><td><form method="post" action="/project_rule_save"><input type="hidden" name="category" value="' + html.escape(r['name'], quote=True) + '"><input name="sender_keywords" value="' + sender + '" placeholder="发件人关键词"><input name="subject_keywords" value="' + subject + '" placeholder="标题关键词"><input name="body_keywords" value="' + body + '" placeholder="正文关键词"><input name="max_results" type="number" min="1" max="50" value="' + max_results + '" style="max-width:120px"><label><input class="account-check" type="checkbox" name="enabled" value="1"' + checked + '> 启用</label><button type="submit">保存规则</button> <a href="/mailboxes?category=' + urllib.parse.quote(r['name']) + '">查看邮箱</a></form></td></tr>')
    project_rows = ''.join(project_rows) or '<tr><td colspan="3" class="muted">还没有项目。可以先在邮箱管理里设置分类，或在这里创建项目。</td></tr>'
    content = f'''
<section class="section-stack">
  <div class="toolbox-hero"><div class="toolbox-kicker">📁 Projects</div><h1 class="toolbox-title">项目管理</h1><p class="toolbox-desc">项目 = 邮箱分类 + 查询规则。外部 API 使用 category=项目名 时，会自动套用这里配置的发件人、标题、正文关键词和返回数量。</p></div>
  <div class="card"><h3>新增项目</h3><form method="post" action="/category_create" class="action-row"><input name="name" placeholder="项目名称 / 分类名" style="max-width:320px"><button type="submit">创建项目</button></form><p class="muted">未分类邮箱：{uncategorized} 个 · <a href="/mailboxes?category={urllib.parse.quote('未分类')}">查看</a></p></div>
  <div class="card scroll"><h3>项目规则</h3><p class="muted">关键词支持逗号、分号、竖线或换行分隔。留空表示不过滤该字段。</p><table><tr><th style="width:220px">项目</th><th>查询规则</th></tr>{project_rows}</table></div>
</section>'''
    return app_page('项目管理', 'projects', content)

def render_api_manage_page():
    content = '''
<section class="section-stack">
  <div class="toolbox-hero"><div class="toolbox-kicker">🔌 Private API</div><h1 class="toolbox-title">查询 API</h1><p class="toolbox-desc">这里先做成与参考站一致的 API 入口页。当前私有版接口默认使用登录态保护；如果你要开放给外部系统调用，下一步建议增加 API Key、项目规则和访问频率限制。</p></div>
  <div class="tool-grid">
    <div class="tool-card"><div class="tool-icon">📥</div><h3>获取当前邮箱验证码</h3><p><code>GET /api/codes</code><br>返回后台当前选中邮箱的验证码摘要。</p></div>
    <div class="tool-card"><div class="tool-icon">📤</div><h3>导出账号文本</h3><p><code>GET /export.txt</code><br>导出已保存邮箱账号文本，仅登录后可用。</p></div>
    <div class="tool-card"><div class="tool-icon">🔑</div><h3>令牌刷新接口</h3><p>后台已有刷新逻辑；外部 API 版可继续封装为 API Key 调用。</p></div>
    <div class="tool-card"><div class="tool-icon">🏷️</div><h3>项目过滤接口</h3><p>可按分类/项目扩展查询指定邮箱组的最新邮件。</p></div>
  </div>
  <div class="card"><h3>可用外部接口</h3><pre>curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/latest-code?email=xxx@outlook.com"
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/latest-code?category=项目名"
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/account-status?email=xxx@outlook.com"
curl -H "X-API-Key: 你的APIKEY" https://token.seoyh.net/api/v1/health
curl -H "X-API-Key: 你的APIKEY" https://token.seoyh.net/api/v1/projects
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/accounts?category=项目名"
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/latest-code?category=项目名&amp;limit=5&amp;subject=验证码"</pre><p class="muted">推荐使用请求头：<code>X-API-Key: 你的APIKEY</code>。URL 参数 <code>?key=</code> 暂时兼容但不推荐，可通过环境变量 <code>RTWEB_API_ALLOW_QUERY_KEY=0</code> 禁用。接口按 API Key scopes 授权；latest-code 只返回验证码摘要，account-status 返回综合状态与安全摘要，不返回 Refresh Token / Access Token / 密码明文。</p></div>
</section>'''
    return app_page('查询 API', 'api', content)


def render_help_page():
    content = '<section class="section-stack">\n  <div class="toolbox-hero">\n    <div class="toolbox-kicker">Guide</div>\n    <h1 class="toolbox-title">一点微软工具箱使用说明</h1>\n    <p class="toolbox-desc">推荐流程：导入邮箱 → 检测/刷新令牌 → 综合检测 → 配置项目规则 → 获取验证码/API 调用。本文按实际功能整理，不返回或展示任何账号密码、Refresh Token 明文。</p>\n  </div>\n\n  <div class="quick-actions">\n    <div class="card">\n      <h3>一、导入邮箱</h3>\n      <p class="muted">入口：<code>/mailboxes</code> → “批量导入”。一行一个邮箱。</p>\n      <pre>邮箱----密码----应用ID(Client ID)----Refresh Token----辅助邮箱----辅助密码----分类/项目</pre>\n      <table><tr><th>字段</th><th>是否必填</th><th>说明</th></tr>\n        <tr><td>邮箱</td><td>必填</td><td>Outlook / Hotmail 邮箱地址。</td></tr>\n        <tr><td>密码</td><td>可选</td><td>用于密码 IMAP 兜底读取验证码；不会用于生成 Refresh Token。</td></tr>\n        <tr><td>Client ID</td><td>令牌账号必填</td><td>Microsoft OAuth 应用 ID。</td></tr>\n        <tr><td>Refresh Token</td><td>令牌账号必填</td><td>用于刷新 Access Token，保存后加密入库。</td></tr>\n        <tr><td>辅助邮箱/密码</td><td>可选</td><td>仅作为账号资料保存。</td></tr>\n        <tr><td>分类/项目</td><td>可选</td><td>用于分组、API 按项目取码。</td></tr>\n      </table>\n      <p class="muted">兼容 <code>----</code>、<code>|</code>、逗号、Tab 分隔。旧格式 <code>邮箱----Client ID----Refresh Token</code> 也支持。</p>\n    </div>\n    <div class="card">\n      <h3>二、令牌类型说明</h3>\n      <table><tr><th>类型</th><th>用途</th><th>适合场景</th></tr>\n        <tr><td>Graph 令牌</td><td>通过 Microsoft Graph <code>Mail.Read</code> 读取邮件。</td><td>最推荐，稳定、速度快、可读收件箱和垃圾箱。</td></tr>\n        <tr><td>OAuth IMAP 令牌</td><td>Access Token 通过 IMAP XOAUTH2 登录。</td><td>Graph 不可读时的兼容方案。</td></tr>\n        <tr><td>密码 IMAP</td><td>保存邮箱密码后用 IMAP 直接读取。</td><td>令牌失效时兜底判断账号是否还能收信。</td></tr>\n        <tr><td>无令牌账号</td><td>只保存邮箱/密码/分类。</td><td>不能刷新 token；如有密码可尝试 IMAP 取码。</td></tr>\n      </table>\n      <div class="notice info">系统读取验证码顺序：Graph → OAuth IMAP → 密码 IMAP。综合检测会告诉你哪个通道可用。</div>\n    </div>\n  </div>\n\n  <div class="card"><h3>三、工具能力总览（仓库里已有的工具能力）</h3><p class="muted">下面是后台已经实现并可直接使用的令牌、邮件和批量工具。</p><h3>刷新令牌 / 获取 Access Token</h3>\n    <table><tr><th>功能</th><th>入口</th><th>说明</th></tr>\n      <tr><td>单个刷新</td><td><code>/tokens</code> / <code>/token_tool</code></td><td>输入 Client ID + Refresh Token，换取 Access Token；如 Microsoft 返回新 Refresh Token，会自动轮换保存。</td></tr>\n      <tr><td>批量获取令牌</td><td>邮箱管理 → 批量操作 → 获取令牌</td><td>按勾选邮箱后台执行，显示进度条和当前处理邮箱。</td></tr>\n      <tr><td>批量刷新令牌</td><td>邮箱管理/令牌管理</td><td>后台任务异步执行，避免大批量请求 502。</td></tr>\n      <tr><td>状态检测</td><td>检测状态 / 批量检测</td><td>验证 Refresh Token 是否可换 Access Token，记录 <code>token_ok</code> 或 <code>token_failed</code>。</td></tr>\n    </table>\n    <p class="muted">Token 结果默认隐藏，只提供“显示/复制”按钮，降低误泄露风险。</p>\n  </div>\n\n  <div class="card"><h3>四、获取邮件和验证码</h3>\n    <p>入口：<code>/mails</code> 或外部 API <code>/api/v1/latest-code</code>。</p>\n    <table><tr><th>步骤</th><th>说明</th></tr>\n      <tr><td>1. 选择邮箱</td><td>在邮箱管理里选择当前读取邮箱，或 API 指定 <code>email/category</code>。</td></tr>\n      <tr><td>2. Graph 读取</td><td>用 Refresh Token 换 Access Token 后，通过 Graph 读取收件箱/垃圾邮箱。</td></tr>\n      <tr><td>3. IMAP 兜底</td><td>Graph 失败后尝试 OAuth IMAP；如令牌失败且保存了密码，再尝试密码 IMAP。</td></tr>\n      <tr><td>4. 提取验证码</td><td>从最新邮件主题/摘要中识别验证码，并返回邮件来源、时间、文件夹。</td></tr>\n    </table>\n    <div class="notice info">取不到验证码不一定代表账号挂了：可能是最近没有验证码邮件、规则不匹配、邮件延迟，或 Microsoft 风控。</div>\n  </div>\n\n  <div class="card"><h3>五、项目/分类管理</h3>\n    <p>入口：<code>/project-manage</code>、<code>/categories</code>。</p>\n    <table><tr><th>能力</th><th>用途</th></tr>\n      <tr><td>分类</td><td>给邮箱分组，批量导入时最后一列可直接写分类。</td></tr>\n      <tr><td>项目规则</td><td>按项目配置邮箱组、关键词过滤和返回数量，便于外部系统按项目取码。</td></tr>\n      <tr><td>API 按项目查询</td><td><code>/api/v1/latest-code?category=项目名</code> 可从该项目邮箱中读取验证码。</td></tr>\n      <tr><td>批量改分类</td><td>勾选邮箱后可批量设置分类。</td></tr>\n    </table>\n  </div>\n\n  <div class="card"><h3>六、外部 API 使用</h3>\n    <p class="muted">推荐统一使用请求头传 API Key，不建议 URL 明文传 key。</p>\n    <pre>curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/health"\ncurl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/projects"\ncurl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/accounts?category=项目名"\ncurl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?email=xxx@outlook.com"\ncurl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?category=项目名"\ncurl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/account-status?email=xxx@outlook.com"</pre>\n    <table><tr><th>接口</th><th>scope</th><th>说明</th></tr>\n      <tr><td><code>/health</code></td><td><code>health</code></td><td>健康检查。</td></tr>\n      <tr><td><code>/projects</code></td><td><code>projects</code></td><td>返回项目/分类列表。</td></tr>\n      <tr><td><code>/accounts</code></td><td><code>accounts</code></td><td>返回账号元数据和状态，不返回密码/token。</td></tr>\n      <tr><td><code>/latest-code</code></td><td><code>latest_code</code></td><td>读取最新验证码。</td></tr>\n      <tr><td><code>/account-status</code></td><td><code>accounts + latest_code</code></td><td>返回综合状态摘要。</td></tr>\n    </table>\n    <p class="muted">业务失败会返回 <code>ok:false</code> 和错误信息；服务本身正常时尽量不把账号失败伪装成 HTTP 502。</p>\n  </div>\n\n  <div class="card"><h3>七、免导入接口说明</h3>\n    <div class="notice bad">当前版本默认不提供“外部直接传邮箱密码/Refresh Token 的免导入取码接口”。原因是这类接口会让敏感凭证经过 URL、日志或第三方系统，安全风险更高。</div>\n    <p class="muted">推荐做法：先在后台导入并加密保存账号，再通过 API Key + email/category 调用。这样外部系统不需要持有邮箱密码或 Refresh Token。</p>\n  </div>\n\n  <div class="card"><h3>八、状态说明</h3>\n    <table><tr><th>状态</th><th>显示</th><th>含义</th></tr>\n      <tr><td><code>token_ok</code></td><td class="ok">令牌可用</td><td>Refresh Token 能换 Access Token。</td></tr>\n      <tr><td><code>token_failed</code></td><td class="bad">令牌失效</td><td>Refresh Token 失效、scope 未授权或授权过期。</td></tr>\n      <tr><td><code>graph_ok</code></td><td class="ok">Graph 可读</td><td>Graph Mail.Read 能读取邮件。</td></tr>\n      <tr><td><code>graph_failed</code></td><td class="bad">Graph 失败</td><td>Graph 请求失败或权限不足。</td></tr>\n      <tr><td><code>xoauth2_imap_ok</code></td><td class="ok">OAuth IMAP 可读</td><td>Access Token 可用于 IMAP XOAUTH2。</td></tr>\n      <tr><td><code>imap_password_ok</code></td><td class="ok">密码 IMAP 可读</td><td>保存的邮箱密码可以通过 IMAP 读取邮件。</td></tr>\n      <tr><td><code>all_failed</code></td><td class="bad">全部失败</td><td>令牌、Graph、IMAP 通道均不可用。</td></tr>\n    </table>\n  </div>\n\n  <div class="card"><h3>九、常见问题</h3>\n    <table><tr><th>问题</th><th>原因</th><th>处理建议</th></tr>\n      <tr><td><code>invalid_grant</code></td><td>Refresh Token 被撤销、过期或账号安全状态变化。</td><td>重新 OAuth 授权获取新 Refresh Token。</td></tr>\n      <tr><td><code>AADSTS70000</code></td><td>账号进入 service abuse mode，通常是微软风控/滥用限制。</td><td>暂停重试，网页登录解锁或验证；解锁后重新授权，无法解锁则标记不可用。</td></tr>\n      <tr><td>Graph 失败但密码 IMAP 可读</td><td>Graph 权限/令牌异常，账号本身可能还能收信。</td><td>继续用密码 IMAP 兜底，同时重新生成令牌。</td></tr>\n      <tr><td>OAuth IMAP 提示 authenticated but not connected</td><td>账号未初始化 Exchange mailbox 或邮箱服务不可用。</td><td>登录 Outlook 网页版初始化，或改用 Graph/密码 IMAP。</td></tr>\n      <tr><td>API Key 无效</td><td>Key 错误、被禁用或 scope 不足。</td><td>到 API 密钥页检查启用状态、scope 和限流。</td></tr>\n      <tr><td>批量任务很慢</td><td>Microsoft 网络、邮箱数量、IMAP 超时都会影响速度。</td><td>观察进度条；系统已限制并发并异步执行，避免 502。</td></tr>\n    </table>\n  </div>\n\n  <div class="card"><h3>十、安全和备份</h3>\n    <ul>\n      <li>数据库位置建议放在 <code>/www/server/rtweb/app.db</code>，不要放源码目录。</li>\n      <li><code>saved_accounts.refresh_token</code>、<code>password</code>、<code>aux_password</code> 会加密存储。</li>\n      <li>请备份 <code>RTWEB_DATA_KEY</code>，丢失后加密字段无法解密。</li>\n      <li>API Key 只保存 digest，不保存明文；新建后请立即复制。</li>\n      <li>自动更新默认关闭，建议先备份数据库和 env 后再开启。</li>\n      <li>不要把 <code>.env</code>、<code>*.db</code>、真实 token、账号密码提交到仓库。</li>\n    </ul>\n  </div>\n</section>'
    return app_page('使用指南', 'help', content)

def render_home_page(active_account_id: str = ''):
    summary = dashboard_summary()
    active_email = ''
    if active_account_id:
        active = get_saved_account(active_account_id)
        if active:
            active_email = active['email']
    recent_rows = ''.join(
        '<tr><td>' + html.escape(r['email']) + '</td><td>' + html.escape(r['category'] or '未分类') + '</td><td>' + html.escape(r['last_status'] or '-') + '</td><td>' + html.escape((r['last_error'] or '')[:80]) + '</td><td>' + html.escape(datetime.fromtimestamp(r['updated_at']).strftime('%m-%d %H:%M') if r['updated_at'] else '-') + '</td></tr>'
        for r in summary['recent']
    ) or '<tr><td colspan="5" class="muted">暂无邮箱。</td></tr>'
    content = f'''
<section class="section-stack">
  <div class="toolbox-hero">
    <div class="toolbox-kicker">🧰 Self-hosted 一点微软工具箱</div>
    <h1 class="toolbox-title">控制台总览</h1>
    <p class="toolbox-desc">一点 Outlook / Hotmail 工具箱：账号托管、令牌刷新、验证码读取、项目规则、API Key、查询接口和调用统计都集中在这里。</p>
    <div class="toolbox-stats">
      <span class="stat-pill">邮箱总数：<b>{summary['total_accounts']}</b></span>
      <span class="stat-pill">正常：<b>{summary['ok_accounts']}</b></span>
      <span class="stat-pill">异常：<b>{summary['error_accounts']}</b></span>
      <span class="stat-pill">项目/分类：<b>{summary['categories']}</b></span>
      <span class="stat-pill">启用 API Key：<b>{summary['api_keys']}</b></span>
      <span class="stat-pill">今日 API：<b>{summary['api_calls_today']}</b> / 失败 {summary['api_fail_today']}</span>
      <span class="stat-pill">当前读取邮箱：<b>{html.escape(active_email or '未选择')}</b></span>
    </div>
  </div>
  <div class="tool-grid">
    <a class="tool-card" href="/tokens"><div class="tool-icon">🔄</div><h3>刷新令牌</h3><p>获取 Access Token、检测 Refresh Token 状态、批量刷新令牌。</p></a>
    <a class="tool-card" href="/mailboxes"><div class="tool-icon">📮</div><h3>邮箱管理</h3><p>单个导入、批量导入、分类、导出、选择账号读取验证码。</p></a>
    <a class="tool-card" href="/mails"><div class="tool-icon">📥</div><h3>获取邮件</h3><p>通过 Microsoft Graph 读取 Outlook / Hotmail 最新验证码邮件。</p></a>
    <a class="tool-card" href="/api-key"><div class="tool-icon">🔐</div><h3>API 密钥</h3><p>生成外部接口密钥，查看调用统计、失败日志和限流情况。</p></a>
    <a class="tool-card" href="/api-manage"><div class="tool-icon">🔌</div><h3>查询 API</h3><p>提供 latest-code、account-status、accounts、health 等私有接口。</p></a>
    <a class="tool-card" href="/project-manage"><div class="tool-icon">📁</div><h3>项目管理</h3><p>按项目配置邮箱组、关键词过滤规则和返回数量。</p></a>
    <a class="tool-card" href="/categories"><div class="tool-icon">🏷️</div><h3>分类管理</h3><p>维护邮箱分类，支持批量改分类。</p></a>
    <a class="tool-card" href="/help"><div class="tool-icon">📘</div><h3>使用指南</h3><p>查看导入格式、状态说明和 API 调用方式。</p></a>
    <a class="tool-card" href="/version"><div class="tool-icon">⬆️</div><h3>版本/更新</h3><p>查看 GitHub 仓库地址、检测远程更新，并在启用后执行手动更新。</p></a>
  </div>
  <div class="quick-actions">
    <div class="card scroll"><h3>最近更新邮箱</h3><table><tr><th>邮箱</th><th>项目/分类</th><th>状态</th><th>错误</th><th>更新时间</th></tr>{recent_rows}</table></div>
    <div class="card"><h3>快捷 API</h3><pre>curl -H "X-API-Key: 你的APIKEY" https://token.seoyh.net/api/v1/health
curl -H "X-API-Key: 你的APIKEY" https://token.seoyh.net/api/v1/projects
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/accounts?category=项目名"
curl -H "X-API-Key: 你的APIKEY" "https://token.seoyh.net/api/v1/latest-code?category=项目名"</pre><div class="danger-note">API Key 页面可查看调用次数与限流状态。接口不会返回 Refresh Token / Access Token / 密码明文。</div></div>
  </div>
</section>'''
    return app_page('一点微软工具箱', 'dashboard', content)



def render_secret_textarea(label: str, value: str, min_height: int = 110) -> str:
    if not value:
        return ''
    textarea_id = 'secret-' + secrets.token_urlsafe(10).replace('-', '')
    safe_label = html.escape(label)
    safe_value = html.escape(value)
    return (
        '<div class="secret-block"><div class="action-row" style="justify-content:space-between;margin:8px 0 4px">'
        '<h3 style="margin:0">' + safe_label + '</h3>'
        '<div class="action-row" style="gap:6px">'
        '<button type="button" class="mini-btn" onclick="toggleSecret(\'' + textarea_id + '\', this)">显示</button>'
        '<button type="button" class="mini-btn primary" onclick="copyToClipboard(\'' + textarea_id + '\')">复制</button>'
        '</div></div>'
        '<textarea id="' + textarea_id + '" readonly data-secret-hidden="1" style="min-height:' + str(int(min_height)) + 'px;display:none">' + safe_value + '</textarea>'
        '<p class="muted">为降低泄露风险，默认隐藏；需要时点击“显示”或直接复制。</p></div>'
    )

class Handler(BaseHTTPRequestHandler):
    server_version = 'RTWeb/1.0'

    def log_message(self, fmt, *args):
        safe = fmt % args
        if 'GET /api/batch_job?' in safe:
            return
        safe = re.sub(r'([?&](?:key|refresh_token|access_token|csrf_token)=)[^&\s]+', r'\1***', safe, flags=re.I)
        print(f'{self.address_string()} - {safe}')

    def read_form_raw(self):
        if hasattr(self, '_posted_raw'):
            return self._posted_raw
        length = int(self.headers.get('Content-Length', '0') or '0')
        raw = self.rfile.read(length).decode('utf-8')
        self._posted_raw = urllib.parse.parse_qs(raw)
        return self._posted_raw

    def read_form(self):
        raw = self.read_form_raw()
        return {k: v[0] if v else '' for k, v in raw.items()}

    def send_security_headers(self, content_type: str = 'html'):
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
        if content_type == 'html':
            self.send_header('Content-Security-Policy', "default-src 'self'; img-src 'self' data: https:; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'")

    def send_html(self, body, status=200, headers=None):
        data = body.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_security_headers('html')
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        if self.require_auth() and not verify_csrf_cookie_value(parse_cookies(self.headers.get('Cookie', '')).get('rtweb_csrf', '')):
            self.send_header('Set-Cookie', make_csrf_cookie())
        self.end_headers()
        self.wfile.write(data)

    def send_text_file(self, text, filename='accounts.txt'):
        data = text.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_security_headers('text')
        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, path, headers=None):
        self.send_response(302)
        self.send_header('Location', path)
        self.send_security_headers('redirect')
        if headers:
            for k, v in headers.items():
                if k.lower().startswith('set-cookie'):
                    self.send_header('Set-Cookie', v)
                else:
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
            self.send_html(render_login_page())
            return
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ('/', '') and not self.require_auth():
            self.send_html(render_public_home_page())
            return
        if parsed.path.startswith('/api/v1/'):
            self.handle_public_api(parsed)
            return
        if not self.require_auth():
            self.redirect('/login'); return
        if parsed.path == '/api/batch_job':
            job_id = urllib.parse.parse_qs(parsed.query).get('id', [''])[0]
            job = get_batch_job(job_id)
            if not job:
                self.send_json({'error': 'job_not_found'}, 404)
                return
            self.send_json(job)
            return
        if parsed.path == '/export.txt':
            self.send_text_file(export_accounts_text(), 'outlook-accounts.txt')
            return
        if parsed.path == '/api/codes':
            self.send_codes_json()
            return
        if parsed.path == '/mailboxes':
            qs = urllib.parse.parse_qs(parsed.query)
            category_filter = qs.get('category', [''])[0]
            try:
                page_num = int(qs.get('page', ['1'])[0] or 1)
            except Exception:
                page_num = 1
            try:
                per_page = int(qs.get('per_page', ['50'])[0] or 50)
            except Exception:
                per_page = 50
            status_filter = qs.get('status', [''])[0]
            error_type = qs.get('error_type', [''])[0]
            q = qs.get('q', [''])[0].strip()
            self.send_html(render_mailboxes_page(get_active_account_id(self.headers.get('Cookie', '')), category_filter, page_num, per_page, status_filter, error_type, q))
            return
        if parsed.path == '/tokens':
            self.send_html(render_tokens_page())
            return
        if parsed.path == '/categories':
            self.send_html(render_categories_page())
            return
        if parsed.path == '/api-key':
            self.send_html(render_api_key_page())
            return
        if parsed.path == '/api-manage':
            self.send_html(render_api_manage_page())
            return
        if parsed.path == '/project-manage':
            self.send_html(render_project_manage_page())
            return
        if parsed.path == '/help':
            self.send_html(render_help_page())
            return
        if parsed.path == '/version':
            self.send_html(render_version_page())
            return
        if parsed.path == '/mails':
            self.send_html(render_mails_page(get_active_account_id(self.headers.get('Cookie', ''))))
            return
        if parsed.path == '/select':
            qs = urllib.parse.parse_qs(parsed.query)
            account_id = qs.get('id', [''])[0]
            next_path = qs.get('next', ['/mailboxes'])[0]
            if next_path not in ('/', '/mailboxes', '/tokens', '/categories', '/mails', '/api-key', '/api-manage', '/project-manage', '/help'):
                next_path = '/mailboxes'
            self.redirect(next_path, {'Set-Cookie': make_active_account_cookie(account_id)})
            return
        if parsed.path == '/token':
            account_id = urllib.parse.parse_qs(parsed.query).get('id', [''])[0]
            self.show_saved_token(account_id)
            return
        if parsed.path == '/read':
            account_id = urllib.parse.parse_qs(parsed.query).get('id', [''])[0]
            self.read_saved_account(account_id)
            return
        if parsed.path in ('/', ''):
            self.send_html(render_home_page(get_active_account_id(self.headers.get('Cookie', ''))))
            return
        self.send_error(404)


    def handle_public_api(self, parsed):
        qs = urllib.parse.parse_qs(parsed.query)
        header_key = self.headers.get('X-API-Key', '')
        query_key = qs.get('key', [''])[0]
        if query_key and not header_key and not API_ALLOW_QUERY_KEY:
            self.send_json({'ok': False, 'error': 'query_api_key_disabled', 'message': 'Use X-API-Key header instead.'}, 401)
            return
        key_value = header_key or query_key
        key_transport = 'header' if header_key else ('query_deprecated' if query_key else 'missing')
        key_row = verify_api_key_value(key_value)
        if not key_row:
            self.send_json({'ok': False, 'error': 'invalid_api_key'}, 401)
            return
        allowed, limit_error, minute_count, day_count = api_rate_limit_status(key_row['id'])
        if not allowed:
            log_api_call(key_row['id'], parsed.path, limit_error, '')
            self.send_json({'ok': False, 'error': limit_error, 'limit_per_minute': API_RATE_LIMIT_PER_MINUTE, 'limit_per_day': API_RATE_LIMIT_PER_DAY, 'minute_count': minute_count, 'day_count': day_count}, 429)
            return
        required_scopes = {
            '/api/v1/health': ['health'],
            '/api/v1/accounts': ['accounts'],
            '/api/v1/projects': ['projects'],
            '/api/v1/latest-code': ['latest_code'],
            '/api/v1/account-status': ['accounts', 'latest_code'],
        }.get(parsed.path, [])
        missing_scopes = [scope for scope in required_scopes if not api_key_has_scope(key_row, scope)]
        if missing_scopes:
            required_scope = ','.join(required_scopes)
            log_api_call(key_row['id'], parsed.path, 'forbidden_scope', required_scope)
            self.send_json({'ok': False, 'error': 'forbidden_scope', 'required_scope': required_scope, 'missing_scopes': missing_scopes, 'api_key_transport': key_transport}, 403)
            return
        if parsed.path == '/api/v1/health':
            log_api_call(key_row['id'], parsed.path, 'ok', '')
            self.send_json({'ok': True, 'service': 'rtweb', 'ts': int(time.time()), 'api_key_transport': key_transport, 'scopes': key_row['scopes'] or '', 'limits': {'per_minute': API_RATE_LIMIT_PER_MINUTE, 'per_day': API_RATE_LIMIT_PER_DAY}})
            return
        if parsed.path == '/api/v1/accounts':
            category = qs.get('category', [''])[0]
            try:
                limit = int(qs.get('limit', ['50'])[0] or 50)
            except Exception:
                limit = 50
            rows = public_account_rows(category, limit)
            accounts = [{'id': r['id'], 'email': r['email'], 'category': r['category'] or '未分类', 'status': r['last_status'] or '', 'error': (r['last_error'] or '')[:200], 'updated_at': r['updated_at']} for r in rows]
            log_api_call(key_row['id'], parsed.path, 'ok', category or 'all')
            self.send_json({'ok': True, 'category': category, 'count': len(accounts), 'accounts': accounts, 'api_key_transport': key_transport, 'ts': int(time.time())})
            return
        if parsed.path == '/api/v1/projects':
            rows = public_project_rows()
            projects = []
            for r in rows:
                projects.append({
                    'category': r['category'],
                    'account_count': r['account_count'],
                    'rule': {
                        'enabled': bool(r['enabled']),
                        'sender_keywords': r['sender_keywords'] or '',
                        'subject_keywords': r['subject_keywords'] or '',
                        'body_keywords': r['body_keywords'] or '',
                        'max_results': r['max_results'] or 5,
                    }
                })
            log_api_call(key_row['id'], parsed.path, 'ok', '')
            self.send_json({'ok': True, 'count': len(projects), 'projects': projects, 'api_key_transport': key_transport, 'ts': int(time.time())})
            return
        if parsed.path == '/api/v1/account-status':
            email_addr = qs.get('email', [''])[0]
            category = qs.get('category', [''])[0]
            account = get_saved_account_by_email(email_addr) if email_addr else newest_account_for_category(category)
            if not account:
                log_api_call(key_row['id'], parsed.path, 'not_found', email_addr or category)
                self.send_json({'ok': False, 'error': 'account_not_found', 'api_key_transport': key_transport}, 404)
                return
            result = inspect_saved_account(account, limit=10)
            ok_any = bool(result['token']['ok'] or result['graph_mail']['ok'] or result['password_login']['ok'])
            source_status = 'graph_ok' if result['graph_mail']['ok'] else ('imap_password_ok' if result['password_login']['ok'] else ('token_ok' if result['token']['ok'] else 'all_failed'))
            payload = {
                'ok': ok_any,
                'email': account['email'],
                'category': account.get('category') or '未分类',
                'account_status': status_label(source_status),
                'best_source': result.get('best_source') or '',
                'code_count': len(result.get('best_codes') or []),
                'codes': result.get('best_codes') or [],
                'token': {
                    'configured': bool(result['token']['configured']),
                    'ok': bool(result['token']['ok']),
                    'status': result['token'].get('status') or '',
                    'expires_in': result['token'].get('expires_in') or '',
                    'rotated': bool(result['token'].get('rotated')),
                    'error': (result['token'].get('error') or '')[:300],
                },
                'graph_mail': {
                    'ok': bool(result['graph_mail']['ok']),
                    'mail_count': result['graph_mail'].get('mail_count') or 0,
                    'error': (result['graph_mail'].get('error') or '')[:300],
                },
                'password_imap': {
                    'configured': bool(result['password_login']['configured']),
                    'ok': bool(result['password_login']['ok']),
                    'mail_count': result['password_login'].get('mail_count') or 0,
                    'error': (result['password_login'].get('error') or '')[:300],
                },
                'api_key_transport': key_transport,
                'ts': int(time.time()),
            }
            log_api_call(key_row['id'], parsed.path, 'ok' if ok_any else 'error', account['email'])
            self.send_json(payload, 200)
            return
        if parsed.path == '/api/v1/latest-code':
            email_addr = qs.get('email', [''])[0]
            category = qs.get('category', [''])[0]
            try:
                limit = max(1, min(50, int(qs.get('limit', ['0'])[0] or 0)))
            except Exception:
                limit = 0
            account = get_saved_account_by_email(email_addr) if email_addr else newest_account_for_category(category)
            rule = get_project_rule(category) if category else None
            if qs.get('sender', [''])[0] or qs.get('subject', [''])[0] or qs.get('body', [''])[0]:
                rule = {'sender_keywords': qs.get('sender', [''])[0], 'subject_keywords': qs.get('subject', [''])[0], 'body_keywords': qs.get('body', [''])[0], 'max_results': limit or 5, 'enabled': 1}
            if rule and not limit:
                limit = int(rule['max_results'] or 5)
            if not limit:
                limit = 10
            if not account:
                log_api_call(key_row['id'], parsed.path, 'not_found', email_addr or category)
                self.send_json({'ok': False, 'error': 'account_not_found'}, 404)
                return
            result = fetch_filtered_codes_for_account(account, rule, limit=limit)
            codes = result.get('codes') or []
            log_api_call(key_row['id'], parsed.path, result.get('status', 'ok'), account['email'])
            self.send_json({'ok': result.get('status') == 'ok', 'email': account['email'], 'category': category, 'source': result.get('source_api'), 'channel_status': result.get('channel_status', {}), 'account_status': status_label('graph_ok' if result.get('source_api') == 'graph' else ('imap_password_ok' if result.get('source_api') == 'imap_password' else 'xoauth2_imap_ok')) if result.get('status') == 'ok' else status_label('all_failed'), 'rule_applied': bool(rule), 'count': len(codes), 'codes': codes, 'error': result.get('error', ''), 'api_key_transport': key_transport, 'ts': int(time.time())}, 200)
            return
        log_api_call(key_row['id'], parsed.path, 'not_found', '')
        self.send_json({'ok': False, 'error': 'unknown_endpoint'}, 404)

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_security_headers('json')
        self.end_headers()
        self.wfile.write(data)

    def send_codes_json(self):
        active_id = get_active_account_id(self.headers.get('Cookie', ''))
        results = []
        if active_id:
            account = get_saved_account(active_id)
            if account:
                ok, value = timed_fetch_latest_codes_for_account(account, 10)
                if ok:
                    results = [value]
                else:
                    results = [{'email': account['email'], 'status': 'error', 'error': str(value)[:1000], 'source_api': '', 'channel_status': {'timeout': 'failed'}, 'codes': []}]
        self.send_json({'results': results, 'active_id': active_id, 'ts': int(time.time())})

    def render_check_result(self, result):
        cls = 'ok' if result.get('ok') else 'bad'
        rotated = '是，已自动保存新 Refresh Token' if result.get('rotated') else '否'
        body = f'''<div class="card"><h2>令牌状态检测</h2>
<table>
<tr><th>邮箱</th><td>{html.escape(result.get('email',''))}</td></tr>
<tr><th>状态</th><td class="{cls}">{html.escape(result.get('status',''))}</td></tr>
<tr><th>Access Token 有效期</th><td>{html.escape(str(result.get('expires_in') or ''))}</td></tr>
<tr><th>Refresh Token 是否轮换</th><td>{html.escape(rotated)}</td></tr>
<tr><th>Scope</th><td>{html.escape(result.get('scope') or '')}</td></tr>
<tr><th>Scope 策略</th><td>{html.escape(result.get('scope_mode') or '')}</td></tr>
<tr><th>错误</th><td class="bad">{html.escape(result.get('error') or '')}</td></tr>
</table><p><a href="/">返回</a></p></div>'''
        return page('令牌状态检测', body)

    def show_check_result(self, account_id):
        account = get_saved_account(account_id)
        if not account:
            self.send_html(page('不存在', '<div class="card"><p class="bad">账号不存在。</p><a href="/">返回</a></div>'), 404)
            return
        result = check_saved_account_token(account)
        self.send_html(self.render_check_result(result), 200 if result.get('ok') else 400)

    @staticmethod
    def render_batch_check_result_static(results):
        rows = []
        for r in results:
            cls = 'ok' if r.get('ok') else 'bad'
            rotated = '是' if r.get('rotated') else ''
            rows.append('<tr><td>' + html.escape(r.get('email','')) + '</td><td class="' + cls + '">' + html.escape(r.get('status','')) + '</td><td>' + html.escape(str(r.get('expires_in') or '')) + '</td><td>' + html.escape(rotated) + '</td><td>' + html.escape((r.get('error') or '')[:200]) + '</td></tr>')
        table = '<table><tr><th>邮箱</th><th>状态</th><th>Access Token 有效期</th><th>Refresh Token 轮换</th><th>错误</th></tr>' + ''.join(rows) + '</table>'
        return page('批量检测结果', '<div class="card"><h2>批量检测结果</h2>' + table + '<p><a href="/tokens">返回令牌管理</a></p></div>')

    def render_batch_check_result(self, results):
        return self.render_batch_check_result_static(results)

    @staticmethod
    def render_batch_refresh_result_static(rows, title='批量刷新结果'):
        result_rows = []
        for account, ok, payload in rows:
            if ok:
                result_rows.append('<tr><td>' + html.escape(account['email']) + '</td><td class="ok">成功</td><td>' + html.escape(str(payload.get('expires_in', ''))) + '</td></tr>')
            else:
                err = (payload.get('error_description') or payload.get('error') or '刷新失败')[:200]
                result_rows.append('<tr><td>' + html.escape(account['email']) + '</td><td class="bad">失败</td><td>' + html.escape(err) + '</td></tr>')
        table = '<table><tr><th>邮箱</th><th>结果</th><th>信息</th></tr>' + ''.join(result_rows) + '</table>'
        return page(title, '<div class="card"><h2>' + html.escape(title) + '</h2>' + table + '</div>')

    def render_batch_refresh_result(self, rows, title='批量刷新结果'):
        return self.render_batch_refresh_result_static(rows, title)

    @staticmethod
    def render_batch_token_result_static(rows):
        parts = ['<div class="card"><h2>批量获取令牌结果</h2><p class="muted">以下结果仅包含已勾选邮箱。Token 默认隐藏，可按需显示或复制。</p>']
        for account, ok, payload, refresh_token in rows:
            parts.append('<div class="card" style="margin:10px 0"><h3>' + html.escape(account['email']) + '</h3>')
            if ok:
                access_token = payload.get('access_token', '')
                new_refresh = payload.get('refresh_token') or refresh_token
                parts.append('<p class="ok">成功</p>')
                if access_token:
                    parts.append(render_secret_textarea('Access Token', access_token, 100))
                if new_refresh:
                    parts.append(render_secret_textarea('Refresh Token', new_refresh, 100))
            else:
                err = payload.get('error_description') or payload.get('error') or '获取失败'
                parts.append('<p class="bad">失败：' + html.escape(err[:500]) + '</p>')
            parts.append('</div>')
        parts.append('</div>')
        return page('批量获取令牌结果', ''.join(parts))

    def render_batch_token_result(self, rows):
        return self.render_batch_token_result_static(rows)

    def render_token_result(self, account_email, ok, payload, refresh_token):
        shown = {k: v for k, v in payload.items() if k not in ('id_token',)}
        access_token = shown.get('access_token', '')
        new_refresh = shown.get('refresh_token', refresh_token)
        token_html = ''
        if access_token:
            token_html += render_secret_textarea('Access Token', access_token, 130)
        if new_refresh:
            token_html += render_secret_textarea('Refresh Token', new_refresh, 130)
        shown_safe = {k: v for k, v in shown.items() if k not in ('access_token', 'refresh_token')}
        status = '<p class="ok">令牌获取/刷新成功。</p>' if ok else '<p class="bad">令牌获取/刷新失败。</p>'
        return page('令牌结果', '<div class="card"><h2>' + html.escape(account_email) + '</h2>' + status + token_html + '<h3>响应信息</h3><pre>' + html.escape(json.dumps(shown_safe or payload, ensure_ascii=False, indent=2)) + '</pre><p><a href="/">返回</a></p></div>')

    def show_saved_token(self, account_id):
        account = get_saved_account(account_id)
        if not account:
            self.send_html(page('不存在', '<div class="card"><p class="bad">账号不存在。</p><a href="/">返回</a></div>'), 404)
            return
        ok, payload, refresh_token = refresh_saved_account_token(account)
        self.send_html(self.render_token_result(account['email'], ok, payload, refresh_token), 200 if ok else 400)

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
        ok, payload = exchange_refresh_token_compatible(account['client_id'], account['refresh_token'], scope=default_oauth_scope(), tenant='consumers')
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
        post_path = urllib.parse.urlparse(self.path).path
        if post_path != '/' and post_path.endswith('/'):
            post_path = post_path.rstrip('/')
        if post_path == '/login':
            identity = self.client_address[0] if self.client_address else 'unknown'
            limited, hits = login_rate_limited(identity)
            if limited:
                self.send_html(render_login_page('登录失败次数过多，请稍后再试。'), 429)
                return
            form = self.read_form()
            user_ok = check_password(form.get('username', '').strip(), LOGIN_USERNAME)
            login_ok = check_password(form.get('login_password', ''), LOGIN_PASSWORD)
            admin_ok = check_password(form.get('admin_password', ''), ADMIN_PASSWORD)
            if user_ok and login_ok and admin_ok:
                clear_login_failures(identity)
                self.redirect('/', {'Set-Cookie': make_session_cookie(), 'Set-Cookie-2': make_csrf_cookie()})
            else:
                record_login_failure(identity)
                self.send_html(render_login_page('登录失败：用户账号、用户密码或应用管理密码不正确。'), 403)
            return
        if not self.require_auth():
            self.redirect('/login'); return
        raw_for_csrf = self.read_form_raw()
        if not verify_csrf_form(self.headers.get('Cookie', ''), raw_for_csrf):
            self.send_html(page('请求已拦截', '<div class="card"><h2>请求已拦截</h2><p class="bad">安全校验失败，请刷新页面后重试。</p><p><a href="/">返回控制台</a></p></div>'), 403)
            return
        if post_path == '/api_key_create':
            form = self.read_form()
            scopes = ','.join(raw_for_csrf.get('scopes', []))
            key = create_api_key(form.get('name', ''), form.get('note', ''), scopes)
            self.send_html(render_api_key_page(key))
            return
        if post_path == '/api_key_toggle':
            form = self.read_form()
            set_api_key_enabled(form.get('id', ''), form.get('enabled', '') == '1')
            self.redirect('/api-key')
            return
        if post_path == '/api_key_delete':
            form = self.read_form()
            delete_api_key(form.get('id', ''))
            self.redirect('/api-key')
            return
        if post_path == '/project_rule_save':
            form = self.read_form()
            upsert_project_rule(form.get('category', ''), form.get('sender_keywords', ''), form.get('subject_keywords', ''), form.get('body_keywords', ''), form.get('max_results', '5'), form.get('enabled', '') == '1')
            self.redirect('/project-manage')
            return
        if post_path == '/check':
            form = self.read_form()
            self.show_check_result(form.get('id', ''))
            return
        if post_path == '/check_all':
            rows = saved_accounts()
            job_id = create_batch_job('批量检测结果', 'check', rows, lambda account: check_saved_account_token(account))
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/refresh':
            form = self.read_form()
            self.show_saved_token(form.get('id', ''))
            return
        if post_path == '/refresh_all':
            rows = saved_accounts()
            def worker(account):
                ok, payload, _ = refresh_saved_account_token(account)
                return (account, ok, payload)
            job_id = create_batch_job('批量刷新结果', 'refresh', rows, worker)
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/check_selected':
            raw = self.read_form_raw()
            rows = selected_saved_accounts(raw.get('ids', []))
            job_id = create_batch_job('检查已选邮箱状态', 'check', rows, lambda account: check_saved_account_token(account))
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/inspect_selected':
            raw = self.read_form_raw()
            rows = selected_saved_accounts(raw.get('ids', []))
            job_id = create_batch_job('批量综合检测', 'inspect', rows, lambda account: inspect_saved_account(account, limit=10))
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/refresh_selected':
            raw = self.read_form_raw()
            rows = selected_saved_accounts(raw.get('ids', []))
            def worker(account):
                ok, payload, _ = refresh_saved_account_token(account)
                return (account, ok, payload)
            job_id = create_batch_job('刷新已选邮箱令牌结果', 'refresh', rows, worker)
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/token_selected':
            raw = self.read_form_raw()
            rows = selected_saved_accounts(raw.get('ids', []))
            def worker(account):
                ok, payload, refresh_token = refresh_saved_account_token(account)
                return (account, ok, payload, refresh_token)
            job_id = create_batch_job('获取已选邮箱令牌', 'token', rows, worker)
            
            if not job_id:
                self.send_json({'error': 'too_many_batch_jobs', 'max_running': BATCH_MAX_RUNNING}, 429)
                return
            self.send_json({'job_id': job_id})
            return
        if post_path == '/bulk_category':
            raw = self.read_form_raw()
            ids = raw.get('ids', [])
            category = raw.get('category', [''])[0] if raw.get('category') else ''
            updated = update_accounts_category(ids, category)
            label = normalize_category_name(category) or '未分类'
            self.send_html(page('批量设置分类', '<div class="card"><h2>批量设置分类</h2><p class="ok">已更新 ' + html.escape(str(updated)) + ' 个邮箱。</p><p>目标分类：<b>' + html.escape(label) + '</b></p></div>'))
            return
        if post_path == '/bulk_delete':
            raw = self.read_form_raw()
            ids = raw.get('ids', [])
            deleted = delete_saved_accounts(ids)
            self.send_html(page('批量删除邮箱', '<div class="card"><h2>批量删除邮箱</h2><p class="ok">已删除 ' + html.escape(str(deleted)) + ' 个邮箱。</p><p class="muted">如误删，可从操作前数据库备份恢复。</p><p><a href="/mailboxes">返回邮箱管理</a></p></div>'))
            return
        if post_path == '/category_create':
            form = self.read_form()
            create_category(form.get('name', ''))
            self.redirect('/categories')
            return
        if post_path == '/category_rename':
            form = self.read_form()
            rename_category(form.get('id', ''), form.get('name', ''))
            self.redirect('/categories')
            return
        if post_path == '/category_delete':
            form = self.read_form()
            delete_category(form.get('id', ''))
            self.redirect('/categories')
            return
        if post_path == '/category':
            form = self.read_form()
            update_account_category(form.get('id', ''), form.get('category', ''))
            self.redirect('/mailboxes')
            return
        if post_path == '/delete':
            form = self.read_form()
            delete_saved_account(form.get('id', ''))
            self.redirect('/mailboxes')
            return
        if post_path == '/batch':
            form = self.read_form()
            batch_category = form.get('category', '').strip()
            rows = parse_batch_accounts(form.get('batch', ''))
            for row in rows:
                save_account(row['email'], row['client_id'], row['refresh_token'], row.get('password', ''), row.get('aux_email', ''), row.get('aux_password', ''), row.get('category', '') or batch_category)
            self.redirect('/?saved=' + str(len(rows)))
            return
        if post_path == '/inspect_account':
            form = self.read_form()
            account = get_saved_account(form.get('id', '').strip())
            if not account:
                self.send_html(page('不存在', '<div class="card"><p class="bad">账号不存在。</p><a href="/mailboxes">返回</a></div>'), 404)
                return
            result = inspect_saved_account(account, limit=10)
            self.send_html(render_account_inspect_result(result), 200 if (result['token']['ok'] or result['graph_mail']['ok'] or result['password_login']['ok']) else 400)
            return
        if post_path == '/check-update':
            code, out = run_shell_command('git fetch origin ' + sh_quote(UPDATE_BRANCH) + ' --tags', timeout=30)
            info = version_info(False)
            latest_tag = info.get('latest_tag') or '未检测到 Release Tag'
            release_behind = str(info.get('release_behind') or '')
            release_status = '未检测' if release_behind == '' else ('已是最新 Release' if release_behind == '0' else '落后最新 Release ' + release_behind + ' 个提交')
            msg = ('检查完成。最新 Release：' + latest_tag + '；' + release_status) if code == 0 else ('检查失败：' + out[:500])
            self.send_html(render_version_page(msg))
            return
        if post_path == '/update':
            if not AUTO_UPDATE_ENABLED:
                self.send_html(render_version_page('手动更新已关闭。请移除 RTWEB_AUTO_UPDATE_ENABLED=0 或设置为 1 后再执行。'))
                return
            code, out = run_shell_command(UPDATE_COMMAND, timeout=180)
            msg = ('更新命令执行完成。' if code == 0 else '更新命令执行失败。') + '\n' + out[:1200]
            self.send_html(render_version_page(msg))
            return
        if post_path == '/token_tool':
            form = self.read_form()
            client_id = form.get('client_id', '').strip()
            refresh_token = form.get('refresh_token', '').strip()
            email_addr = form.get('email', '').strip()
            category = form.get('category', '').strip()
            scope = (form.get('scope', '').strip() or default_oauth_scope())
            tenant = form.get('tenant', 'consumers').strip() or 'consumers'
            ok, payload = exchange_refresh_token_compatible(client_id, refresh_token, scope=scope, tenant=tenant)
            if ok:
                new_refresh = payload.get('refresh_token') or refresh_token
                insert_record(client_id, new_refresh, 'ok', account=email_addr or None, scope=payload.get('scope'), expires_in=payload.get('expires_in'))
                if form.get('save_account') == '1' and email_addr:
                    save_account(email_addr, client_id, new_refresh, category=category)
                    update_saved_status(email_addr, 'ok')
                self.send_html(self.render_token_result(email_addr or '临时刷新', True, payload, new_refresh))
            else:
                err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
                insert_record(client_id, refresh_token, 'error', account=email_addr or None, error=err)
                if email_addr:
                    update_saved_status(email_addr, 'error', err)
                self.send_html(self.render_token_result(email_addr or '临时刷新', False, payload, refresh_token), 400)
            return
        if post_path == '/decode':
            form = self.read_form()
            client_id = form.get('client_id', '').strip()
            refresh_token = form.get('refresh_token', '').strip()
            email_addr = form.get('email', '').strip()
            scope = default_oauth_scope()
            tenant = 'consumers'
            save_account(email_addr, client_id, refresh_token, category=form.get('category', '').strip())
            ok, payload = exchange_refresh_token_compatible(client_id, refresh_token, scope=scope, tenant=tenant)
            if ok:
                new_refresh = payload.get('refresh_token') or refresh_token
                save_account(email_addr, client_id, new_refresh, category=form.get('category', '').strip())
                insert_record(client_id, new_refresh, 'ok', account=email_addr, scope=payload.get('scope'), expires_in=payload.get('expires_in'))
                update_saved_status(email_addr, 'ok')
                self.send_html(self.render_token_result(email_addr, True, payload, new_refresh))
            else:
                err = (payload.get('error', '') + ': ' + payload.get('error_description', ''))[:1000]
                insert_record(client_id, refresh_token, 'error', account=email_addr, error=err)
                update_saved_status(email_addr, 'error', err)
                self.send_html(self.render_token_result(email_addr, False, payload, refresh_token), 400)
            return
        self.send_error(404)


def main():
    init_db()
    httpd = ThreadingHTTPServer((BIND_HOST, BIND_PORT), Handler)
    print(f'RTWeb listening on http://{BIND_HOST}:{BIND_PORT}, db={DB_PATH}')
    httpd.serve_forever()

if __name__ == '__main__':
    main()
