"""
database.py — בסיס הנתונים של שבצ"ק (SQLite)

טבלאות:
  users        משתמשי המערכת (מפקדים). הסיסמה נשמרת כגיבוב (hash) בלבד.
  soldiers     החיילים של כל משתמש
  tasks        המשימות של כל משתמש
  runs         כל הרצת שיבוץ: מתי, של מי, ומה היו המדדים
  assignments  שורה לכל תקן בכל משמרת בכל הרצה (מי שובץ, או חוסר)
  events       יומן שימוש: כניסות, הרשמות, הרצות, הורדות

קשרים:  users 1─∞ soldiers, tasks, runs, events    |    runs 1─∞ assignments

מיקום הקובץ נקבע במשתנה הסביבה SHAVTZAK_DB (ברירת מחדל: shavtzak.db ליד הקוד).
"""
import datetime
import hashlib
import json
import os
import secrets
import sqlite3

import pandas as pd

DB_PATH = os.environ.get('SHAVTZAK_DB', 'shavtzak.db')
ADMIN_USERNAME = 'admin'

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT    NOT NULL UNIQUE,
    unit_name   TEXT    NOT NULL DEFAULT '',
    pw_hash     TEXT    NOT NULL,
    pw_salt     TEXT    NOT NULL,
    is_admin    INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT    NOT NULL
);
CREATE TABLE IF NOT EXISTS soldiers (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id       INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    personal_id   TEXT    NOT NULL,
    full_name     TEXT    NOT NULL,
    exemptions    TEXT    NOT NULL DEFAULT '',
    roles         TEXT    NOT NULL DEFAULT '',
    blocked_hours TEXT    NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS tasks (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id        INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    task_no        INTEGER NOT NULL,
    name           TEXT    NOT NULL,
    required       INTEGER NOT NULL,
    shift_hours    REAL    NOT NULL,
    rest_hours     REAL    NOT NULL DEFAULT 0,
    shared         INTEGER NOT NULL DEFAULT 0,
    active_hours   TEXT    NOT NULL DEFAULT '',
    required_roles TEXT    NOT NULL DEFAULT '',
    intensity      INTEGER NOT NULL DEFAULT 1,
    blocked_roles  TEXT    NOT NULL DEFAULT '',
    priority       INTEGER NOT NULL DEFAULT 5
);
CREATE TABLE IF NOT EXISTS runs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at      TEXT    NOT NULL,
    label           TEXT    NOT NULL,
    day_start       INTEGER NOT NULL,
    engine          TEXT    NOT NULL,
    num_soldiers    INTEGER NOT NULL,
    num_tasks       INTEGER NOT NULL,
    short_hours     REAL    NOT NULL,
    coverage_pct    REAL    NOT NULL,
    rest_violations INTEGER NOT NULL,
    full_rest       INTEGER NOT NULL,
    load_gap        REAL    NOT NULL,
    solve_seconds   REAL    NOT NULL,
    roster_json     TEXT    NOT NULL,
    compare_json    TEXT    NOT NULL
);
CREATE TABLE IF NOT EXISTS assignments (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id       INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    task_no      INTEGER NOT NULL,
    task_name    TEXT    NOT NULL,
    personal_id  TEXT,                 -- NULL = תקן לא מאויש (חוסר)
    soldier_name TEXT,
    role         TEXT,
    start_min    INTEGER NOT NULL,     -- דקות מתחילת יממת השיבוץ
    end_min      INTEGER NOT NULL,
    shared       INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS events (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id    INTEGER REFERENCES users(id) ON DELETE SET NULL,
    ts         TEXT    NOT NULL,
    event_type TEXT    NOT NULL,       -- register / login / run / download / demo / upload
    details    TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_soldiers_user ON soldiers(user_id);
CREATE INDEX IF NOT EXISTS idx_tasks_user    ON tasks(user_id);
CREATE INDEX IF NOT EXISTS idx_runs_user     ON runs(user_id);
CREATE INDEX IF NOT EXISTS idx_assign_run    ON assignments(run_id);
CREATE INDEX IF NOT EXISTS idx_events_ts     ON events(ts);
"""

EVENT_LABELS = {'register': 'הרשמה', 'login': 'כניסה', 'run': 'הרצת שיבוץ',
                'download': 'הורדת קובץ', 'demo': 'טעינת נתוני דוגמה', 'upload': 'העלאת נתונים'}


def _connect():
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys = ON')
    return con


def _now() -> str:
    try:
        from zoneinfo import ZoneInfo
        now = datetime.datetime.now(ZoneInfo('Asia/Jerusalem'))
    except Exception:
        now = datetime.datetime.now()
    return now.strftime('%Y-%m-%d %H:%M:%S')


def init_db():
    con = _connect()
    con.executescript(SCHEMA)
    con.commit()
    con.close()


# ── משתמשים ──────────────────────────────────────────────────────
def _hash(password: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'),
                               bytes.fromhex(salt_hex), 200_000).hex()


def _user_dict(row) -> dict:
    return {'id': row['id'], 'username': row['username'], 'unit_name': row['unit_name'],
            'is_admin': bool(row['is_admin'])}


def create_user(username: str, password: str, unit_name: str = ''):
    """מחזיר (user, None) בהצלחה, או (None, הודעת שגיאה)."""
    username = (username or '').strip()
    if len(username) < 3:
        return None, 'שם המשתמש חייב להכיל לפחות 3 תווים.'
    if username.lower() == ADMIN_USERNAME:
        return None, 'שם המשתמש הזה שמור.'
    if len(password or '') < 6:
        return None, 'הסיסמה חייבת להכיל לפחות 6 תווים.'
    salt = secrets.token_hex(16)
    con = _connect()
    try:
        cur = con.execute(
            'INSERT INTO users (username, unit_name, pw_hash, pw_salt, is_admin, created_at) '
            'VALUES (?,?,?,?,0,?)',
            (username, (unit_name or '').strip(), _hash(password, salt), salt, _now()))
        con.commit()
        uid = cur.lastrowid
    except sqlite3.IntegrityError:
        return None, 'שם המשתמש כבר תפוס.'
    finally:
        con.close()
    log_event(uid, 'register')
    return {'id': uid, 'username': username, 'unit_name': (unit_name or '').strip(),
            'is_admin': False}, None


def authenticate(username: str, password: str):
    con = _connect()
    row = con.execute('SELECT * FROM users WHERE username = ?',
                      ((username or '').strip(),)).fetchone()
    con.close()
    if row is None:
        return None
    if not secrets.compare_digest(_hash(password or '', row['pw_salt']), row['pw_hash']):
        return None
    log_event(row['id'], 'login')
    return _user_dict(row)


def ensure_admin(password: str):
    """יוצר או מעדכן את משתמש המנהל לפי הסיסמה שהוגדרה מחוץ לקוד."""
    if not password:
        return
    salt = secrets.token_hex(16)
    con = _connect()
    row = con.execute('SELECT id, pw_hash, pw_salt FROM users WHERE username = ?',
                      (ADMIN_USERNAME,)).fetchone()
    if row is None:
        con.execute(
            'INSERT INTO users (username, unit_name, pw_hash, pw_salt, is_admin, created_at) '
            'VALUES (?,?,?,?,1,?)',
            (ADMIN_USERNAME, 'מנהל המערכת', _hash(password, salt), salt, _now()))
    elif _hash(password, row['pw_salt']) != row['pw_hash']:
        con.execute('UPDATE users SET pw_hash = ?, pw_salt = ?, is_admin = 1 WHERE id = ?',
                    (_hash(password, salt), salt, row['id']))
    con.commit()
    con.close()


# ── יומן שימוש ───────────────────────────────────────────────────
def log_event(user_id, event_type: str, details: str = ''):
    con = _connect()
    con.execute('INSERT INTO events (user_id, ts, event_type, details) VALUES (?,?,?,?)',
                (user_id, _now(), event_type, details))
    con.commit()
    con.close()


# ── חיילים ומשימות ───────────────────────────────────────────────
def _txt(v) -> str:
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s


def _num(v, default=0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def save_roster(user_id: int, s_rows: list, t_rows: list):
    """מחליף את רשימות החיילים והמשימות של המשתמש."""
    con = _connect()
    con.execute('DELETE FROM soldiers WHERE user_id = ?', (user_id,))
    con.execute('DELETE FROM tasks WHERE user_id = ?', (user_id,))
    con.executemany(
        'INSERT INTO soldiers (user_id, personal_id, full_name, exemptions, roles, blocked_hours) '
        'VALUES (?,?,?,?,?,?)',
        [(user_id, _txt(r.get('personal_id')).removesuffix('.0'), _txt(r.get('full_name')),
          _txt(r.get('exemptions')), _txt(r.get('roles')), _txt(r.get('blocked_hours')))
         for r in s_rows])
    con.executemany(
        'INSERT INTO tasks (user_id, task_no, name, required, shift_hours, rest_hours, shared, '
        'active_hours, required_roles, intensity, blocked_roles, priority) '
        'VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
        [(user_id, int(_num(r.get('task_no'))), _txt(r.get('name')), int(_num(r.get('required'), 1)),
          _num(r.get('shift_hours'), 1), _num(r.get('rest_hours'), 0),
          1 if (r.get('shared') is True or _txt(r.get('shared')).lower() in ('true', '1', 'כן')) else 0,
          _txt(r.get('active_hours')), _txt(r.get('required_roles')),
          int(_num(r.get('intensity'), 1)), _txt(r.get('blocked_roles')),
          int(_num(r.get('priority'), 5) or 5))
         for r in t_rows])
    con.commit()
    con.close()


def load_roster(user_id: int):
    con = _connect()
    s_rows = [dict(r) for r in con.execute(
        'SELECT personal_id, full_name, exemptions, roles, blocked_hours '
        'FROM soldiers WHERE user_id = ? ORDER BY id', (user_id,))]
    t_rows = [dict(r) for r in con.execute(
        'SELECT task_no, name, required, shift_hours, rest_hours, shared, active_hours, '
        'required_roles, intensity, blocked_roles, priority '
        'FROM tasks WHERE user_id = ? ORDER BY id', (user_id,))]
    con.close()
    return s_rows, t_rows


# ── הרצות ושיבוצים ───────────────────────────────────────────────
def save_run(user_id: int, label: str, day_start: int, engine: str, num_tasks: int,
             metrics: dict, seconds: float, roster: list, compare: list, rows: list) -> int:
    con = _connect()
    cur = con.execute(
        'INSERT INTO runs (user_id, created_at, label, day_start, engine, num_soldiers, num_tasks, '
        'short_hours, coverage_pct, rest_violations, full_rest, load_gap, solve_seconds, '
        'roster_json, compare_json) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
        (user_id, _now(), label, day_start, engine, len(roster), num_tasks,
         round(metrics['short_minutes'] / 60, 2), round(metrics['coverage_pct'], 2),
         len(metrics['rest_violations']), metrics['full_rest'], round(metrics['load_gap'], 2),
         seconds, json.dumps(roster, ensure_ascii=False), json.dumps(compare, ensure_ascii=False)))
    run_id = cur.lastrowid
    con.executemany(
        'INSERT INTO assignments (run_id, task_no, task_name, personal_id, soldier_name, role, '
        'start_min, end_min, shared) VALUES (?,?,?,?,?,?,?,?,?)',
        [(run_id, r['task_no'], r['task_name'], r['personal_id'], r['soldier_name'], r['role'],
          r['start_min'], r['end_min'], r['shared']) for r in rows])
    con.commit()
    con.close()
    return run_id


def list_runs(user_id: int) -> pd.DataFrame:
    con = _connect()
    df = pd.read_sql_query(
        'SELECT id, created_at, label, num_soldiers, num_tasks, short_hours, coverage_pct, '
        'rest_violations, full_rest, load_gap, solve_seconds '
        'FROM runs WHERE user_id = ? ORDER BY id DESC', con, params=(user_id,))
    con.close()
    return df


def load_run(run_id: int, user_id: int):
    """מחזיר (פרטי ההרצה, שורות השיבוץ) — רק אם ההרצה שייכת למשתמש."""
    con = _connect()
    run = con.execute('SELECT * FROM runs WHERE id = ? AND user_id = ?',
                      (run_id, user_id)).fetchone()
    if run is None:
        con.close()
        return None, []
    rows = [dict(r) for r in con.execute(
        'SELECT task_no, task_name, personal_id, soldier_name, role, start_min, end_min, shared '
        'FROM assignments WHERE run_id = ? ORDER BY id', (run_id,))]
    con.close()
    run = dict(run)
    run['roster'] = [tuple(x) for x in json.loads(run.pop('roster_json'))]
    run['compare'] = json.loads(run.pop('compare_json'))
    return run, rows


def delete_runs(user_id: int):
    con = _connect()
    con.execute('DELETE FROM runs WHERE user_id = ?', (user_id,))
    con.commit()
    con.close()


# ── נתוני שימוש למנהל המערכת ─────────────────────────────────────
def admin_overview(days: int = 30) -> dict:
    con = _connect()
    since = (datetime.datetime.strptime(_now(), '%Y-%m-%d %H:%M:%S')
             - datetime.timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    week = (datetime.datetime.strptime(_now(), '%Y-%m-%d %H:%M:%S')
            - datetime.timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
    one = lambda q, p=(): con.execute(q, p).fetchone()[0]
    totals = {
        'users': one('SELECT COUNT(*) FROM users WHERE is_admin = 0'),
        'active_week': one('SELECT COUNT(DISTINCT e.user_id) FROM events e JOIN users u '
                           'ON u.id = e.user_id WHERE u.is_admin = 0 AND e.ts >= ?', (week,)),
        'logins': one("SELECT COUNT(*) FROM events e JOIN users u ON u.id = e.user_id "
                      "WHERE u.is_admin = 0 AND e.event_type = 'login'"),
        'runs': one('SELECT COUNT(*) FROM runs r JOIN users u ON u.id = r.user_id '
                    'WHERE u.is_admin = 0'),
        'avg_seconds': one('SELECT COALESCE(AVG(solve_seconds), 0) FROM runs'),
        'avg_coverage': one('SELECT COALESCE(AVG(coverage_pct), 0) FROM runs'),
    }
    users = pd.read_sql_query(
        "SELECT u.username AS 'שם משתמש', u.unit_name AS 'יחידה', u.created_at AS 'נרשם', "
        "(SELECT MAX(ts) FROM events e WHERE e.user_id = u.id) AS 'פעילות אחרונה', "
        "(SELECT COUNT(*) FROM events e WHERE e.user_id = u.id AND e.event_type = 'login') AS 'כניסות', "
        "(SELECT COUNT(*) FROM runs r WHERE r.user_id = u.id) AS 'הרצות', "
        "(SELECT COUNT(*) FROM soldiers s WHERE s.user_id = u.id) AS 'חיילים שמורים' "
        "FROM users u WHERE u.is_admin = 0 ORDER BY u.id DESC", con)
    daily = pd.read_sql_query(
        "SELECT substr(e.ts, 1, 10) AS day, e.event_type, COUNT(*) AS n, "
        "COUNT(DISTINCT e.user_id) AS users "
        "FROM events e JOIN users u ON u.id = e.user_id "
        "WHERE u.is_admin = 0 AND e.ts >= ? GROUP BY day, e.event_type ORDER BY day", con,
        params=(since,))
    daily['event'] = daily['event_type'].map(EVENT_LABELS).fillna(daily['event_type'])
    recent = pd.read_sql_query(
        "SELECT e.ts AS 'מתי', COALESCE(u.username, '—') AS 'משתמש', e.event_type AS 'פעולה', "
        "e.details AS 'פרטים' FROM events e LEFT JOIN users u ON u.id = e.user_id "
        "ORDER BY e.id DESC LIMIT 100", con)
    recent['פעולה'] = recent['פעולה'].map(EVENT_LABELS).fillna(recent['פעולה'])
    con.close()
    return {'totals': totals, 'users': users, 'daily': daily, 'recent': recent}
