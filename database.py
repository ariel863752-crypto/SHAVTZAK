"""
database.py — בסיס הנתונים של שבצ"ק

עובד מול שני סוגים של בסיס נתונים, עם אותו קוד בדיוק:

  * Postgres חיצוני (Neon וכדומה) — כשמוגדרת שורת חיבור DATABASE_URL.
    זה המצב הנכון לאתר החי: הנתונים נשמרים גם כשהשרת של האתר מאותחל.
  * קובץ SQLite מקומי (shavtzak.db) — כשאין שורת חיבור. מתאים לפיתוח על המחשב.

טבלאות:
  users        משתמשי המערכת (מפקדים). הסיסמה נשמרת כגיבוב (hash) בלבד.
  sessions     התחברויות פעילות, כדי שרענון הדף לא ינתק את המשתמש
  soldiers     החיילים של כל משתמש
  tasks        המשימות של כל משתמש
  runs         כל הרצת שיבוץ: מתי, של מי, ומה היו המדדים
  assignments  שורה לכל תקן בכל משמרת בכל הרצה (מי שובץ, או חוסר)
  events       יומן שימוש: כניסות, הרשמות, הרצות, הורדות

קשרים:  users 1─∞ sessions, soldiers, tasks, runs, events    |    runs 1─∞ assignments
"""
import datetime
import hashlib
import json
import os
import secrets
import sqlite3
import threading

import pandas as pd

try:
    import psycopg2
    import psycopg2.extras
except Exception:  # הספרייה נדרשת רק כשעובדים מול Postgres
    psycopg2 = None

ADMIN_USERNAME = 'admin'
SESSION_DAYS = 30

_url = (os.environ.get('DATABASE_URL') or '').strip()
_sqlite_path = os.environ.get('SHAVTZAK_DB', 'shavtzak.db')
_conn = None
_lock = threading.RLock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id          {PK},
    username    TEXT    NOT NULL UNIQUE,
    unit_name   TEXT    NOT NULL DEFAULT '',
    pw_hash     TEXT    NOT NULL,
    pw_salt     TEXT    NOT NULL,
    is_admin    INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT    NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
    token_hash  TEXT    PRIMARY KEY,
    user_id     INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at  TEXT    NOT NULL,
    expires_at  TEXT    NOT NULL
);
CREATE TABLE IF NOT EXISTS soldiers (
    id            {PK},
    user_id       INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    personal_id   TEXT    NOT NULL,
    full_name     TEXT    NOT NULL,
    exemptions    TEXT    NOT NULL DEFAULT '',
    roles         TEXT    NOT NULL DEFAULT '',
    blocked_hours TEXT    NOT NULL DEFAULT '',
    absences      TEXT    NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS tasks (
    id             {PK},
    user_id        INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    task_no        INTEGER NOT NULL,
    name           TEXT    NOT NULL,
    required       INTEGER NOT NULL,
    shift_hours    DOUBLE PRECISION NOT NULL,
    rest_hours     DOUBLE PRECISION NOT NULL DEFAULT 0,
    shared         INTEGER NOT NULL DEFAULT 0,
    active_hours   TEXT    NOT NULL DEFAULT '',
    required_roles TEXT    NOT NULL DEFAULT '',
    intensity      INTEGER NOT NULL DEFAULT 1,
    blocked_roles  TEXT    NOT NULL DEFAULT '',
    priority       INTEGER NOT NULL DEFAULT 5
);
CREATE TABLE IF NOT EXISTS runs (
    id              {PK},
    user_id         INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at      TEXT    NOT NULL,
    label           TEXT    NOT NULL,
    day_start       INTEGER NOT NULL,
    engine          TEXT    NOT NULL,
    num_soldiers    INTEGER NOT NULL,
    num_tasks       INTEGER NOT NULL,
    short_hours     DOUBLE PRECISION NOT NULL,
    coverage_pct    DOUBLE PRECISION NOT NULL,
    rest_violations INTEGER NOT NULL,
    full_rest       INTEGER NOT NULL,
    load_gap        DOUBLE PRECISION NOT NULL,
    solve_seconds   DOUBLE PRECISION NOT NULL,
    roster_json     TEXT    NOT NULL,
    compare_json    TEXT    NOT NULL,
    extra_json      TEXT    NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS assignments (
    id           {PK},
    run_id       INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    task_no      INTEGER NOT NULL,
    task_name    TEXT    NOT NULL,
    personal_id  TEXT,
    soldier_name TEXT,
    role         TEXT,
    start_min    INTEGER NOT NULL,
    end_min      INTEGER NOT NULL,
    shared       INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS events (
    id         {PK},
    user_id    INTEGER REFERENCES users(id) ON DELETE SET NULL,
    ts         TEXT    NOT NULL,
    event_type TEXT    NOT NULL,
    details    TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_soldiers_user ON soldiers(user_id);
CREATE INDEX IF NOT EXISTS idx_tasks_user    ON tasks(user_id);
CREATE INDEX IF NOT EXISTS idx_runs_user     ON runs(user_id);
CREATE INDEX IF NOT EXISTS idx_assign_run    ON assignments(run_id);
CREATE INDEX IF NOT EXISTS idx_events_ts     ON events(ts);
"""

EVENT_LABELS = {'register': 'הרשמה', 'login': 'כניסה', 'run': 'הרצת שיבוץ',
                'download': 'הורדת קובץ', 'demo': 'טעינת נתוני דוגמה', 'upload': 'עדכון נתונים'}


# ══════════════════════════════════════════════════════════════════
# חיבור
# ══════════════════════════════════════════════════════════════════
def configure(database_url=None):
    """קובע לאיזה בסיס נתונים מתחברים. שורת חיבור ריקה = קובץ SQLite מקומי."""
    global _url, _conn
    new = (database_url or '').strip()
    with _lock:
        if new != _url:
            _close()
            _url = new


def backend() -> str:
    return 'postgres' if _url else 'sqlite'


def _close():
    global _conn
    if _conn is not None:
        try:
            _conn.close()
        except Exception:
            pass
    _conn = None


def _open():
    if _url:
        if psycopg2 is None:
            raise RuntimeError('כדי להתחבר ל-Postgres צריך להתקין את psycopg2-binary '
                               '(מופיע ב-requirements.txt).')
        return psycopg2.connect(_url, connect_timeout=20,
                                cursor_factory=psycopg2.extras.DictCursor)
    con = sqlite3.connect(_sqlite_path, timeout=15, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys = ON')
    return con


def _sql(query: str) -> str:
    return query.replace('?', '%s') if _url else query


def _transaction(work):
    """
    מריץ work(cursor) כיחידה אחת: או שהכול נשמר, או ששום דבר.
    שירותי ענן סוגרים חיבור שלא היה בשימוש, ולכן אם החיבור נפל מתחברים מחדש ומנסים שוב.
    """
    global _conn
    dropped = (psycopg2.OperationalError, psycopg2.InterfaceError) if psycopg2 else ()
    with _lock:
        for attempt in (1, 2):
            try:
                if _conn is None:
                    _conn = _open()
                cur = _conn.cursor()
                try:
                    result = work(cur)
                    _conn.commit()
                    return result
                except Exception:
                    _conn.rollback()
                    raise
                finally:
                    try:
                        cur.close()
                    except Exception:
                        pass
            except dropped:
                _close()
                if attempt == 2:
                    raise


def _exec(query, params=()):
    return _transaction(lambda cur: cur.execute(_sql(query), params))


def _rows(query, params=()):
    def work(cur):
        cur.execute(_sql(query), params)
        return [dict(r) for r in cur.fetchall()]
    return _transaction(work)


def _one(query, params=()):
    rows = _rows(query, params)
    return rows[0] if rows else None


def _value(query, params=()):
    def work(cur):
        cur.execute(_sql(query), params)
        row = cur.fetchone()
        return row[0] if row else None
    return _transaction(work)


def _df(query, params=()) -> pd.DataFrame:
    def work(cur):
        cur.execute(_sql(query), params)
        cols = [d[0] for d in cur.description]
        return pd.DataFrame([tuple(r) for r in cur.fetchall()], columns=cols)
    return _transaction(work)


def _insert_id(cur, query, params):
    """INSERT של שורה אחת שמחזיר את המזהה שנוצר"""
    if _url:
        cur.execute(_sql(query) + ' RETURNING id', params)
        return cur.fetchone()[0]
    cur.execute(query, params)
    return cur.lastrowid


def _insert_rows(cur, table, cols, rows, chunk=60):
    """INSERT של הרבה שורות בפנייה אחת (חשוב כשבסיס הנתונים רחוק)"""
    one = '(' + ','.join(['?'] * len(cols)) + ')'
    for i in range(0, len(rows), chunk):
        part = rows[i:i + chunk]
        query = f"INSERT INTO {table} ({','.join(cols)}) VALUES " + ','.join([one] * len(part))
        cur.execute(_sql(query), [v for r in part for v in r])


def _now(shift_days: int = 0) -> str:
    try:
        from zoneinfo import ZoneInfo
        now = datetime.datetime.now(ZoneInfo('Asia/Jerusalem'))
    except Exception:
        now = datetime.datetime.now()
    return (now + datetime.timedelta(days=shift_days)).strftime('%Y-%m-%d %H:%M:%S')


# עמודות שנוספו אחרי הגרסה הראשונה. בסיס נתונים קיים מקבל אותן אוטומטית בעלייה.
MIGRATIONS = [
    ('soldiers', 'absences', "TEXT NOT NULL DEFAULT ''"),
    ('runs', 'extra_json', "TEXT NOT NULL DEFAULT '{}'"),
]


def _migrate():
    for table, column, ddl in MIGRATIONS:
        if _url:
            have = _value('SELECT COUNT(*) FROM information_schema.columns '
                          'WHERE table_name = ? AND column_name = ?', (table, column))
        else:
            have = any(r['name'] == column for r in _rows(f'PRAGMA table_info({table})'))
        if not have:
            _exec(f'ALTER TABLE {table} ADD COLUMN {column} {ddl}')


def init_db():
    pk = 'SERIAL PRIMARY KEY' if _url else 'INTEGER PRIMARY KEY AUTOINCREMENT'
    schema = SCHEMA.replace('{PK}', pk)
    if _url:
        _transaction(lambda cur: cur.execute(schema))
    else:
        with _lock:
            global _conn
            if _conn is None:
                _conn = _open()
            _conn.executescript(schema)
            _conn.commit()
    _migrate()


def status() -> dict:
    """למסך המנהל: איפה הנתונים נשמרים, והאם החיבור תקין"""
    try:
        n = _value('SELECT COUNT(*) FROM users')
        return {'backend': backend(), 'ok': True, 'users': n}
    except Exception as e:
        return {'backend': backend(), 'ok': False, 'error': str(e)}


# ══════════════════════════════════════════════════════════════════
# משתמשים והתחברות
# ══════════════════════════════════════════════════════════════════
def _hash(password: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac('sha256', str(password).encode('utf-8'),
                               bytes.fromhex(salt_hex), 200_000).hex()


def _user_dict(row) -> dict:
    return {'id': row['id'], 'username': row['username'], 'unit_name': row['unit_name'],
            'is_admin': bool(row['is_admin'])}


def _integrity_errors():
    errs = [sqlite3.IntegrityError]
    if psycopg2:
        errs.append(psycopg2.IntegrityError)
    return tuple(errs)


def create_user(username: str, password: str, unit_name: str = ''):
    """מחזיר (user, None) בהצלחה, או (None, הודעת שגיאה)."""
    username = (username or '').strip()
    unit_name = (unit_name or '').strip()
    if len(username) < 3:
        return None, 'שם המשתמש צריך להכיל לפחות 3 תווים.'
    if username.lower() == ADMIN_USERNAME:
        return None, 'שם המשתמש הזה שמור. בחרו שם אחר.'
    if len(password or '') < 6:
        return None, 'הסיסמה צריכה להכיל לפחות 6 תווים.'
    salt = secrets.token_hex(16)
    try:
        uid = _transaction(lambda cur: _insert_id(
            cur, 'INSERT INTO users (username, unit_name, pw_hash, pw_salt, is_admin, created_at) '
                 'VALUES (?,?,?,?,0,?)',
            (username, unit_name, _hash(password, salt), salt, _now())))
    except _integrity_errors():
        return None, 'שם המשתמש כבר תפוס. בחרו שם אחר.'
    log_event(uid, 'register')
    return {'id': uid, 'username': username, 'unit_name': unit_name, 'is_admin': False}, None


def authenticate(username: str, password: str):
    row = _one('SELECT * FROM users WHERE username = ?', ((username or '').strip(),))
    if row is None:
        return None
    if not secrets.compare_digest(_hash(password or '', row['pw_salt']), row['pw_hash']):
        return None
    log_event(row['id'], 'login')
    return _user_dict(row)


def ensure_admin(password):
    """יוצר או מעדכן את משתמש המנהל לפי הסיסמה שהוגדרה מחוץ לקוד."""
    if password is None or str(password) == '':
        return
    password = str(password)
    row = _one('SELECT id, pw_hash, pw_salt FROM users WHERE username = ?', (ADMIN_USERNAME,))
    salt = secrets.token_hex(16)
    if row is None:
        _exec('INSERT INTO users (username, unit_name, pw_hash, pw_salt, is_admin, created_at) '
              'VALUES (?,?,?,?,1,?)',
              (ADMIN_USERNAME, 'מנהל המערכת', _hash(password, salt), salt, _now()))
    elif _hash(password, row['pw_salt']) != row['pw_hash']:
        _exec('UPDATE users SET pw_hash = ?, pw_salt = ?, is_admin = 1 WHERE id = ?',
              (_hash(password, salt), salt, row['id']))


def _token_hash(token: str) -> str:
    return hashlib.sha256(token.encode('utf-8')).hexdigest()


def create_session(user_id: int) -> str:
    """מחזיר אסימון אקראי שנשמר בדפדפן. בבסיס הנתונים נשמר רק הגיבוב שלו."""
    token = secrets.token_urlsafe(32)
    _exec('INSERT INTO sessions (token_hash, user_id, created_at, expires_at) VALUES (?,?,?,?)',
          (_token_hash(token), user_id, _now(), _now(SESSION_DAYS)))
    return token


def user_from_session(token: str):
    if not token:
        return None
    row = _one('SELECT u.* FROM sessions s JOIN users u ON u.id = s.user_id '
               'WHERE s.token_hash = ? AND s.expires_at > ?', (_token_hash(token), _now()))
    if row is None:
        return None
    log_event(row['id'], 'login', 'חזרה לאתר')
    return _user_dict(row)


def delete_session(token: str):
    if token:
        _exec('DELETE FROM sessions WHERE token_hash = ?', (_token_hash(token),))


# ══════════════════════════════════════════════════════════════════
# יומן שימוש
# ══════════════════════════════════════════════════════════════════
def log_event(user_id, event_type: str, details: str = ''):
    _exec('INSERT INTO events (user_id, ts, event_type, details) VALUES (?,?,?,?)',
          (user_id, _now(), event_type, details))


# ══════════════════════════════════════════════════════════════════
# חיילים ומשימות
# ══════════════════════════════════════════════════════════════════
def _txt(v) -> str:
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'none') else s


def _num(v, default=0.0) -> float:
    try:
        f = float(v)
        return default if f != f else f      # NaN
    except (TypeError, ValueError):
        return default


def _flag(v) -> int:
    return 1 if (v is True or _txt(v).lower() in ('true', '1', '1.0', 'כן')) else 0


def save_roster(user_id: int, s_rows: list, t_rows: list):
    """מחליף את רשימות החיילים והמשימות של המשתמש."""
    soldiers = [(user_id, _txt(r.get('personal_id')).removesuffix('.0'), _txt(r.get('full_name')),
                 _txt(r.get('exemptions')).removesuffix('.0'), _txt(r.get('roles')),
                 _txt(r.get('blocked_hours')), _txt(r.get('absences')).removesuffix('.0'))
                for r in s_rows]
    tasks = [(user_id, int(_num(r.get('task_no'))), _txt(r.get('name')),
              max(1, int(_num(r.get('required'), 1))), _num(r.get('shift_hours'), 1) or 1,
              _num(r.get('rest_hours'), 0), _flag(r.get('shared')), _txt(r.get('active_hours')),
              _txt(r.get('required_roles')), int(_num(r.get('intensity'), 1) or 1),
              _txt(r.get('blocked_roles')), int(_num(r.get('priority'), 5) or 5)) for r in t_rows]

    def work(cur):
        cur.execute(_sql('DELETE FROM soldiers WHERE user_id = ?'), (user_id,))
        cur.execute(_sql('DELETE FROM tasks WHERE user_id = ?'), (user_id,))
        if soldiers:
            _insert_rows(cur, 'soldiers', ['user_id', 'personal_id', 'full_name', 'exemptions',
                                           'roles', 'blocked_hours', 'absences'], soldiers)
        if tasks:
            _insert_rows(cur, 'tasks', ['user_id', 'task_no', 'name', 'required', 'shift_hours',
                                        'rest_hours', 'shared', 'active_hours', 'required_roles',
                                        'intensity', 'blocked_roles', 'priority'], tasks)
    _transaction(work)


def load_roster(user_id: int):
    s_rows = _rows('SELECT personal_id, full_name, exemptions, roles, blocked_hours, absences '
                   'FROM soldiers WHERE user_id = ? ORDER BY id', (user_id,))
    t_rows = _rows('SELECT task_no, name, required, shift_hours, rest_hours, shared, active_hours, '
                   'required_roles, intensity, blocked_roles, priority '
                   'FROM tasks WHERE user_id = ? ORDER BY id', (user_id,))
    return s_rows, t_rows


# ══════════════════════════════════════════════════════════════════
# הרצות ושיבוצים
# ══════════════════════════════════════════════════════════════════
def save_run(user_id: int, label: str, day_start: int, engine: str, num_tasks: int,
             metrics: dict, seconds: float, roster: list, compare: list, rows: list,
             extra: dict = None) -> int:
    """extra: מידע נוסף על ההרצה (מספר ימים, מי לא נמצא באיזה יום)"""
    def work(cur):
        run_id = _insert_id(
            cur, 'INSERT INTO runs (user_id, created_at, label, day_start, engine, num_soldiers, '
                 'num_tasks, short_hours, coverage_pct, rest_violations, full_rest, load_gap, '
                 'solve_seconds, roster_json, compare_json, extra_json) '
                 'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (user_id, _now(), label, day_start, engine, len(roster), num_tasks,
             round(metrics['short_minutes'] / 60, 2), round(metrics['coverage_pct'], 2),
             len(metrics['rest_violations']), metrics['full_rest'], round(metrics['load_gap'], 2),
             seconds, json.dumps(roster, ensure_ascii=False),
             json.dumps(compare, ensure_ascii=False),
             json.dumps(extra or {}, ensure_ascii=False)))
        _insert_rows(cur, 'assignments',
                     ['run_id', 'task_no', 'task_name', 'personal_id', 'soldier_name', 'role',
                      'start_min', 'end_min', 'shared'],
                     [(run_id, r['task_no'], r['task_name'], r['personal_id'], r['soldier_name'],
                       r['role'], int(r['start_min']), int(r['end_min']), _flag(r['shared'])) for r in rows])
        return run_id
    return _transaction(work)


def list_runs(user_id: int) -> pd.DataFrame:
    return _df('SELECT id, created_at, label, num_soldiers, num_tasks, short_hours, coverage_pct, '
               'rest_violations, full_rest, load_gap, solve_seconds '
               'FROM runs WHERE user_id = ? ORDER BY id DESC', (user_id,))


def load_run(run_id: int, user_id: int):
    """מחזיר (פרטי ההרצה, שורות השיבוץ) — רק אם ההרצה שייכת למשתמש."""
    run = _one('SELECT * FROM runs WHERE id = ? AND user_id = ?', (run_id, user_id))
    if run is None:
        return None, []
    rows = _rows('SELECT task_no, task_name, personal_id, soldier_name, role, start_min, end_min, '
                 'shared FROM assignments WHERE run_id = ? ORDER BY id', (run_id,))
    run['roster'] = [tuple(x) for x in json.loads(run.pop('roster_json'))]
    run['compare'] = json.loads(run.pop('compare_json'))
    run['extra'] = json.loads(run.pop('extra_json', None) or '{}')
    return run, rows


def delete_runs(user_id: int):
    _exec('DELETE FROM runs WHERE user_id = ?', (user_id,))


# ══════════════════════════════════════════════════════════════════
# נתוני שימוש למנהל המערכת
# ══════════════════════════════════════════════════════════════════
def admin_overview(days: int = 30) -> dict:
    since, week = _now(-days), _now(-7)
    totals = {
        'users': _value('SELECT COUNT(*) FROM users WHERE is_admin = 0'),
        'active_week': _value('SELECT COUNT(DISTINCT e.user_id) FROM events e JOIN users u '
                              'ON u.id = e.user_id WHERE u.is_admin = 0 AND e.ts >= ?', (week,)),
        'logins': _value("SELECT COUNT(*) FROM events e JOIN users u ON u.id = e.user_id "
                         "WHERE u.is_admin = 0 AND e.event_type = 'login'"),
        'runs': _value('SELECT COUNT(*) FROM runs r JOIN users u ON u.id = r.user_id '
                       'WHERE u.is_admin = 0'),
        'avg_seconds': float(_value('SELECT COALESCE(AVG(solve_seconds), 0) FROM runs') or 0),
        'avg_coverage': float(_value('SELECT COALESCE(AVG(coverage_pct), 0) FROM runs') or 0),
    }
    users = _df(
        "SELECT u.username, u.unit_name, u.created_at, "
        "(SELECT MAX(ts) FROM events e WHERE e.user_id = u.id) AS last_seen, "
        "(SELECT COUNT(*) FROM events e WHERE e.user_id = u.id AND e.event_type = 'login') AS logins, "
        "(SELECT COUNT(*) FROM runs r WHERE r.user_id = u.id) AS runs, "
        "(SELECT COUNT(*) FROM soldiers s WHERE s.user_id = u.id) AS soldiers "
        "FROM users u WHERE u.is_admin = 0 ORDER BY u.id DESC")
    users.columns = ['שם משתמש', 'יחידה', 'נרשם', 'פעילות אחרונה', 'כניסות', 'הרצות', 'חיילים שמורים']
    daily = _df(
        "SELECT substr(e.ts, 1, 10) AS day, e.event_type, COUNT(*) AS n, "
        "COUNT(DISTINCT e.user_id) AS users "
        "FROM events e JOIN users u ON u.id = e.user_id "
        "WHERE u.is_admin = 0 AND e.ts >= ? "
        "GROUP BY substr(e.ts, 1, 10), e.event_type ORDER BY 1", (since,))
    daily['event'] = daily['event_type'].map(EVENT_LABELS).fillna(daily['event_type'])
    recent = _df(
        "SELECT e.ts, COALESCE(u.username, '—') AS username, e.event_type, e.details "
        "FROM events e LEFT JOIN users u ON u.id = e.user_id ORDER BY e.id DESC LIMIT 100")
    recent['event_type'] = recent['event_type'].map(EVENT_LABELS).fillna(recent['event_type'])
    recent.columns = ['מתי', 'משתמש', 'פעולה', 'פרטים']
    return {'totals': totals, 'users': users, 'daily': daily, 'recent': recent}
