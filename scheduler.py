"""
scheduler.py — מנוע השיבוץ של שבצ"ק (ללא תלות ב-Streamlit)

עקרונות:
  * ציר הזמן הוא יממה רציפה אחת של 24 שעות, שמתחילה ב"שעת תחילת היממה"
    (ברירת מחדל 06:00). משימה שחוצה חצות פשוט ממשיכה על הציר.
  * כל הזמנים נשמרים בדקות. טווח "8-20" הוא 08:00 עד 20:00 (הסיום אינו כלול),
    ו-"08:30-16:00" הוא בדיוק זה.
  * משתנה ההחלטה הוא (חייל, משמרת, תקן) ולא (חייל, שעה). משמרת היא יחידה שלמה.
  * האופטימיזציה נעשית בשלבים, לפי סדר עדיפות קשיח:
        1. כיסוי   — מינימום עמדות לא מאוישות (גובר על הכול)
        2. רווחה   — מנוחה אחרי משמרת + מנוחה רצופה, עדיף בלילה
        3. הוגנות — איזון עומס משוקלל (שעות × עצימות, תוספת ללילה)
    כל שלב ננעל לפני שעוברים לבא אחריו.
  * פונקציה אחת, evaluate(), מודדת כל שבצ"ק (של המערכת, של השיבוץ הבסיסי
    או ידני) באותם מדדים בדיוק.

אילוצים קשיחים:
  H1  כל תקן מאויש בחייל אחד, או מסומן כחוסר כשאין אפשרות מתמטית
  H2  הסמכה נדרשת לתקן
  H3  חסימות אישיות, פטורים ותפקידים חסומים
  H4  חייל לא נמצא בשתי משימות בלעדיות באותו זמן.
      משימה שיתופית (כוננות) מותרת במקביל לכל דבר.
  H5  חייל לא תופס שני תקנים באותה משמרת
  H6  הנחיות מפקד: חובה / אסור
"""
from __future__ import annotations

import math
import os
import re
import threading
import time
from dataclasses import dataclass, field
from functools import reduce
from typing import Callable, Optional

import pandas as pd

try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except Exception:  # pragma: no cover
    cp_model = None
    HAS_ORTOOLS = False

DAY = 24 * 60


# ══════════════════════════════════════════════════════════════════
# 1. הגדרות — מקור אמת אחד לכל המשקלים
# ══════════════════════════════════════════════════════════════════
@dataclass
class Config:
    day_start: int = 6 * 60        # תחילת יממת השיבוץ (דקות מחצות)
    night_start: int = 22 * 60     # חלון הלילה המועדף למנוחה
    night_end: int = 6 * 60
    sleep_target: int = 7 * 60     # יעד מנוחה רצופה לחייל ביממה
    sleep_min: int = 4 * 60        # מתחת לזה נחשב "אין מנוחה רצופה"

    # שלב 2 — רווחה
    w_rest: int = 1                # × (חצאי-שעות חסרות ממנוחה אחרי משמרת)²
    w_sleep_short: int = 40        # × (שעות חסרות מיעד המנוחה הרצופה)²
    w_sleep_none: int = 2000       # אין בלוק מנוחה של sleep_min לפחות
    w_sleep_day: int = 2           # לכל חצי שעה של בלוק המנוחה שמחוץ ללילה

    # שלב 3 — הוגנות
    night_load_factor: float = 1.5  # שעת לילה במשימה בלעדית שווה 1.5 שעות יום

    wellbeing_slack: float = 0.0    # 0 = סדר עדיפות קשיח: הוגנות לא באה על חשבון מנוחה
    time_limit: float = 30.0        # תקרת זמן כוללת (שניות)
    patience: float = 5.0           # שלב נעצר מוקדם אם לא נמצא שיפור במשך כך וכך שניות
    workers: int = max(1, min(8, os.cpu_count() or 1))
    seed: int = 1


# ══════════════════════════════════════════════════════════════════
# 2. פענוח זמנים
# ══════════════════════════════════════════════════════════════════
_ALL_WORDS = {'all', 'הכל', 'הכול', 'כל היום', '24/7'}
_TRUE_WORDS = {'true', '1', '1.0', 'כן', 'yes', 'y', 'v', 'אמת', 'נכון'}


def _is_blank(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    return str(val).strip().lower() in ('', 'nan', 'none', 'nat')


def parse_clock(tok: str) -> Optional[int]:
    """'8' / '08' / '8:30' / '08:30:00' / '8.5' → דקות מחצות. לא תקין → None"""
    tok = str(tok).strip()
    m = re.fullmatch(r'(\d{1,2})(?::(\d{1,2}))?(?::\d{1,2})?', tok)
    if m:
        h, mi = int(m.group(1)), int(m.group(2) or 0)
    else:
        m = re.fullmatch(r'(\d{1,2})\.(\d{1,2})', tok)
        if not m:
            return None
        h, frac = int(m.group(1)), m.group(2)
        # '8.30' = 08:30 (כתיב שעון) ; '8.5' = 08:30 ; '8.0' = 08:00 (שבר עשרוני)
        mi = int(frac) if len(frac) == 2 else int(frac) * 6
    if h > 24 or mi > 59 or (h == 24 and mi):
        return None
    return h * 60 + mi


def parse_ranges(val, blank_means_all: bool):
    """
    מפענח תא של שעות. מחזיר (intervals, bad_tokens).
    intervals: רשימת (התחלה, סיום) בדקות שעון; הסיום אינו כלול ויכול לעבור 1440
               כשהטווח חוצה חצות ("22-6" → (1320, 1800)).
    מספר בודד ("7") הוא השעה 07:00–08:00.
    """
    if _is_blank(val):
        return ([(0, DAY)] if blank_means_all else []), []
    text = str(val).strip()
    if text.lower() in _ALL_WORDS:
        return [(0, DAY)], []

    out, bad = [], []
    for part in re.split(r'[,;\n]+', text):
        part = part.strip().replace('–', '-').replace('—', '-').replace('־', '-')
        if not part:
            continue
        if '-' in part:
            pieces = [p for p in part.split('-')]
            s, e = parse_clock(pieces[0]), parse_clock(pieces[-1])
            if s is None or e is None:
                bad.append(part)
                continue
            s %= DAY
            if e <= s:
                e += DAY          # חוצה חצות ("22-6"), או יממה שלמה ("8-8")
            out.append((s, e))
        else:
            s = parse_clock(part)
            if s is None:
                bad.append(part)
                continue
            s %= DAY
            out.append((s, s + 60))
    return out, bad


def _merge(intervals):
    res = []
    for a, b in sorted(intervals):
        if res and a <= res[-1][1]:
            res[-1] = (res[-1][0], max(res[-1][1], b))
        else:
            res.append((a, b))
    return res


def to_horizon(clock_intervals, day_start: int):
    """ממיר טווחי שעון לדקות יחסיות לתחילת יממת השיבוץ, בתוך [0, 1440)."""
    out = []
    for s, e in clock_intervals:
        length = min(e - s, DAY)
        a = (s - day_start) % DAY
        b = a + length
        if b <= DAY:
            out.append((a, b))
        else:
            out.append((a, DAY))
            out.append((0, b - DAY))
    return _merge(out)


def fmt_clock(rel_min: int, day_start: int) -> str:
    m = (day_start + rel_min) % DAY
    return f"{m // 60:02d}:{m % 60:02d}"


def fmt_dur(minutes: int) -> str:
    h, m = divmod(int(minutes), 60)
    return f"{h}:{m:02d}"


def _overlaps(a1, b1, a2, b2) -> bool:
    return a1 < b2 and a2 < b1


def _overlap_len(a, b, intervals) -> int:
    return sum(max(0, min(b, y) - max(a, x)) for x, y in intervals)


# ══════════════════════════════════════════════════════════════════
# 3. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def _split_list(val) -> list:
    if _is_blank(val):
        return []
    return [p.strip() for p in re.split(r'[,;]+', str(val)) if p.strip()]


def _to_int(val, default: int) -> int:
    try:
        if _is_blank(val):
            return default
        return int(float(val))
    except (TypeError, ValueError):
        return default


def _hours_to_min(val, default_hours: float) -> int:
    try:
        if _is_blank(val):
            return int(round(default_hours * 60))
        return int(round(float(val) * 60))
    except (TypeError, ValueError):
        return int(round(default_hours * 60))


def _norm_id(val) -> str:
    s = str(val).strip()
    return s[:-2] if s.endswith('.0') else s


@dataclass
class Soldier:
    sid: str
    name: str
    exempt: set = field(default_factory=set)        # מס"ד משימות שהחייל פטור מהן
    roles: set = field(default_factory=set)         # הסמכות / תפקידים
    blocked_clock: list = field(default_factory=list)
    blocked: list = field(default_factory=list)     # על ציר היממה (מתמלא ב-prepare)


@dataclass
class Task:
    tid: int
    name: str
    req: int
    shift: int                 # משך משמרת בדקות
    rest: int                  # מנוחה נדרשת אחרי משמרת, בדקות
    shared: bool               # True = שיתופית (כוננות): מותרת במקביל לכל דבר
    hours_clock: list
    slots: list                # הסמכה נדרשת לכל תקן, או None
    intensity: int = 5
    block_roles: set = field(default_factory=set)
    priority: int = 5          # קריטיות: על מה מוותרים אחרון כשיש חוסר

    @property
    def groups(self):
        """[(הסמכה או None, כמות)] — תקנים זהים מקובצים יחד"""
        order, counts = [], {}
        for r in self.slots:
            if r not in counts:
                order.append(r)
                counts[r] = 0
            counts[r] += 1
        order.sort(key=lambda r: r is None)   # תקנים עם הסמכה קודם
        return [(r, counts[r]) for r in order]


@dataclass
class Shift:
    idx: int
    task: Task
    start: int       # דקות יחסיות לתחילת היממה
    end: int
    load: int = 0    # עומס משוקלל (מתמלא ב-build_shifts)
    night: int = 0   # דקות בתוך חלון הלילה

    @property
    def dur(self) -> int:
        return self.end - self.start


@dataclass
class Directive:
    """הנחיית מפקד (H6). kind: 'must' = חובה, 'must_not' = אסור"""
    sid: str
    tid: int
    kind: str
    hours_clock: list = field(default_factory=list)   # ריק = כל משמרות המשימה


def soldier_from_row(row: dict) -> tuple:
    blocked, bad = parse_ranges(row.get('blocked_hours'), blank_means_all=False)
    s = Soldier(
        sid=_norm_id(row.get('personal_id')),
        name=str(row.get('full_name', '')).strip(),
        exempt={int(float(x)) for x in _split_list(row.get('exemptions'))
                if re.fullmatch(r'\d+(\.0)?', x)},
        roles=set(_split_list(row.get('roles'))),
        blocked_clock=blocked,
    )
    warn = [f'חייל {s.name}: לא הצלחתי לפענח את שעות החסימה "{b}"' for b in bad]
    return s, warn


def task_from_row(row: dict) -> tuple:
    hours, bad = parse_ranges(row.get('active_hours'), blank_means_all=True)
    req = max(1, _to_int(row.get('required'), 1))
    slots = _split_list(row.get('required_roles'))
    slots = slots + [None] * max(0, req - len(slots))
    name = str(row.get('name', '')).strip()
    shared_raw = row.get('shared')
    shared = (shared_raw is True) or (str(shared_raw).strip().lower() in _TRUE_WORDS)
    t = Task(
        tid=_to_int(row.get('task_no'), 0),
        name=name,
        req=len(slots),
        shift=max(15, _hours_to_min(row.get('shift_hours'), 1)),
        rest=max(0, _hours_to_min(row.get('rest_hours'), 0)),
        shared=shared,
        hours_clock=hours,
        slots=slots,
        intensity=max(1, min(10, _to_int(row.get('intensity'), 1))),
        block_roles=set(_split_list(row.get('blocked_roles'))),
        priority=max(1, min(10, _to_int(row.get('priority'), 5))),
    )
    warn = [f'משימה {name}: לא הצלחתי לפענח את שעות הפעילות "{b}"' for b in bad]
    return t, warn


def directive_from_row(row: dict) -> Optional[Directive]:
    kind_raw = str(row.get('kind', '')).strip().lower()
    if kind_raw in ('חובה', 'must', 'חייב'):
        kind = 'must'
    elif kind_raw in ('אסור', 'must_not', 'must not', 'אסור לשבץ'):
        kind = 'must_not'
    else:
        return None
    hours, _ = parse_ranges(row.get('hours'), blank_means_all=False)
    return Directive(sid=_norm_id(row.get('personal_id')),
                     tid=_to_int(row.get('task_no'), 0), kind=kind, hours_clock=hours)


# ── קריאת קובצי Excel לשורות בפורמט אחיד ─────────────────────────
def _find_col(df: pd.DataFrame, keywords) -> Optional[str]:
    for kw in keywords:
        for c in df.columns:
            if kw in str(c):
                return c
    return None


def _cell(row, col, default=''):
    if col is None:
        return default
    v = row.get(col, default)
    return default if _is_blank(v) else v


def rows_from_excel(s_df: pd.DataFrame, t_df: pd.DataFrame):
    """ממיר את שתי טבלאות הקלט (בעמודות העבריות של התבנית) לרשימות של dict."""
    sc = {
        'personal_id':   _find_col(s_df, ['מספר אישי', 'מספר', 'אישי']),
        'full_name':     _find_col(s_df, ['שם']),
        'exemptions':    _find_col(s_df, ['פטור']),
        'roles':         _find_col(s_df, ['הסמכ', 'תפקיד']),
        'blocked_hours': _find_col(s_df, ['חסימה', 'חסום']),
    }
    tc = {
        'task_no':        _find_col(t_df, ['מס"ד', 'מסד', 'מספר']),
        'name':           _find_col(t_df, ['שם']),
        'required':       _find_col(t_df, ['סד"כ', 'סדכ', 'נדרש למשימה']),
        'shift_hours':    _find_col(t_df, ['משך']),
        'rest_hours':     _find_col(t_df, ['מנוחה']),
        'shared':         _find_col(t_df, ['חפיפה', 'שיתופי']),
        'active_hours':   _find_col(t_df, ['פעילות']),
        'required_roles': _find_col(t_df, ['הסמכה']),
        'intensity':      _find_col(t_df, ['עצימות']),
        'blocked_roles':  _find_col(t_df, ['חסומים']),
        'priority':       _find_col(t_df, ['קריטיות', 'עדיפות']),
    }
    missing = [k for k in ('personal_id', 'full_name') if sc[k] is None]
    missing += [k for k in ('task_no', 'name', 'required') if tc[k] is None]
    if missing:
        raise ValueError('לא נמצאו עמודות חובה בקבצים. ודאו שהקבצים בנויים לפי התבנית.')

    s_rows = []
    for _, r in s_df.iterrows():
        if _is_blank(r.get(sc['personal_id'])) or _is_blank(r.get(sc['full_name'])):
            continue
        s_rows.append({k: _cell(r, c) for k, c in sc.items()})
    t_rows = []
    for _, r in t_df.iterrows():
        if _is_blank(r.get(tc['task_no'])) or _is_blank(r.get(tc['name'])) \
                or _is_blank(r.get(tc['required'])):
            continue
        d = {k: _cell(r, c) for k, c in tc.items()}
        if tc['shared'] is not None:
            d['shared'] = (r.get(tc['shared']) is True) or \
                (str(r.get(tc['shared'])).strip().lower() in _TRUE_WORDS)
        t_rows.append(d)
    return s_rows, t_rows


def directives_from_excel(d_df: pd.DataFrame) -> list:
    cols = {
        'personal_id': _find_col(d_df, ['מספר אישי', 'אישי']),
        'task_no':     _find_col(d_df, ['מס"ד', 'מסד', 'משימה']),
        'kind':        _find_col(d_df, ['סוג', 'הנחיה']),
        'hours':       _find_col(d_df, ['שעות']),
    }
    if cols['personal_id'] is None or cols['task_no'] is None or cols['kind'] is None:
        raise ValueError('קובץ הנחיות המפקד חייב לכלול: מספר אישי, מס"ד משימה, סוג.')
    out = []
    for _, r in d_df.iterrows():
        if _is_blank(r.get(cols['personal_id'])):
            continue
        out.append({k: _cell(r, c) for k, c in cols.items()})
    return out


def build_objects(s_rows, t_rows, d_rows=None):
    """שורות גולמיות → אובייקטים. מחזיר (soldiers, tasks, directives, warnings)."""
    warnings = []
    soldiers, seen = [], set()
    for r in s_rows:
        s, w = soldier_from_row(r)
        warnings += w
        if s.sid in seen:
            warnings.append(f'מספר אישי {s.sid} מופיע פעמיים — נלקחה רק השורה הראשונה')
            continue
        seen.add(s.sid)
        soldiers.append(s)
    tasks, seen_t = [], set()
    for r in t_rows:
        t, w = task_from_row(r)
        warnings += w
        if t.tid in seen_t:
            warnings.append(f'מס"ד משימה {t.tid} מופיע פעמיים — נלקחה רק השורה הראשונה')
            continue
        seen_t.add(t.tid)
        tasks.append(t)
    directives = []
    for r in (d_rows or []):
        d = directive_from_row(r)
        if d is None:
            warnings.append(f'הנחיית מפקד לא מזוהה (סוג חייב להיות "חובה" או "אסור"): {r}')
        else:
            directives.append(d)
    return soldiers, tasks, directives, warnings


# ══════════════════════════════════════════════════════════════════
# 4. בניית משמרות על ציר היממה
# ══════════════════════════════════════════════════════════════════
def night_intervals(cfg: Config):
    end = cfg.night_end + (DAY if cfg.night_end <= cfg.night_start else 0)
    return to_horizon([(cfg.night_start, end)], cfg.day_start)


def prepare(soldiers, tasks, cfg: Config):
    """ממקם חסימות ומשמרות על ציר היממה. מחזיר (shifts, warnings)."""
    for s in soldiers:
        s.blocked = to_horizon(s.blocked_clock, cfg.day_start)

    night = night_intervals(cfg)
    nf2 = int(round(cfg.night_load_factor * 2))
    shifts, warnings = [], []
    for t in tasks:
        # משימה שהחלון שלה עובר דרך שעת תחילת היממה נחתכת לשני חלקים
        cut = any(e - s < DAY and ((s - cfg.day_start) % DAY) + (e - s) > DAY
                  for s, e in t.hours_clock)
        if cut:
            warnings.append(
                f'משימה {t.name} עוברת דרך שעת תחילת היממה ({fmt_clock(0, cfg.day_start)}) '
                f'ולכן נחתכה לשני חלקים. כדי שתישאר רצופה, בחרו שעת תחילת יממה שמחוץ לשעות המשימה.')
        for a, b in to_horizon(t.hours_clock, cfg.day_start):
            cur = a
            while cur < b:
                end = min(cur + t.shift, b)
                if end - cur < t.shift and not cut:
                    warnings.append(
                        f'משימה {t.name}: חלון הפעילות אינו כפולה של משך המשמרת — '
                        f'המשמרת {fmt_clock(cur, cfg.day_start)}–{fmt_clock(end, cfg.day_start)} '
                        f'קצרה ({fmt_dur(end - cur)} שעות)')
                sh = Shift(idx=len(shifts), task=t, start=cur, end=end)
                sh.night = 0 if t.shared else _overlap_len(cur, end, night)
                day_min = sh.dur - sh.night
                sh.load = t.intensity * (2 * day_min + nf2 * sh.night)
                shifts.append(sh)
                cur = end
    return shifts, warnings


def _base_eligible(s: Soldier, t: Task) -> bool:
    return t.tid not in s.exempt and not (s.roles & t.block_roles)


def _available(s: Soldier, sh: Shift) -> bool:
    return not any(_overlaps(sh.start, sh.end, a, b) for a, b in s.blocked)


def eligibility(soldiers, shifts):
    """elig[si][k] = רשימת אינדקסים של קבוצות-תקן שהחייל יכול למלא במשמרת k"""
    elig = []
    for s in soldiers:
        row = []
        for sh in shifts:
            if not _base_eligible(s, sh.task) or not _available(s, sh):
                row.append([])
            else:
                row.append([gi for gi, (role, _) in enumerate(sh.task.groups)
                            if role is None or role in s.roles])
        elig.append(row)
    return elig


def _grid(shifts) -> int:
    """רזולוציית הזמן הנדרשת: 60, 30 או 15 דקות"""
    vals = [60]
    for sh in shifts:
        vals += [sh.start, sh.end]
    g = reduce(math.gcd, vals)
    return 60 if g % 60 == 0 else (30 if g % 30 == 0 else 15)


def sleep_windows(cfg: Config, grid: int, night):
    """כל חלונות המנוחה האפשריים: (התחלה, סיום, עלות). עלות 0 = יעד מלא בלילה."""
    out = []
    dur = cfg.sleep_target
    while dur >= cfg.sleep_min:
        miss_h = math.ceil((cfg.sleep_target - dur) / 60)
        for a in range(0, DAY - dur + 1, grid):
            outside = dur - _overlap_len(a, a + dur, night)
            cost = cfg.w_sleep_short * miss_h * miss_h + cfg.w_sleep_day * math.ceil(outside / 30)
            out.append((a, a + dur, cost))
        dur -= 60
    return out


def _directive_shifts(d: Directive, shifts, cfg: Config):
    iv = to_horizon(d.hours_clock, cfg.day_start) if d.hours_clock else None
    return [sh for sh in shifts if sh.task.tid == d.tid
            and (iv is None or any(_overlaps(sh.start, sh.end, a, b) for a, b in iv))]


# ══════════════════════════════════════════════════════════════════
# 5. שיבוץ בסיסי — מחקה שיבוץ ידני: עובר על המשמרות לפי סדר הזמן
#    ולוקח בכל פעם את החייל הפנוי שעבד הכי פחות.
#    משמש כנקודת פתיחה לפותר וכבסיס להשוואת "לפני / אחרי".
# ══════════════════════════════════════════════════════════════════
def greedy_baseline(soldiers, shifts, elig, cfg: Config, directives=None):
    n = len(soldiers)
    sid_idx = {s.sid: i for i, s in enumerate(soldiers)}
    busy = {i: [] for i in range(n)}      # משמרות בלעדיות של החייל
    load = [0] * n
    in_shift = set()
    filled = {}
    assign = []

    banned = set()
    for d in (directives or []):
        if d.kind == 'must_not' and d.sid in sid_idx:
            for sh in _directive_shifts(d, shifts, cfg):
                banned.add((sid_idx[d.sid], sh.idx))

    def rest_short(si, sh):
        worst = 0
        for o in busy[si]:
            if o.end <= sh.start:
                worst = max(worst, o.end + o.task.rest - sh.start)
            elif sh.end <= o.start:
                worst = max(worst, sh.end + sh.task.rest - o.start)
        return max(0, worst)

    def can(si, sh, gi):
        if gi not in elig[si][sh.idx] or (si, sh.idx) in in_shift or (si, sh.idx) in banned:
            return False
        if not sh.task.shared and any(_overlaps(sh.start, sh.end, o.start, o.end) for o in busy[si]):
            return False
        return True

    def put(si, sh, gi):
        assign.append((soldiers[si].sid, sh.idx, sh.task.groups[gi][0]))
        in_shift.add((si, sh.idx))
        filled[sh.idx, gi] = filled.get((sh.idx, gi), 0) + 1
        load[si] += sh.load
        if not sh.task.shared:
            busy[si].append(sh)

    for d in (directives or []):
        if d.kind != 'must' or d.sid not in sid_idx:
            continue
        si = sid_idx[d.sid]
        done = False
        for sh in sorted(_directive_shifts(d, shifts, cfg), key=lambda x: x.start):
            for gi, (_, cnt) in enumerate(sh.task.groups):
                if filled.get((sh.idx, gi), 0) < cnt and can(si, sh, gi):
                    put(si, sh, gi)
                    done = True
                    break
            if done:
                break

    order = sorted(shifts, key=lambda sh: (sh.start, sh.task.shared, -sh.task.priority, sh.idx))
    for sh in order:
        for gi, (_, cnt) in enumerate(sh.task.groups):
            while filled.get((sh.idx, gi), 0) < cnt:
                cands = [si for si in range(n) if can(si, sh, gi)]
                if not cands:
                    break
                best = min(cands, key=lambda si: (
                    0 if sh.task.shared else rest_short(si, sh),
                    load[si], len(soldiers[si].roles), si))
                put(best, sh, gi)
    return assign


# ══════════════════════════════════════════════════════════════════
# 6. הערכה — פונקציה אחת שמודדת כל שבצ"ק באותם מדדים
# ══════════════════════════════════════════════════════════════════
def _free_blocks(excl_shifts):
    """בלוקים פנויים ממשימה בלעדית בתוך היממה"""
    blocks, cur = [], 0
    for sh in sorted(excl_shifts, key=lambda x: x.start):
        if sh.start > cur:
            blocks.append((cur, sh.start))
        cur = max(cur, sh.end)
    if cur < DAY:
        blocks.append((cur, DAY))
    return blocks


def evaluate(soldiers, shifts, assign, cfg: Config) -> dict:
    """
    assign: רשימת (מספר אישי, אינדקס משמרת, הסמכת התקן או None)
    מחזיר את כל המדדים, כולל בדיקה עצמאית של האילוצים הקשיחים.
    """
    by_sid = {s.sid: s for s in soldiers}
    night = night_intervals(cfg)
    grid = _grid([sh for sh in shifts if not sh.task.shared] or shifts)
    windows = sleep_windows(cfg, grid, night)
    nf2 = int(round(cfg.night_load_factor * 2))

    hard = []
    per_shift_group = {}
    per_soldier = {s.sid: [] for s in soldiers}
    seen_pair = set()
    for sid, k, role in assign:
        sh = shifts[k]
        s = by_sid.get(sid)
        if s is None:
            hard.append(f'חייל לא מוכר: {sid}')
            continue
        if (sid, k) in seen_pair:
            hard.append(f'H5: {s.name} משובץ פעמיים במשמרת {sh.task.name}')
        seen_pair.add((sid, k))
        if role is not None and role not in s.roles:
            hard.append(f'H2: ל{s.name} אין הסמכת {role} ({sh.task.name})')
        if not _base_eligible(s, sh.task):
            hard.append(f'H3: {s.name} פטור או חסום מ{sh.task.name}')
        if not _available(s, sh):
            hard.append(f'H3: {s.name} משובץ ל{sh.task.name} בשעות החסימה שלו')
        per_shift_group[k, role] = per_shift_group.get((k, role), 0) + 1
        per_soldier[sid].append(sh)

    # ── כיסוי ─────────────────────────────────────────────────────
    shortages, short_min, short_weighted, slots_total, slot_min_total = [], 0, 0, 0, 0
    for sh in shifts:
        for role, cnt in sh.task.groups:
            slots_total += cnt
            slot_min_total += cnt * sh.dur
            got = per_shift_group.get((sh.idx, role), 0)
            if got > cnt:
                hard.append(f'H1: יותר מדי חיילים בתקן של {sh.task.name}')
            miss = max(0, cnt - got)
            if miss:
                short_min += miss * sh.dur
                short_weighted += miss * sh.dur * sh.task.priority
                shortages.append({'shift': sh.idx, 'role': role, 'missing': miss})

    # ── לכל חייל ──────────────────────────────────────────────────
    stats, rest_viol, rest_cost, sleep_cost = {}, [], 0, 0
    for s in soldiers:
        mine = per_soldier[s.sid]
        excl = sorted([sh for sh in mine if not sh.task.shared], key=lambda x: x.start)
        for i in range(len(excl)):
            for j in range(i + 1, len(excl)):
                if _overlaps(excl[i].start, excl[i].end, excl[j].start, excl[j].end):
                    hard.append(f'H4: {s.name} בשתי משימות בלעדיות באותו זמן '
                                f'({excl[i].task.name}, {excl[j].task.name})')
        # מנוחה אחרי משמרת: עלות לפי כל זוג (כמו במודל), דיווח לפי זוגות עוקבים
        for i in range(len(excl)):
            for j in range(len(excl)):
                a, b = excl[i], excl[j]
                if i != j and a.end <= b.start < a.end + a.task.rest:
                    u = math.ceil((a.end + a.task.rest - b.start) / 30)
                    rest_cost += cfg.w_rest * u * u
        for a, b in zip(excl, excl[1:]):
            if a.end <= b.start < a.end + a.task.rest:
                rest_viol.append({
                    'sid': s.sid, 'name': s.name,
                    'after': a.task.name, 'before': b.task.name,
                    'at': a.end, 'required': a.task.rest, 'actual': b.start - a.end,
                })
        # מנוחה רצופה
        blocks = _free_blocks(excl)
        longest = max((b - a for a, b in blocks), default=0)
        night_rest = 0
        for a, b in blocks:
            for x, y in night:
                night_rest = max(night_rest, min(b, y) - max(a, x))
        best = cfg.w_sleep_none
        for wa, wb, cost in windows:
            if cost < best and not any(_overlaps(wa, wb, sh.start, sh.end) for sh in excl):
                best = cost
        sleep_cost += best

        excl_min = sum(sh.dur for sh in excl)
        shared_min = sum(sh.dur for sh in mine if sh.task.shared)
        stats[s.sid] = {
            'name': s.name,
            'minutes': excl_min,
            'shared_minutes': shared_min,
            'night_minutes': sum(sh.night for sh in excl),
            'load': sum(sh.load for sh in mine),
            'shifts': len(excl),
            'longest_rest': longest,
            'night_rest': max(0, night_rest),
            'sleep_cost': best,
        }

    loads = [st['load'] for st in stats.values()] or [0]
    mins = [st['minutes'] for st in stats.values()] or [0]
    unit = 2 * 60  # יחידת עומס = שעת יום אחת בעצימות 1
    return {
        'hard_violations': hard,
        'slots_total': slots_total,
        'short_slots': sum(x['missing'] for x in shortages),
        'short_minutes': short_min,
        'short_weighted': short_weighted,
        'coverage_pct': 100.0 * (1 - short_min / slot_min_total) if slot_min_total else 100.0,
        'shortages': shortages,
        'rest_cost': rest_cost,
        'sleep_cost': sleep_cost,
        'wellbeing': rest_cost + sleep_cost,
        'rest_violations': rest_viol,
        'rest_short_minutes': sum(v['required'] - v['actual'] for v in rest_viol),
        'full_rest': sum(1 for st in stats.values() if st['longest_rest'] >= cfg.sleep_target),
        'full_night_rest': sum(1 for st in stats.values() if st['night_rest'] >= cfg.sleep_target),
        'no_rest': sum(1 for st in stats.values() if st['longest_rest'] < cfg.sleep_min),
        'soldiers': stats,
        'load_max': max(loads) / unit,
        'load_min': min(loads) / unit,
        'load_gap': (max(loads) - min(loads)) / unit,
        'load_std': (sum((x - sum(loads) / len(loads)) ** 2 for x in loads) / len(loads)) ** 0.5 / unit,
        'hours_gap': (max(mins) - min(mins)) / 60,
        'hours_avg': sum(mins) / 60 / max(1, len(mins)),
        '_loads_raw': {sid: st['load'] for sid, st in stats.items()},
    }


def explain_shortages(soldiers, shifts, assign, metrics, cfg: Config) -> list:
    """לכל חוסר: מה האילוץ שמונע את האיוש"""
    busy = {s.sid: [] for s in soldiers}
    for sid, k, _ in assign:
        if sid in busy and not shifts[k].task.shared:
            busy[sid].append(shifts[k])
    out = []
    for item in metrics['shortages']:
        sh, role = shifts[item['shift']], item['role']
        qualified = [s for s in soldiers if _base_eligible(s, sh.task)
                     and (role is None or role in s.roles)]
        avail = [s for s in qualified if _available(s, sh)]
        need = dict(sh.task.groups)[role]
        if not qualified:
            reason = (f'אין אף חייל עם הסמכת {role}' if role
                      else 'אין אף חייל שמותר לשבץ למשימה (פטורים / תפקידים חסומים)')
        elif len(avail) < need:
            reason = (f'רק {len(avail)} חיילים מתאימים זמינים בשעות האלה (נדרשים {need}); '
                      f'{len(qualified) - len(avail)} בחסימה אישית')
        elif sh.task.shared:
            reason = f'{len(avail)} חיילים מתאימים, אך כולם כבר משובצים בתקן אחר של המשימה'
        else:
            reason = (f'כל {len(avail)} החיילים המתאימים והזמינים משובצים באותן שעות '
                      f'במשימות בלעדיות אחרות')
        out.append({
            'משימה': sh.task.name,
            'התחלה': fmt_clock(sh.start, cfg.day_start),
            'סיום': fmt_clock(sh.end, cfg.day_start),
            'הסמכה': role or '—',
            'חסרים': item['missing'],
            'סיבה': reason,
        })
    return out


# ══════════════════════════════════════════════════════════════════
# 7. הפותר — CP-SAT בשלבים
# ══════════════════════════════════════════════════════════════════
def _new_solver(cfg: Config, seconds: float):
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(0.5, seconds)
    solver.parameters.num_search_workers = cfg.workers
    solver.parameters.random_seed = cfg.seed
    solver.parameters.log_search_progress = False
    return solver


def _solve_with_patience(model, solver, patience: float):
    """מריץ את הפותר, ועוצר מוקדם אם עבר 'patience' שניות בלי שיפור בפתרון."""
    class _Progress(cp_model.CpSolverSolutionCallback):
        def __init__(self):
            super().__init__()
            self.best = None
            self.last = time.time()

        def on_solution_callback(self):
            v = self.ObjectiveValue()
            if self.best is None or v < self.best - 0.5:
                self.best, self.last = v, time.time()

    cb = _Progress()
    done = threading.Event()

    def watchdog():
        while not done.wait(0.2):
            if cb.best is not None and time.time() - cb.last > patience:
                try:
                    cb.StopSearch()
                except Exception:
                    pass
                return

    th = threading.Thread(target=watchdog, daemon=True)
    th.start()
    try:
        return solver.Solve(model, cb)
    finally:
        done.set()
        th.join(timeout=1)


_STATUS_HE = {'OPTIMAL': 'אופטימלי (מוכח)', 'FEASIBLE': 'פתרון טוב (לא הוכחה אופטימליות)',
              'INFEASIBLE': 'אין פתרון', 'UNKNOWN': 'לא נמצא פתרון בזמן שהוקצב'}


def _solve_cpsat(soldiers, shifts, elig, cfg: Config, directives, hint_assign,
                 use_must: bool, log: Callable):
    t_start = time.time()
    n = len(soldiers)
    sid_idx = {s.sid: i for i, s in enumerate(soldiers)}
    model = cp_model.CpModel()
    notes = []

    # ── משתני החלטה ──────────────────────────────────────────────
    y = {}   # (חייל, משמרת, קבוצת-תקן) → 0/1
    a = {}   # (חייל, משמרת) → 0/1   (H5: לכל היותר תקן אחד במשמרת)
    for si in range(n):
        for sh in shifts:
            gs = elig[si][sh.idx]
            if not gs:
                continue
            vs = []
            for gi in gs:
                v = model.NewBoolVar(f'y_{si}_{sh.idx}_{gi}')
                y[si, sh.idx, gi] = v
                vs.append(v)
            if len(vs) == 1:
                a[si, sh.idx] = vs[0]
            else:
                av = model.NewBoolVar(f'a_{si}_{sh.idx}')
                model.Add(sum(vs) == av)
                a[si, sh.idx] = av

    # ── H1: כיסוי, עם משתנה חוסר ─────────────────────────────────
    u, short_terms = {}, []
    for sh in shifts:
        for gi, (role, cnt) in enumerate(sh.task.groups):
            vs = [y[si, sh.idx, gi] for si in range(n) if (si, sh.idx, gi) in y]
            uv = model.NewIntVar(0, cnt, f'u_{sh.idx}_{gi}')
            model.Add(sum(vs) + uv == cnt)
            u[sh.idx, gi] = uv
            short_terms.append(uv * (sh.dur * sh.task.priority))

    # ── H4: לא שתי משימות בלעדיות באותו זמן ──────────────────────
    excl = [sh for sh in shifts if not sh.task.shared]
    excl_idx = {sh.idx for sh in excl}
    seen = set()
    for t in sorted({sh.start for sh in excl}):
        active = tuple(sh.idx for sh in excl if sh.start <= t < sh.end)
        if len(active) < 2 or active in seen:
            continue
        seen.add(active)
        for si in range(n):
            vs = [a[si, k] for k in active if (si, k) in a]
            if len(vs) > 1:
                model.AddAtMostOne(vs)

    # ── H6: הנחיות מפקד ──────────────────────────────────────────
    for d in (directives or []):
        if d.sid not in sid_idx:
            notes.append(f'הנחיית מפקד: מספר אישי {d.sid} לא נמצא ברשימת החיילים')
            continue
        si = sid_idx[d.sid]
        ds = _directive_shifts(d, shifts, cfg)
        if not ds:
            notes.append(f'הנחיית מפקד: אין משמרות למשימה {d.tid} בשעות שצוינו')
            continue
        if d.kind == 'must_not':
            for sh in ds:
                if (si, sh.idx) in a:
                    model.Add(a[si, sh.idx] == 0)
        elif use_must:
            vs = [a[si, sh.idx] for sh in ds if (si, sh.idx) in a]
            if vs:
                model.AddBoolOr(vs)
            else:
                notes.append(f'הנחיית "חובה" לא יושמה: {soldiers[si].name} לא יכול לבצע את '
                             f'{ds[0].task.name} (הסמכה, פטור או חסימה אישית)')

    # משתני העזר של שלבים 2 ו-3 נרשמים כאן, כדי שנוכל לתת לפותר נקודת פתיחה מלאה
    rest_vars = []    # (משתנה, חייל, משמרת1, משמרת2)
    sleep_vars = []   # (חייל, [(משתנה, התנגשויות, עלות)], משתנה "חלון חופשי", עלות חופשי)
    fair = {}         # משתני שלב ההוגנות

    def add_hints(assign_list):
        """נקודת פתיחה מלאה לפותר: ערך לכל משתנה במודל, מחושב מתוך שיבוץ נתון."""
        model.ClearHints()
        done = set()

        def hint(var, val):
            if var.Index() not in done:
                done.add(var.Index())
                model.AddHint(var, int(val))

        on, inside, per_group = set(), set(), {}
        for sid, k, role in assign_list:
            si = sid_idx.get(sid)
            if si is None:
                continue
            for gi, (r, _) in enumerate(shifts[k].task.groups):
                if r == role and (si, k, gi) in y:
                    on.add((si, k, gi))
                    inside.add((si, k))
                    per_group[k, gi] = per_group.get((k, gi), 0) + 1
                    break
        for key, var in y.items():
            hint(var, key in on)
        for key, var in a.items():
            hint(var, key in inside)
        for (k, gi), var in u.items():
            hint(var, max(0, shifts[k].task.groups[gi][1] - per_group.get((k, gi), 0)))
        for var, si, k1, k2 in rest_vars:
            hint(var, (si, k1) in inside and (si, k2) in inside)
        for si, zs, zf, free_cost in sleep_vars:
            mine = {k for (s2, k) in inside if s2 == si and k in excl_idx}
            pick, pick_cost = None, free_cost
            for wi, (_, conf, cost) in enumerate(zs):
                if cost < pick_cost and not (conf & mine):
                    pick, pick_cost = wi, cost
            for wi, (var, _, _) in enumerate(zs):
                hint(var, wi == pick)
            hint(zf, pick is None)
        if fair:
            vals = [sum(shifts[k].load // fair['g'] for (s2, k) in inside if s2 == si)
                    for si in range(n)]
            for si in range(n):
                hint(fair['loads'][si], vals[si])
                dev = abs(vals[si] - fair['avg'])
                hint(fair['devs'][si], dev)
                for j, ev in fair['extra'][si]:
                    hint(ev, max(0, dev - j * fair['step']))
            hint(fair['lmax'], max(vals))
            hint(fair['lmin'], min(vals))

    def extract(solver):
        return [(soldiers[si].sid, k, shifts[k].task.groups[gi][0])
                for (si, k, gi), var in y.items() if solver.Value(var)]

    def remaining():
        return cfg.time_limit - (time.time() - t_start)

    stages = []

    def run_stage(name, objective, budget):
        model.Minimize(objective)
        solver = _new_solver(cfg, budget)
        t0 = time.time()
        status = _solve_with_patience(model, solver, cfg.patience)
        sname = solver.StatusName(status)
        ok = status in (cp_model.OPTIMAL, cp_model.FEASIBLE)
        val = int(round(solver.ObjectiveValue())) if ok else None
        stages.append({'שלב': name, 'סטטוס': _STATUS_HE.get(sname, sname),
                       'ערך': val, 'שניות': round(time.time() - t0, 2), '_status': sname})
        log(f'{name}: {_STATUS_HE.get(sname, sname)}')
        return (solver if ok else None), val, sname

    # ══ שלב 1: כיסוי ═════════════════════════════════════════════
    add_hints(hint_assign)
    obj1 = sum(short_terms)
    solver, v1, sname = run_stage('1. כיסוי עמדות', obj1, max(2.0, cfg.time_limit * 0.3))
    if solver is None:
        return None, stages, notes, sname
    best = extract(solver)
    model.Add(obj1 <= v1)
    # נקודת הפתיחה לשלב הבא: הטוב מבין פתרון שלב 1 והשיבוץ הבסיסי (אם הכיסוי שלו זהה)
    m_best = evaluate(soldiers, shifts, best, cfg)
    m_hint = evaluate(soldiers, shifts, hint_assign, cfg)
    if (not m_hint['hard_violations'] and m_hint['short_weighted'] <= v1
            and m_hint['wellbeing'] < m_best['wellbeing']):
        best = list(hint_assign)

    # ══ שלב 2: רווחה — מנוחה אחרי משמרת + מנוחה רצופה (עדיף בלילה) ══
    well_terms, well_const = [], 0
    pairs = []
    for k1 in excl:
        if k1.task.rest <= 0:
            continue
        for k2 in excl:
            if k1.idx != k2.idx and k1.end <= k2.start < k1.end + k1.task.rest:
                uu = math.ceil((k1.end + k1.task.rest - k2.start) / 30)
                pairs.append((k1.idx, k2.idx, cfg.w_rest * uu * uu))
    for si in range(n):
        for k1, k2, cost in pairs:
            if (si, k1) in a and (si, k2) in a:
                v = model.NewBoolVar(f'r_{si}_{k1}_{k2}')
                model.AddBoolOr([a[si, k1].Not(), a[si, k2].Not(), v])
                well_terms.append(v * cost)
                rest_vars.append((v, si, k1, k2))

    night = night_intervals(cfg)
    windows = sleep_windows(cfg, _grid(excl or shifts), night)
    for si in range(n):
        mine = [sh for sh in excl if (si, sh.idx) in a]
        cand, free_cost = {}, cfg.w_sleep_none
        for wa, wb, cost in windows:
            conf = frozenset(sh.idx for sh in mine if _overlaps(wa, wb, sh.start, sh.end))
            if not conf:
                free_cost = min(free_cost, cost)
            elif conf not in cand or cost < cand[conf]:
                cand[conf] = cost
        kept = []
        for conf, cost in sorted(cand.items(), key=lambda kv: (kv[1], len(kv[0]))):
            if cost >= free_cost or any(kc <= conf for kc, _ in kept):
                continue          # נשלט ע"י חלון זול יותר עם פחות התנגשויות
            kept.append((conf, cost))
        if not kept:
            well_const += free_cost
            continue
        zs = []
        for wi, (conf, cost) in enumerate(kept):
            z = model.NewBoolVar(f'z_{si}_{wi}')
            for k in conf:
                model.AddImplication(z, a[si, k].Not())
            if cost:
                well_terms.append(z * cost)
            zs.append((z, conf, cost))
        zf = model.NewBoolVar(f'zf_{si}')
        if free_cost:
            well_terms.append(zf * free_cost)
        model.AddExactlyOne([z for z, _, _ in zs] + [zf])
        sleep_vars.append((si, zs, zf, free_cost))

    if not well_terms:
        stages.append({'שלב': '2. מנוחה ושינה', 'סטטוס': 'אין מה לשפר', 'ערך': well_const,
                       'שניות': 0.0, '_status': 'OPTIMAL'})
    elif remaining() > 1:
        obj2 = sum(well_terms) + well_const
        add_hints(best)
        solver, val, _ = run_stage('2. מנוחה ושינה', obj2, max(2.0, remaining() * 0.5))
        if solver is not None:
            best = extract(solver)
        # נועלים לפי הערך של השיבוץ הטוב ביותר שבידינו (מחושב באופן בלתי תלוי)
        v2 = evaluate(soldiers, shifts, best, cfg)['wellbeing']
        model.Add(obj2 <= v2 + int(cfg.wellbeing_slack * max(0, v2 - well_const)))

    # ══ שלב 3: הוגנות — עומס משוקלל ═══════════════════════════════
    if n >= 2 and remaining() > 1:
        g = reduce(math.gcd, [sh.load for sh in shifts if sh.load] or [1])
        top = max(sum(sh.load // g for sh in shifts), 1)
        loads = []
        for si in range(n):
            terms = [a[si, sh.idx] * (sh.load // g) for sh in shifts if (si, sh.idx) in a]
            lv = model.NewIntVar(0, top, f'L_{si}')
            model.Add(lv == (sum(terms) if terms else 0))
            loads.append(lv)
        lmax = model.NewIntVar(0, top, 'Lmax')
        lmin = model.NewIntVar(0, top, 'Lmin')
        model.AddMaxEquality(lmax, loads)
        model.AddMinEquality(lmin, loads)
        avg = int(round(sum(shifts[k].load // g for _, k, _ in best) / n))
        # קנס קמור על הסטייה מהממוצע (קירוב מקטעי לריבוע הסטייה, כלומר לשונות):
        # סטייה גדולה של חייל אחד יקרה יותר מכמה סטיות קטנות.
        step = max(1, avg // 4)
        fair_terms, devs, extra = [], [], []
        for si in range(n):
            dv = model.NewIntVar(0, top, f'dev_{si}')
            model.Add(dv >= loads[si] - avg)
            model.Add(dv >= avg - loads[si])
            devs.append(dv)
            fair_terms.append(dv)
            row = []
            for j in range(1, 8):
                ev = model.NewIntVar(0, top, f'e_{si}_{j}')
                model.Add(ev >= dv - j * step)
                fair_terms.append(ev)
                row.append((j, ev))
            extra.append(row)
        fair.update(g=g, avg=avg, step=step, loads=loads, devs=devs, extra=extra,
                    lmax=lmax, lmin=lmin)
        obj3 = sum(fair_terms) + (lmax - lmin)
        add_hints(best)
        solver, _, _ = run_stage('3. הוגנות', obj3, max(2.0, remaining()))
        if solver is not None:
            best = extract(solver)

    return best, stages, notes, 'OK'


def solve_schedule(soldiers, tasks, cfg: Optional[Config] = None, directives=None,
                   log: Optional[Callable] = None) -> dict:
    """
    נקודת הכניסה הראשית.
    מחזיר dict עם: shifts, assign (השבצ"ק הסופי), baseline (השיבוץ הבסיסי),
    metrics, baseline_metrics, stages, warnings, engine, seconds.
    """
    cfg = cfg or Config()
    log = log or (lambda msg: None)
    t0 = time.time()

    shifts, warnings = prepare(soldiers, tasks, cfg)
    elig = eligibility(soldiers, shifts)
    baseline = greedy_baseline(soldiers, shifts, elig, cfg, directives)
    log('שיבוץ בסיסי הושלם')

    final, stages, engine = baseline, [], 'בסיסי'
    if not shifts:
        warnings.append('לא נוצרו משמרות — בדקו את שעות הפעילות של המשימות.')
    elif not soldiers:
        warnings.append('אין חיילים ברשימה.')
    elif not HAS_ORTOOLS:
        warnings.append('הספרייה OR-Tools אינה מותקנת — הוחזר השיבוץ הבסיסי בלבד.')
    else:
        res, stages, notes, state = _solve_cpsat(
            soldiers, shifts, elig, cfg, directives, baseline, use_must=True, log=log)
        if res is None and state == 'INFEASIBLE' and any(
                d.kind == 'must' for d in (directives or [])):
            warnings.append('הנחיות ה"חובה" של המפקד סותרות זו את זו (או אילוץ קשיח אחר), '
                            'ולכן השיבוץ חושב בלעדיהן. הנחיות ה"אסור" נשמרו.')
            res, stages, notes, state = _solve_cpsat(
                soldiers, shifts, elig, cfg, directives, baseline, use_must=False, log=log)
        warnings += notes
        if res is not None:
            final, engine = res, 'CP-SAT'
        else:
            warnings.append('הפותר לא החזיר פתרון בזמן שהוקצב — הוחזר השיבוץ הבסיסי.')

    metrics = evaluate(soldiers, shifts, final, cfg)
    base_metrics = evaluate(soldiers, shifts, baseline, cfg)
    if metrics['hard_violations'] and not base_metrics['hard_violations']:
        warnings.append('נמצאה הפרה של אילוץ קשיח בפתרון — הוחזר השיבוץ הבסיסי.')
        final, metrics, engine = baseline, base_metrics, 'בסיסי'

    return {
        'shifts': shifts, 'assign': final, 'baseline': baseline,
        'metrics': metrics, 'baseline_metrics': base_metrics,
        'shortage_reasons': explain_shortages(soldiers, shifts, final, metrics, cfg),
        'stages': stages, 'warnings': warnings, 'engine': engine,
        'seconds': round(time.time() - t0, 2), 'cfg': cfg,
    }


# ══════════════════════════════════════════════════════════════════
# 8. טבלאות תוצאה
# ══════════════════════════════════════════════════════════════════
def assignment_rows(soldiers, shifts, assign, cfg: Config) -> list:
    """שורה לכל תקן בכל משמרת (כולל תקנים חסרים) — זה מה שנשמר בבסיס הנתונים."""
    name_of = {s.sid: s.name for s in soldiers}
    taken = {}
    for sid, k, role in assign:
        taken.setdefault((k, role), []).append(sid)
    rows = []
    for sh in sorted(shifts, key=lambda x: (x.task.tid, x.start)):
        for role, cnt in sh.task.groups:
            sids = sorted(taken.get((sh.idx, role), []), key=lambda x: name_of.get(x, ''))
            for i in range(cnt):
                sid = sids[i] if i < len(sids) else None
                rows.append({
                    'task_no': sh.task.tid, 'task_name': sh.task.name,
                    'personal_id': sid, 'soldier_name': name_of.get(sid) if sid else None,
                    'role': role, 'start_min': sh.start, 'end_min': sh.end,
                    'shared': int(sh.task.shared),
                })
    return rows


def task_table(rows: list, day_start: int) -> pd.DataFrame:
    out = []
    for r in rows:
        out.append({
            'משימה': r['task_name'],
            'התחלה': fmt_clock(r['start_min'], day_start),
            'סיום': fmt_clock(r['end_min'], day_start),
            'הסמכה נדרשת': r['role'] or '—',
            'חייל משובץ': r['soldier_name'] or '⚠️ חוסר',
            'סטטוס': '✅ מאויש' if r['personal_id'] else '❌ חוסר',
        })
    df = pd.DataFrame(out)
    df.index = range(1, len(df) + 1)
    return df


def soldier_grid(rows: list, roster: list, day_start: int) -> pd.DataFrame:
    """
    שורה לכל חייל, עמודה לכל משבצת זמן (שעה, או חצי שעה אם יש משמרות לא עגולות).
    roster: רשימת (מספר אישי, שם) של כל החיילים, גם אלה שלא שובצו.
    """
    vals = [60] + [r['start_min'] for r in rows] + [r['end_min'] for r in rows]
    g = reduce(math.gcd, vals)
    step = 60 if g % 60 == 0 else (30 if g % 30 == 0 else 15)
    cols = list(range(0, DAY, step))
    table = {str(sid): {c: [] for c in cols} for sid, _ in roster}
    missing = {c: [] for c in cols}
    for r in rows:
        target = table.get(str(r['personal_id'])) if r['personal_id'] else missing
        if target is None:
            continue
        for c in cols:
            if _overlaps(c, c + step, r['start_min'], r['end_min']):
                if r['task_name'] not in target[c]:
                    target[c].append(r['task_name'])
    labels = {c: fmt_clock(c, day_start) for c in cols}
    data = []
    for sid, name in roster:
        row = {'שם': name}
        row.update({labels[c]: ' + '.join(table[str(sid)][c]) or '—' for c in cols})
        data.append(row)
    if any(missing.values()):
        row = {'שם': '⚠️ חוסר כוח אדם'}
        row.update({labels[c]: ('❌ ' + ' + '.join(missing[c])) if missing[c] else '—' for c in cols})
        data.append(row)
    df = pd.DataFrame(data)
    df.index = range(1, len(df) + 1)
    return df


def summary_table(metrics: dict) -> pd.DataFrame:
    out = []
    for st in metrics['soldiers'].values():
        out.append({
            'שם': st['name'],
            'שעות במשימות': round(st['minutes'] / 60, 1),
            'מתוכן בלילה': round(st['night_minutes'] / 60, 1),
            'שעות כוננות': round(st['shared_minutes'] / 60, 1),
            'עומס משוקלל': round(st['load'] / 120, 1),
            'מנוחה רצופה (שעות)': round(st['longest_rest'] / 60, 1),
            'רצף מנוחה בלילה': round(st['night_rest'] / 60, 1),
        })
    df = pd.DataFrame(out)
    df.index = range(1, len(df) + 1)
    return df


def violations_table(metrics: dict, cfg: Config) -> pd.DataFrame:
    """חריגות מהיעדים הרכים — לאישור המפקד"""
    out = []
    for v in metrics['rest_violations']:
        out.append({
            'חייל': v['name'], 'סוג': 'מנוחה קצרה בין משמרות',
            'פירוט': f"אחרי {v['after']} ({fmt_clock(v['at'], cfg.day_start)}) ולפני {v['before']}: "
                     f"{fmt_dur(v['actual'])} שעות מנוחה במקום {fmt_dur(v['required'])}",
        })
    for st in metrics['soldiers'].values():
        if st['longest_rest'] < cfg.sleep_target:
            out.append({
                'חייל': st['name'], 'סוג': 'אין מנוחה רצופה מלאה',
                'פירוט': f"הרצף הארוך ביותר: {fmt_dur(st['longest_rest'])} שעות "
                         f"(יעד: {fmt_dur(cfg.sleep_target)})",
            })
    df = pd.DataFrame(out, columns=['חייל', 'סוג', 'פירוט'])
    df.index = range(1, len(df) + 1)
    return df


def comparison_table(base: dict, final: dict, n_soldiers: int) -> pd.DataFrame:
    """השוואת השיבוץ הבסיסי (מחקה שיבוץ ידני) מול השיבוץ הממוטב"""
    hours = lambda m: f'{m / 60:.1f}'
    one = lambda x: f'{x:.1f}'
    whole = lambda x: str(int(x))
    spec = [
        ('שעות-עמדה לא מאוישות', lambda m: hours(m['short_minutes'])),
        ('אחוז כיסוי', lambda m: one(m['coverage_pct']) + '%'),
        ('מקרים של מנוחה קצרה מהנדרש', lambda m: whole(len(m['rest_violations']))),
        ('סך שעות מנוחה שקוצצו', lambda m: hours(m['rest_short_minutes'])),
        (f'חיילים עם מנוחה רצופה מלאה (מתוך {n_soldiers})', lambda m: whole(m['full_rest'])),
        ('מתוכם מנוחה מלאה בלילה', lambda m: whole(m['full_night_rest'])),
        ('פער עומס משוקלל (העמוס ביותר מול הפנוי ביותר)', lambda m: one(m['load_gap'])),
        ('סטיית תקן של העומס המשוקלל', lambda m: one(m['load_std'])),
        ('פער שעות', lambda m: one(m['hours_gap'])),
    ]
    return pd.DataFrame([{'מדד': label, 'שיבוץ בסיסי': f(base), 'שיבוץ ממוטב': f(final)}
                         for label, f in spec])
