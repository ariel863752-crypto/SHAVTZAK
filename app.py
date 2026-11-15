"""
app.py — שבצ"ק: מערכת שיבוץ כוחות

הרצה מקומית:
    pip install -r requirements.txt
    streamlit run app.py

הגדרות שנקבעות מחוץ לקוד (ב-Streamlit Cloud: Settings → Secrets):
    ADMIN_PASSWORD = "..."      סיסמת המשתמש admin, שרואה את מסך הניהול
    DATABASE_URL   = "..."      שורת החיבור לבסיס נתונים חיצוני (Postgres).
                                בלי השורה הזו הנתונים נשמרים בקובץ מקומי, שנמחק באתחול השרת.

קבצים:  scheduler.py (האלגוריתם) · database.py (בסיס הנתונים) · test_scheduler.py (בדיקות)
"""
import datetime
import html as html_lib
import io
import os
import traceback

import pandas as pd
import streamlit as st
from streamlit import config as _st_config

import database as db
import scheduler as sc

# ── ערכת צבעים בהירה וקבועה, גם אם הדפדפן של המשתמש במצב כהה ─────────
INK, BOARD, RULE, OLIVE, KHAKI, ALERT = '#1C2620', '#F5F6F3', '#D8DDD4', '#3F5B34', '#ECEDE6', '#B93A2E'
try:
    _changed = False
    for _k, _v in {'theme.base': 'light', 'theme.primaryColor': OLIVE, 'theme.backgroundColor': BOARD,
                   'theme.secondaryBackgroundColor': KHAKI, 'theme.textColor': INK}.items():
        if _st_config.get_option(_k) != _v:
            _st_config.set_option(_k, _v)
            _changed = True
    if _changed:
        st.rerun()
except Exception:
    pass

st.set_page_config(page_title='שבצ"ק', page_icon='🗓️', layout='wide')

_VER = tuple(int(x) for x in st.__version__.split('.')[:2])
FULL = {'width': 'stretch'} if _VER >= (1, 52) else {'use_container_width': True}

# צבעי המשימות בלוח: סדר קבוע לפי מס"ד המשימה; משימה תשיעית ומעלה מקבלת אפור
TASK_COLORS = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948']
OTHER_COLOR = '#8a8a86'
EVENT_COLORS = {'כניסה': '#2a78d6', 'הרצת שיבוץ': '#eb6834', 'הרשמה': '#1baf7a'}
COOKIE = 'shavtzak_session'

# ══════════════════════════════════════════════════════════════════
# עיצוב
# ══════════════════════════════════════════════════════════════════
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Hebrew:wght@400;500;600;700&display=swap');
:root{--ink:#1C2620;--soft:#5C6A5F;--board:#F5F6F3;--paper:#FFFFFF;--rule:#D8DDD4;--olive:#3F5B34;
  --khaki:#ECEDE6;--alert:#B93A2E;--amber:#8A5A00;--night:#E9ECF1;
  --font:'IBM Plex Sans Hebrew','Heebo','Arial Hebrew',Arial,sans-serif}
html,body,.stApp{direction:rtl}
.stApp,.stApp p,.stApp li,.stApp label,.stApp input,.stApp textarea,.stApp button,.stApp select,
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp table,.stApp summary,
[data-testid="stMarkdownContainer"],[data-testid="stCaptionContainer"],[data-baseweb="select"] *{
  font-family:var(--font)!important}
[data-testid="stIconMaterial"],.material-symbols-rounded{font-family:'Material Symbols Rounded'!important;direction:ltr!important}
.stApp p,.stApp li,.stApp label,.stApp h1,.stApp h2,.stApp h3,.stApp h4,
[data-testid="stMarkdownContainer"],[data-testid="stCaptionContainer"],[data-testid="stAlert"],
[data-testid="stExpander"] summary,[data-testid="stFileUploader"]{direction:rtl;text-align:right}
.block-container{padding:1.2rem 2rem 4rem;max-width:1240px}
header[data-testid="stHeader"]{display:none!important}
[data-testid="stToolbar"],[data-testid="stDecoration"],#MainMenu,footer{display:none!important}
h1,h2,h3{letter-spacing:0;color:var(--ink)}
.stApp h3{font-size:1.15rem;font-weight:600;margin:0 0 .2rem}
/* טבלאות-נתונים ותרשימים מצוירים על canvas ואינם תומכים בכיוון RTL */
[data-testid="stDataFrame"],[data-testid="stDataFrame"] *,[data-testid="stDataEditor"],[data-testid="stDataEditor"] *,
[data-testid="stPlotlyChart"],[data-testid="stPlotlyChart"] *,.js-plotly-plot,.js-plotly-plot *{
  direction:ltr!important;text-align:left!important}

/* סרגל עליון */
.bar{display:flex;align-items:baseline;gap:14px;padding:2px 0 10px;border-bottom:2px solid var(--ink)}
.bar .mark{font-size:1.65rem;font-weight:700;color:var(--ink);line-height:1}
.bar .who{font-size:.92rem;color:var(--soft)}

/* לשוניות */
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:22px;border-bottom:1px solid var(--rule)}
[data-testid="stTabs"] [data-baseweb="tab"]{padding:10px 2px;height:auto;white-space:nowrap;background:transparent}
[data-testid="stTabs"] [data-baseweb="tab"] p{font-size:.98rem;font-weight:500;color:var(--soft)}
[data-testid="stTabs"] [aria-selected="true"] p{color:var(--ink)!important;font-weight:600}

/* כפתורים */
.stApp button[kind="primary"],[data-testid="stBaseButton-primary"],[data-testid="stFormSubmitButton"] button{
  background:var(--olive)!important;border:1px solid var(--olive)!important;border-radius:6px!important;min-height:44px}
.stApp button[kind="primary"] p,[data-testid="stBaseButton-primary"] p,[data-testid="stFormSubmitButton"] button p{
  color:#fff!important;font-weight:600;font-size:1rem}
.stApp button[kind="primary"]:hover,[data-testid="stBaseButton-primary"]:hover,
[data-testid="stFormSubmitButton"] button:hover{background:#2F4627!important}
.stApp button[kind="secondary"],[data-testid="stBaseButton-secondary"],[data-testid="stDownloadButton"] button{
  border-radius:6px!important;border:1px solid #B9C1B5!important;background:var(--paper)!important}
.stApp button:focus-visible{outline:2px solid var(--olive)!important;outline-offset:2px}
[data-testid="stExpander"] details{border:1px solid var(--rule);border-radius:6px;background:var(--paper)}
[data-testid="stFileUploader"] section{background:var(--paper);border:1px dashed #B9C1B5;border-radius:6px}

/* הודעות */
.note{border-inline-start:3px solid var(--olive);background:var(--paper);padding:10px 14px;margin:8px 0;
  font-size:.95rem;line-height:1.65;color:var(--ink);border-radius:0 4px 4px 0}
.note.warn{border-color:#C98A00;background:#FFFBF0}
.note.bad{border-color:var(--alert);background:#FFF6F4}
.lead{font-size:1.02rem;color:var(--ink);margin:6px 0 2px;line-height:1.6}
.hint{font-size:.88rem;color:var(--soft);line-height:1.6;margin:2px 0 8px}

/* שורת התוצאה */
.verdict{font-size:1.5rem;font-weight:700;color:var(--ink);margin:18px 0 10px;line-height:1.3}
.verdict.bad{color:var(--alert)}
.facts{display:flex;flex-wrap:wrap;border:1px solid var(--rule);border-radius:6px;background:var(--paper);margin:0 0 6px}
.fact{flex:1 1 170px;padding:12px 16px;border-inline-start:1px solid var(--rule)}
.fact:first-child{border-inline-start:none}
.fact b{display:block;font-size:1.45rem;font-weight:700;color:var(--ink);line-height:1.15;font-variant-numeric:tabular-nums}
.fact span{font-size:.86rem;color:var(--soft);line-height:1.45;display:block;margin-top:3px}
.fact.bad b{color:var(--alert)}

/* הלוח */
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;margin:10px 0 8px;font-size:.86rem;color:var(--ink)}
.legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-inline-end:6px;vertical-align:-1px}
.board-wrap{overflow-x:auto;border:1px solid var(--rule);border-radius:6px;background:var(--paper)}
.board{min-width:900px;font-variant-numeric:tabular-nums}
.b-row{display:flex;align-items:stretch;border-top:1px solid #EDF0EA}
.b-row:first-child{border-top:none}
.b-row:hover .b-track{background-color:#FAFBF8}
.b-name{flex:0 0 172px;padding:0 12px;display:flex;align-items:center;gap:6px;font-size:.9rem;font-weight:500;
  color:var(--ink);background:var(--paper);position:sticky;right:0;z-index:2;border-inline-end:1px solid var(--rule);
  white-space:nowrap;overflow:hidden}
.b-name .nm{overflow:hidden;text-overflow:ellipsis;min-width:0}
.b-name small{flex:0 0 auto;font-weight:400;font-size:.72rem;color:var(--soft);border:1px solid var(--rule);border-radius:10px;padding:0 6px}
.b-sum{flex:0 0 58px;display:flex;align-items:center;justify-content:center;font-size:.82rem;color:var(--soft);
  border-inline-start:1px solid var(--rule)}
.b-track{flex:1 1 auto;position:relative;height:36px;
  background-image:repeating-linear-gradient(to left,#E7EBE3 0,#E7EBE3 1px,transparent 1px,transparent calc(100% / 12))}
.b-head .b-track{height:26px;background-image:none}
.b-head{border-bottom:1px solid var(--rule);background:#FAFBF8}
.b-head .b-name,.b-head .b-sum{background:#FAFBF8;font-size:.78rem;font-weight:400;color:var(--soft)}
.tick{position:absolute;top:5px;font-size:.76rem;color:var(--soft);padding-inline-start:4px}
.night{position:absolute;top:0;bottom:0;background:var(--night);opacity:.75}
.blk{position:absolute;top:5px;height:26px;box-sizing:border-box;border-radius:3px;padding:0 7px;font-size:.8rem;
  line-height:26px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  border-inline-start:3px solid;margin-inline-end:2px}
.stand{position:absolute;bottom:2px;height:2px;border-radius:2px;opacity:.55}
.b-row.short .b-name{color:var(--alert)}
.b-row.short .blk{background:#FBE4E0;border-color:var(--alert);color:#7A1F16;font-weight:600}

/* טבלאות HTML */
table.t{width:100%;border-collapse:collapse;background:var(--paper);font-size:.92rem;direction:rtl;
  border:1px solid var(--rule);font-variant-numeric:tabular-nums}
table.t th{text-align:right;font-weight:600;color:var(--soft);font-size:.82rem;padding:8px 12px;
  border-bottom:1px solid var(--rule);background:#FAFBF8}
table.t td{text-align:right;padding:8px 12px;border-bottom:1px solid #EDF0EA;color:var(--ink);vertical-align:top;line-height:1.5}
table.t tr:last-child td{border-bottom:none}
table.t td.k{font-weight:600;white-space:nowrap}
table.t td.time{white-space:nowrap;color:var(--soft)}
table.t td.better{font-weight:700}
table.t .miss{color:var(--alert);font-weight:600}
.t-wrap{overflow-x:auto}
.t-wrap.narrow{max-width:820px}

/* עומס */
.load{display:grid;grid-template-columns:150px 1fr 150px;gap:4px 12px;align-items:center;font-size:.9rem;
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;padding:14px 16px;font-variant-numeric:tabular-nums}
.load .bar-bg{position:relative;height:14px;background:#EFF1EC;border-radius:2px}
.load .bar-fg{position:absolute;right:0;top:0;bottom:0;background:var(--olive);border-radius:2px 0 0 2px}
.load .avg{position:absolute;top:-3px;bottom:-3px;width:2px;background:var(--ink)}
.load .v{color:var(--soft);font-size:.84rem;white-space:nowrap}
.load .v b{color:var(--ink);font-weight:500}

/* מסך כניסה */
.hello{margin:9vh 0 18px}
.hello h1{font-size:2.3rem;font-weight:700;margin:0 0 8px;color:var(--ink)}
.hello p{font-size:1.05rem;color:var(--soft);line-height:1.7;margin:0}
@media (max-width:640px){.block-container{padding:1rem .9rem 3rem}.verdict{font-size:1.25rem}
  .b-name{flex-basis:118px;font-size:.82rem}.b-name small{display:none}.load{grid-template-columns:92px 1fr 118px;padding:10px}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def html(body: str):
    """HTML גולמי. בגרסאות ישנות של Streamlit נופלים חזרה ל-markdown."""
    if hasattr(st, 'html'):
        st.html(body)
    else:
        st.markdown(body, unsafe_allow_html=True)


def esc(v) -> str:
    return html_lib.escape('' if v is None else str(v), quote=True)


def note(text: str, kind: str = ''):
    html(f'<div class="note {kind}">{text}</div>')


def run_js(code: str):
    try:
        st.html(f'<script>{code}</script>', unsafe_allow_javascript=True)
    except Exception:
        try:
            import streamlit.components.v1 as components
            components.html(f'<script>{code.replace("document.cookie", "window.parent.document.cookie")}</script>',
                            height=0)
        except Exception:
            pass


def span(a: int, b: int, day_start: int) -> str:
    """טווח שעות שמוצג תמיד משמאל לימין (06:00–10:00), גם בתוך טקסט עברי"""
    return f'<span dir="ltr">{sc.fmt_clock(a, day_start)}–{sc.fmt_clock(b, day_start)}</span>'


def html_table(df: pd.DataFrame, key_col: int = 0, bold_col=None, narrow=False, code_key=False) -> str:
    head = ''.join(f'<th>{esc(c)}</th>' for c in df.columns)
    body = ''
    for _, r in df.iterrows():
        cells = ''
        for i, c in enumerate(df.columns):
            cls = 'k' if i == key_col else ('better' if c == bold_col else '')
            text = esc(r[c])
            if code_key and i == key_col:       # דוגמת קלט: מוצגת משמאל לימין, כמו שמקלידים אותה
                text = f'<span dir="ltr">{text}</span>'
            cells += f'<td class="{cls}">{text}</td>'
        body += f'<tr>{cells}</tr>'
    return f'<div class="t-wrap{" narrow" if narrow else ""}"><table class="t"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


# ══════════════════════════════════════════════════════════════════
# אתחול: בסיס נתונים ומשתמש מנהל
# ══════════════════════════════════════════════════════════════════
def _secret(name: str, env: str):
    val = os.environ.get(env)
    if val:
        return val
    try:
        val = st.secrets.get(name)
        return str(val) if val is not None else None
    except Exception:
        return None


@st.cache_resource
def _boot(database_url, admin_pw):
    # הפרמטרים הם חלק מהמפתח: שינוי ב-Secrets מפעיל את האתחול מחדש
    db.configure(database_url)
    db.init_db()
    db.ensure_admin(admin_pw)
    return True


try:
    _boot(_secret('DATABASE_URL', 'DATABASE_URL'), _secret('ADMIN_PASSWORD', 'SHAVTZAK_ADMIN_PASSWORD'))
except Exception as e:
    st.error('לא הצלחתי להתחבר לבסיס הנתונים. בדקו ששורת החיבור DATABASE_URL ב-Secrets הועתקה במלואה, '
             'בתוך מירכאות.')
    st.code(str(e))
    st.stop()

ss = st.session_state
ss.setdefault('user', None)
ss.setdefault('view', None)
ss.setdefault('ed_ver', 0)


# ══════════════════════════════════════════════════════════════════
# תבניות ונתוני דוגמה
# ══════════════════════════════════════════════════════════════════
S_COLS = {'personal_id': 'מספר אישי', 'full_name': 'שם מלא', 'exemptions': 'פטורים',
          'roles': 'הסמכות', 'blocked_hours': 'שעות חסימה'}
T_COLS = {'task_no': 'מס"ד משימה', 'name': 'שם המשימה', 'required': 'סד"כ נדרש למשימה',
          'shift_hours': 'משך משמרת', 'rest_hours': 'שעות מנוחה בין משימות',
          'shared': 'אישור חפיפה בין משימות', 'active_hours': 'שעות פעילות',
          'required_roles': 'הסמכה נדרשת', 'intensity': 'דירוג עצימות משימה (1-10)',
          'blocked_roles': 'תפקידים חסומים', 'priority': 'קריטיות (1-10)'}
D_COLS = ['חייל', 'משימה', 'הנחיה', 'שעות (לא חובה)']


def demo_rows():
    names = ['נועם לוי', 'איתי כהן', 'יובל מזרחי', 'עומר פרץ', 'דניאל ביטון', 'רועי אברהם',
             'אורי פרידמן', 'עידו דהן', 'תומר אזולאי', 'גיא שפירא', 'אלון חדד', 'יונתן גבאי',
             'ניר סויסה', 'שחר אוחיון', 'עמית קליין', 'מתן רוזן', 'ליאור נחום', 'אביב שלום',
             'בן וקנין', 'אסף מלכה', 'טל גולן', 'דור ברק', 'יאיר סגל', 'רן אשכנזי']
    roles = {0: 'מפקד', 6: 'מפקד', 12: 'מפקד', 18: 'מפקד',
             1: 'נהג', 7: 'נהג', 13: 'נהג', 19: 'נהג', 4: 'נהג,חובש', 10: 'חובש'}
    blocked = {3: '08:30-12:00', 9: '14:00-18:00', 15: '06:00-10:30'}
    exempt = {5: '104', 20: '105'}
    s_rows = [{'personal_id': str(7001 + i), 'full_name': n, 'exemptions': exempt.get(i, ''),
               'roles': roles.get(i, ''), 'blocked_hours': blocked.get(i, '')}
              for i, n in enumerate(names)]
    t = lambda no, name, req, shift, rest, shared, hours, rr, inten, br, pr: {
        'task_no': no, 'name': name, 'required': req, 'shift_hours': shift, 'rest_hours': rest,
        'shared': shared, 'active_hours': hours, 'required_roles': rr, 'intensity': inten,
        'blocked_roles': br, 'priority': pr}
    t_rows = [
        t(101, 'ש"ג', 2, 4, 8, False, 'all', '', 6, '', 9),
        t(102, 'סיור', 3, 6, 8, False, 'all', 'מפקד,נהג', 8, '', 8),
        t(103, 'תצפית', 1, 3, 6, False, 'all', '', 5, '', 7),
        t(104, 'מטבח', 2, 7, 4, False, '06:00-20:00', '', 4, 'מפקד', 3),
        t(105, 'שמירה היקפית', 1, 4, 8, False, '22-6', '', 7, '', 8),
        t(106, 'כוננות', 4, 24, 0, True, 'all', 'מפקד,נהג,חובש', 1, '', 6),
    ]
    return s_rows, t_rows


def _filled(v) -> bool:
    """האם בתא יש ערך (ולא None / NaN / מחרוזת ריקה)"""
    if v is None or v is pd.NA:
        return False
    if isinstance(v, float) and v != v:
        return False
    return str(v).strip().lower() not in ('', 'none', 'nan', '<na>')


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """טבלה שנערכה באתר → ערכים רגילים, עם None בתאים ריקים"""
    out = df.astype(object)
    return out.apply(lambda col: col.map(lambda v: v if _filled(v) else None))


def _col(kind, label=None, **kw):
    """הגדרת עמודה בטבלת עריכה. בגרסאות ישנות של Streamlit אין קיבוע עמודות, אז מוותרים עליו."""
    try:
        return kind(label, **kw)
    except TypeError:
        kw.pop('pinned', None)
        return kind(label, **kw)


def _is_true(v) -> bool:
    return v is True or str(v).strip().lower() in ('true', '1', '1.0', 'כן')


def rows_to_frames(s_rows, t_rows):
    """שורות מבסיס הנתונים → טבלאות בעמודות העבריות של התבנית"""
    s_df = pd.DataFrame([{S_COLS[k]: r.get(k, '') for k in S_COLS} for r in s_rows],
                        columns=list(S_COLS.values())).astype(str)
    t_df = pd.DataFrame([{T_COLS[k]: r.get(k, '') for k in T_COLS} for r in t_rows],
                        columns=list(T_COLS.values()))
    t_df[T_COLS['shared']] = t_df[T_COLS['shared']].map(_is_true).astype(bool)
    for k in ('task_no', 'required', 'intensity', 'priority'):
        t_df[T_COLS[k]] = pd.to_numeric(t_df[T_COLS[k]], errors='coerce').astype('Int64')
    for k in ('shift_hours', 'rest_hours'):
        t_df[T_COLS[k]] = pd.to_numeric(t_df[T_COLS[k]], errors='coerce').astype(float)
    for k in ('name', 'active_hours', 'required_roles', 'blocked_roles'):
        t_df[T_COLS[k]] = t_df[T_COLS[k]].fillna('').astype(str)
    return s_df, t_df


def to_excel(sheets: dict) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        for name, df in sheets.items():
            df.to_excel(w, index=False, sheet_name=name)
            wb, ws = w.book, w.sheets[name]
            ws.right_to_left()
            ws.freeze_panes(1, 1)
            head = wb.add_format({'bold': True, 'fg_color': OLIVE, 'font_color': 'white',
                                  'border': 1, 'align': 'right', 'valign': 'vcenter'})
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, head)
                longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)])
                ws.set_column(ci, ci, min(longest + 3, 45))
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════
# כניסה, הרשמה והתחברות שנשמרת
# ══════════════════════════════════════════════════════════════════
def sign_in(user: dict):
    ss.user = user
    try:
        ss.cookie_job = ('set', db.create_session(user['id']))
    except Exception:
        ss.cookie_job = None
    st.rerun()


def sign_out():
    try:
        db.delete_session(ss.get('token') or _cookie_token())
    except Exception:
        pass
    for k in ('user', 'view', 'roster', 'runs', 'run_cache', 'token', 'admin_ov'):
        ss[k] = None
    ss.cookie_job = ('clear', '')
    st.rerun()


def _cookie_token():
    try:
        return st.context.cookies.get(COOKIE)
    except Exception:
        return None


def flush_cookie_job():
    """כותב או מוחק את עוגיית ההתחברות בדפדפן (פעם אחת, אחרי כניסה או יציאה)"""
    job = ss.get('cookie_job')
    if not job:
        return
    ss.cookie_job = None
    kind, token = job
    if kind == 'set':
        ss.token = token
        age = db.SESSION_DAYS * 86400
        run_js(f"document.cookie='{COOKIE}={token}; path=/; max-age={age}; SameSite=Lax'"
               "+(location.protocol==='https:'?'; Secure':'');")
    else:
        run_js(f"document.cookie='{COOKIE}=; path=/; max-age=0; SameSite=Lax';")


if ss.user is None and not ss.get('cookie_checked'):
    ss.cookie_checked = True
    _tok = _cookie_token()
    if _tok:
        try:
            _u = db.user_from_session(_tok)
        except Exception:
            _u = None
        if _u:
            ss.user, ss.token = _u, _tok


def auth_screen():
    _, mid, _ = st.columns([1, 1.25, 1])
    with mid:
        html('<div class="hello"><h1>שבצ"ק</h1>'
             '<p>שבצ"ק יומי מוכן בתוך שניות: כל עמדה מאוישת, כל חייל מקבל מנוחה, והעומס מתחלק בהגינות. '
             'את ההחלטות הקטנות משאירים למפקד.</p></div>')
        t_in, t_new = st.tabs(['כניסה', 'משתמש חדש'])
        with t_in:
            with st.form('login', border=False):
                u = st.text_input('שם משתמש')
                p = st.text_input('סיסמה', type='password')
                if st.form_submit_button('כניסה', **FULL):
                    user = db.authenticate(u, p)
                    if user:
                        sign_in(user)
                    st.error('שם המשתמש או הסיסמה לא נכונים.')
        with t_new:
            with st.form('register', border=False):
                u = st.text_input('שם משתמש', help='3 תווים לפחות')
                unit = st.text_input('יחידה או מסגרת', help='לא חובה')
                p = st.text_input('סיסמה', type='password', help='6 תווים לפחות')
                p2 = st.text_input('הסיסמה שוב', type='password')
                if st.form_submit_button('פתיחת משתמש', **FULL):
                    if p != p2:
                        st.error('שתי הסיסמאות לא זהות.')
                    else:
                        user, err = db.create_user(u, p, unit)
                        if err:
                            st.error(err)
                        else:
                            sign_in(user)
        html('<p class="hint">הנתונים שתזינו נשמרים במשתמש שלכם בלבד. '
             'עדיף לעבוד עם שמות ומספרים שאינם מזהים חיילים אמיתיים.</p>')
    flush_cookie_job()


if ss.user is None:
    auth_screen()
    st.stop()

USER = ss.user


# ══════════════════════════════════════════════════════════════════
# נתונים שמורים (נשמרים גם בזיכרון של הגלישה, כדי לא לפנות לבסיס הנתונים בכל לחיצה)
# ══════════════════════════════════════════════════════════════════
def get_roster():
    if ss.get('roster') is None:
        ss.roster = db.load_roster(USER['id'])
    return ss.roster


def set_roster(s_rows, t_rows, event: str, details: str = ''):
    db.save_roster(USER['id'], s_rows, t_rows)
    db.log_event(USER['id'], event, details)
    ss.roster = None
    ss.ed_ver += 1


def get_runs():
    if ss.get('runs') is None:
        ss.runs = db.list_runs(USER['id'])
    return ss.runs


def get_run(run_id: int):
    cache = ss.get('run_cache') or {}
    if run_id not in cache:
        cache[run_id] = db.load_run(run_id, USER['id'])
        ss.run_cache = cache
    return cache[run_id]


# ══════════════════════════════════════════════════════════════════
# תצוגות השבצ"ק
# ══════════════════════════════════════════════════════════════════
def _tint(hex_color: str, amount: float = 0.84) -> str:
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    mix = lambda c: round(c + (255 - c) * amount)
    return f'#{mix(r):02x}{mix(g):02x}{mix(b):02x}'


def task_colors(rows) -> dict:
    by_number = sorted({(r['task_no'], r['task_name']) for r in rows})
    return {name: (TASK_COLORS[i] if i < len(TASK_COLORS) else OTHER_COLOR)
            for i, (_, name) in enumerate(by_number)}


def board_html(rows, roster, day_start: int) -> str:
    """הלוח: שורה לכל חייל, ציר זמן מימין לשמאל, מקטע לכל משמרת."""
    colors = task_colors(rows)
    pct = lambda m: f'{m / sc.DAY * 100:.3f}%'
    clock = lambda m: sc.fmt_clock(m, day_start)
    night = sc.night_intervals(sc.Config(day_start=day_start))
    night_html = ''.join(f'<i class="night" style="right:{pct(a)};width:{pct(b - a)}"></i>' for a, b in night)

    mine = {str(sid): [] for sid, _ in roster}
    missing = {}
    for r in rows:
        if r['personal_id']:
            if str(r['personal_id']) in mine:
                mine[str(r['personal_id'])].append(r)
        else:
            missing.setdefault(r['task_name'], []).append(r)

    ticks = ''.join(f'<span class="tick" style="right:{pct(m)}">{clock(m)}</span>'
                    for m in range(0, sc.DAY, 120))
    out = ['<div class="b-row b-head"><div class="b-name">חייל</div>'
           f'<div class="b-track">{ticks}</div><div class="b-sum">שעות</div></div>']

    def block(r, label):
        c = colors.get(r['task_name'], OTHER_COLOR)
        tip = f"{r['task_name']}{' (' + r['role'] + ')' if r['role'] else ''}, {clock(r['start_min'])}-{clock(r['end_min'])}"
        return (f'<div class="blk" title="{esc(tip)}" style="right:{pct(r["start_min"])};'
                f'width:calc({pct(r["end_min"] - r["start_min"])} - 2px);background:{_tint(c)};'
                f'border-color:{c}">{esc(label)}</div>')

    for sid, name in roster:
        items = mine[str(sid)]
        excl = [r for r in items if not r['shared']]
        shared = [r for r in items if r['shared']]
        tags = ''.join(f'<small>{esc(n)}</small>' for n in dict.fromkeys(r['task_name'] for r in shared))
        track = night_html + ''.join(block(r, r['task_name']) for r in excl)
        track += ''.join(
            f'<i class="stand" title="{esc(r["task_name"])}, {clock(r["start_min"])}-{clock(r["end_min"])}" '
            f'style="right:{pct(r["start_min"])};width:{pct(r["end_min"] - r["start_min"])};'
            f'background:{colors.get(r["task_name"], OTHER_COLOR)}"></i>' for r in shared)
        hours = sum(r['end_min'] - r['start_min'] for r in excl) / 60
        out.append(f'<div class="b-row"><div class="b-name"><span class="nm">{esc(name)}</span>{tags}</div>'
                   f'<div class="b-track">{track}</div><div class="b-sum">{hours:g}</div></div>')

    for task_name, items in missing.items():
        merged = {}
        for r in items:
            merged.setdefault((r['start_min'], r['end_min']), []).append(r)
        track = night_html
        for (a, b), group in sorted(merged.items()):
            label = 'לא אויש' + (f' ×{len(group)}' if len(group) > 1 else '')
            track += (f'<div class="blk" title="{esc(task_name)}, {clock(a)}-{clock(b)}" '
                      f'style="right:{pct(a)};width:calc({pct(b - a)} - 2px)">{label}</div>')
        out.append(f'<div class="b-row short"><div class="b-name">חסר: {esc(task_name)}</div>'
                   f'<div class="b-track">{track}</div><div class="b-sum"></div></div>')

    legend = ''.join(f'<span><i style="background:{c}"></i>{esc(n)}</span>' for n, c in colors.items())
    legend += '<span><i style="background:#DDE2EA"></i>לילה</span>'
    return (f'<div class="legend">{legend}</div>'
            f'<div class="board-wrap"><div class="board">{"".join(out)}</div></div>')


def by_task_html(rows, day_start: int) -> str:
    """לכל משימה: מי נמצא בה בכל משמרת"""
    clock = lambda m: sc.fmt_clock(m, day_start)
    tasks = {}
    for r in sorted(rows, key=lambda x: (x['task_no'], x['start_min'])):
        tasks.setdefault(r['task_name'], {}).setdefault((r['start_min'], r['end_min']), []).append(r)
    body = ''
    for name, shifts in tasks.items():
        first = True
        for (a, b), group in shifts.items():
            people = []
            for r in group:
                if r['personal_id']:
                    people.append(esc(r['soldier_name']) + (f' ({esc(r["role"])})' if r['role'] else ''))
                else:
                    people.append('<span class="miss">לא אויש' + (f' ({esc(r["role"])})' if r['role'] else '') + '</span>')
            lead = f'<td class="k" rowspan="{len(shifts)}">{esc(name)}</td>' if first else ''
            body += f'<tr>{lead}<td class="time">{span(a, b, day_start)}</td><td>{", ".join(people)}</td></tr>'
            first = False
    return ('<div class="t-wrap narrow"><table class="t"><thead><tr><th>משימה</th><th>שעות</th><th>משובצים</th></tr></thead>'
            f'<tbody>{body}</tbody></table></div>')


def rest_issues_html(violations: pd.DataFrame) -> str:
    """שורה לכל חייל: איפה קוצרה לו מנוחה בין משמרות, ומה רצף המנוחה הארוך ביותר שלו"""
    people = {}
    for _, r in violations.iterrows():
        entry = people.setdefault(r['חייל'], {'between': [], 'block': ''})
        text = esc(r['פירוט']).replace(': 0:00 שעות מנוחה במקום ', ': בלי מנוחה, במקום ')
        if r['סוג'].startswith('אין מנוחה רצופה'):
            text = text.replace('הרצף הארוך ביותר: ', '')
            entry['block'] = 'אין מנוחה בכלל' if text.startswith('0:00') else text
        else:
            entry['between'].append(text)
    body = ''
    for name, e in people.items():
        between = '<br>'.join(e['between']) if e['between'] else 'מלאה'
        body += (f'<tr><td class="k">{esc(name)}</td><td>{between}</td>'
                 f'<td>{e["block"] or "מלא"}</td></tr>')
    return ('<div class="t-wrap"><table class="t"><thead><tr><th>חייל</th><th>מנוחה בין משמרות</th>'
            f'<th>רצף המנוחה הארוך ביותר</th></tr></thead><tbody>{body}</tbody></table></div>')


def load_html(summary: pd.DataFrame) -> str:
    df = summary.sort_values('עומס משוקלל', ascending=False)
    top = max(float(df['עומס משוקלל'].max()), 1.0)
    avg = float(df['עומס משוקלל'].mean())
    cells = ''
    for _, r in df.iterrows():
        w = float(r['עומס משוקלל']) / top * 100
        extra = f", מהן {r['מתוכן בלילה']:g} בלילה" if r['מתוכן בלילה'] else ''
        cells += (f'<div>{esc(r["שם"])}</div>'
                  f'<div class="bar-bg" title="{r["שעות במשימות"]:g} שעות{extra}">'
                  f'<div class="bar-fg" style="width:{w:.1f}%"></div>'
                  f'<div class="avg" style="right:{avg / top * 100:.1f}%"></div></div>'
                  f'<div class="v"><b>{r["שעות במשימות"]:g} שעות</b>, עומס {r["עומס משוקלל"]:g}</div>')
    return f'<div class="load">{cells}</div>'


def schedule_excel(rows, roster, day_start, extra: dict = None) -> bytes:
    sheets = {'לפי משימה': sc.task_table(rows, day_start), 'לפי חייל': sc.soldier_grid(rows, roster, day_start)}
    for name, df in (extra or {}).items():
        if df is not None and not df.empty:
            sheets[name] = df
    return to_excel(sheets)


# ══════════════════════════════════════════════════════════════════
# הרצת שיבוץ
# ══════════════════════════════════════════════════════════════════
def run_schedule(s_rows, t_rows, d_rows, cfg: sc.Config):
    soldiers, tasks, directives, warns = sc.build_objects(s_rows, t_rows, d_rows)
    res = sc.solve_schedule(soldiers, tasks, cfg, directives)
    m = res['metrics']
    rows = sc.assignment_rows(soldiers, res['shifts'], res['assign'], cfg)
    roster = [(s.sid, s.name) for s in soldiers]
    compare = sc.comparison_table(res['baseline_metrics'], m, len(soldiers))
    label = 'שיבוץ ' + datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
    run_id = db.save_run(USER['id'], label, cfg.day_start, res['engine'], len(tasks), m,
                         res['seconds'], roster, compare.to_dict('records'), rows)
    db.log_event(USER['id'], 'run',
                 f"{len(soldiers)} חיילים, {len(tasks)} משימות, כיסוי {m['coverage_pct']:.0f}%")
    ss.runs = None
    return {
        'run_id': run_id, 'rows': rows, 'roster': roster, 'day_start': cfg.day_start,
        'sleep_target': cfg.sleep_target, 'engine': res['engine'], 'seconds': res['seconds'],
        'warnings': warns + res['warnings'],
        'stages': [{k: v for k, v in s.items() if not k.startswith('_')} for s in res['stages']],
        'compare': compare, 'summary': sc.summary_table(m),
        'violations': sc.violations_table(m, cfg),
        'shortages': pd.DataFrame(res['shortage_reasons']),
        'hard': m['hard_violations'],
        'm': {k: m[k] for k in ('short_minutes', 'coverage_pct', 'full_rest', 'full_night_rest',
                                'load_gap', 'hours_gap', 'slots_total', 'short_slots')},
        'n_rest_viol': len(m['rest_violations']),
    }


def render_result(v):
    m, n = v['m'], len(v['roster'])
    if v['hard']:
        note('נמצאה סתירה לאילוץ קשיח. אל תשתמשו בשבצ"ק הזה ודווחו לנו:<br>' + '<br>'.join(map(esc, v['hard'])), 'bad')
    for w in v['warnings']:
        note(esc(w), 'warn')

    short = m['short_minutes'] > 0
    if short:
        left = 'תקן אחד נשאר ריק' if m['short_slots'] == 1 else f'{m["short_slots"]} תקנים נשארו ריקים'
        html(f'<div class="verdict bad">השבצ"ק מוכן, אבל {left}</div>')
    elif v['n_rest_viol']:
        html(f'<div class="verdict">השבצ"ק מוכן: כל העמדות מאוישות, ב-{v["n_rest_viol"]} מקרים קוצרה מנוחה</div>')
    else:
        html('<div class="verdict">השבצ"ק מוכן: כל העמדות מאוישות</div>')

    rest_bad = v['n_rest_viol'] > 0
    html('<div class="facts">'
         f'<div class="fact {"bad" if short else ""}"><b>{m["slots_total"] - m["short_slots"]} מתוך {m["slots_total"]}</b>'
         '<span>תקנים מאוישים</span></div>'
         f'<div class="fact"><b>{m["full_rest"]} מתוך {n}</b>'
         f'<span>חיילים עם {sc.fmt_dur(v["sleep_target"]).split(":")[0]} שעות מנוחה רצופות, {m["full_night_rest"]} מהם בלילה</span></div>'
         f'<div class="fact {"bad" if rest_bad else ""}"><b>{v["n_rest_viol"]}</b>'
         '<span>מקרים של מנוחה קצרה מהנדרש בין משמרות</span></div>'
         f'<div class="fact"><b>{m["hours_gap"]:g} שעות</b>'
         '<span>ההפרש בין החייל העמוס ביותר לפנוי ביותר</span></div>'
         '</div>')

    n_issues = v['violations']['חייל'].nunique() + len(v['shortages'])
    t_board, t_task, t_issues, t_load, t_compare = st.tabs(
        ['לוח', 'לפי משימה', f'לאישור המפקד ({n_issues})' if n_issues else 'לאישור המפקד', 'עומס', 'לפני ואחרי'])
    with t_board:
        html(board_html(v['rows'], v['roster'], v['day_start']))
        html('<p class="hint">כל שורה היא חייל, והזמן מתקדם מימין לשמאל. '
             'העמידו את העכבר על מקטע כדי לראות שעות מדויקות. פס דק בתחתית השורה מסמן כוננות.</p>')
    with t_task:
        html(by_task_html(v['rows'], v['day_start']))
    with t_issues:
        if not n_issues:
            note('אין מה לאשר: כל העמדות מאוישות, וכל חייל מקבל את המנוחה שהוגדרה.')
        if not v['shortages'].empty:
            st.markdown('### תקנים שלא אוישו')
            html('<p class="hint">אי אפשר לאייש אותם בלי לשבור אילוץ קשיח. הסיבה מופיעה ליד כל אחד.</p>')
            html(html_table(v['shortages']))
        if not v['violations'].empty:
            st.markdown('### מנוחה שקוצרה')
            html('<p class="hint">המקומות שבהם המערכת ויתרה על מנוחה כדי שלא תישאר עמדה ריקה.</p>')
            html(rest_issues_html(v['violations']))
    with t_load:
        html('<p class="hint">העומס משקלל שעות לפי עצימות המשימה, ושעת לילה נספרת כשעה וחצי. '
             'הקו הכהה מסמן את הממוצע.</p>')
        html(load_html(v['summary']))
        with st.expander('הטבלה המלאה'):
            html(html_table(v['summary']))
    with t_compare:
        html('<p class="hint">"שיבוץ בסיסי" מחקה את הדרך הידנית: עוברים על המשמרות לפי הסדר ומשבצים בכל פעם '
             'את מי שפנוי ועבד הכי פחות. "שיבוץ ממוטב" הוא מה שהמערכת הפיקה.</p>')
        html(html_table(v['compare'], bold_col='שיבוץ ממוטב'))
        with st.expander('שלבי החישוב'):
            html(f'<p class="hint">החישוב נמשך {v["seconds"]:g} שניות, במנוע {esc(v["engine"])}. '
                 'כל שלב ננעל לפני הבא, כך ששלב מאוחר לא פוגע בשלב שקדם לו.</p>')
            if v['stages']:
                html(html_table(pd.DataFrame(v['stages'])))

    data = schedule_excel(v['rows'], v['roster'], v['day_start'],
                          {'סיכום לחייל': v['summary'], 'לאישור המפקד': v['violations'],
                           'חוסרים': v['shortages'], 'השוואה': v['compare']})
    if st.download_button('הורדה ל-Excel', data=data, file_name=f'shavtzak_{v["run_id"]}.xlsx',
                          key=f'dl_run{v["run_id"]}'):
        db.log_event(USER['id'], 'download', f'run {v["run_id"]}')


# ══════════════════════════════════════════════════════════════════
# מסך ראשי
# ══════════════════════════════════════════════════════════════════
bar_l, bar_r = st.columns([7, 1])
with bar_l:
    who = esc(USER['username']) + (f', {esc(USER["unit_name"])}' if USER['unit_name'] else '')
    html(f'<div class="bar"><span class="mark">שבצ"ק</span><span class="who">{who}</span></div>')
with bar_r:
    if st.button('יציאה', **FULL):
        sign_out()

tab_names = ['שבצ"ק', 'נתונים', 'היסטוריה', 'מדריך'] + (['ניהול'] if USER['is_admin'] else [])
tabs = st.tabs(tab_names)
tab_run, tab_data, tab_history, tab_guide = tabs[:4]
s_rows, t_rows = get_roster()

# ── שבצ"ק ────────────────────────────────────────────────────────
with tab_run:
    if not s_rows or not t_rows:
        html('<p class="lead">כדי ליצור שבצ"ק צריך רשימת חיילים ורשימת משימות.</p>'
             '<p class="hint">אפשר להזין אותן בלשונית "נתונים", או לטעון נתוני דוגמה ולראות איך זה עובד.</p>')
        if st.button('טעינת נתוני דוגמה', type='primary'):
            d_s, d_t = demo_rows()
            set_roster(d_s, d_t, 'demo')
            st.rerun()
    else:
        html(f'<p class="lead">{len(s_rows)} חיילים ו-{len(t_rows)} משימות מוכנים לשיבוץ.</p>')
        with st.expander('הגדרות'):
            g1, g2, g3 = st.columns(3)
            with g1:
                day_start_h = st.selectbox(
                    'היממה מתחילה בשעה', list(range(24)), index=6, format_func=lambda h: f'{h:02d}:00',
                    help='השבצ"ק מכסה 24 שעות מהשעה הזו. 06:00 או 08:00 שומרות על הלילה כרצף אחד.')
            with g2:
                sleep_h = st.selectbox('מנוחה רצופה לחייל', [5, 6, 7, 8], index=2,
                                       format_func=lambda h: f'{h} שעות')
            with g3:
                time_limit = st.selectbox('משך החישוב', [10, 20, 60], index=1,
                                          format_func=lambda s: {10: 'מהיר, עד 10 שניות', 20: 'רגיל, עד 20 שניות',
                                                                 60: 'יסודי, עד דקה'}[s])
        with st.expander('הנחיות מפקד'):
            html('<p class="hint">לא חובה. קבעו מראש מי חייב לקבל משמרת במשימה מסוימת ומי לא ישובץ אליה. '
                 'השעות מצמצמות את ההנחיה לחלון זמן, למשל 22-2.</p>')
            soldier_opts = [f"{r['full_name']} ({r['personal_id']})" for r in s_rows]
            task_opts = [f"{r['name']} ({r['task_no']})" for r in t_rows]
            d_df = st.data_editor(
                pd.DataFrame(columns=D_COLS).astype(str), num_rows='dynamic', hide_index=True,
                key=f'directives_{ss.ed_ver}', **FULL,
                column_config={
                    D_COLS[0]: st.column_config.SelectboxColumn(options=soldier_opts, required=True),
                    D_COLS[1]: st.column_config.SelectboxColumn(options=task_opts, required=True),
                    D_COLS[2]: st.column_config.SelectboxColumn(options=['חובה', 'אסור'], required=True),
                    D_COLS[3]: st.column_config.TextColumn(),
                })

        if st.button('יצירת שבצ"ק', type='primary'):
            try:
                d_rows = []
                for _, r in d_df.iterrows():
                    if not all(_filled(r[c]) for c in D_COLS[:3]):
                        continue
                    d_rows.append({'personal_id': str(r[D_COLS[0]]).rsplit('(', 1)[-1].rstrip(')'),
                                   'task_no': str(r[D_COLS[1]]).rsplit('(', 1)[-1].rstrip(')'),
                                   'kind': r[D_COLS[2]],
                                   'hours': r[D_COLS[3]] if _filled(r[D_COLS[3]]) else ''})
                cfg = sc.Config(day_start=day_start_h * 60, sleep_target=sleep_h * 60,
                                sleep_min=min(4, sleep_h - 1) * 60, time_limit=float(time_limit))
                with st.spinner('משבץ: קודם מאייש עמדות, אחר כך מסדר מנוחה, ובסוף מאזן עומס'):
                    ss.view = run_schedule(s_rows, t_rows, d_rows, cfg)
                st.rerun()
            except Exception:
                ss.view = None
                note('משהו השתבש בזמן החישוב. הפרטים למטה יעזרו לאתר את התקלה.', 'bad')
                st.code(traceback.format_exc())

        if ss.view:
            render_result(ss.view)

# ── נתונים ──────────────────────────────────────────────────────
with tab_data:
    s_df, t_df = rows_to_frames(s_rows, t_rows)
    d_soldiers, d_tasks, d_files = st.tabs([f'חיילים ({len(s_rows)})', f'משימות ({len(t_rows)})', 'קבצים'])
    with d_soldiers:
        html('<p class="hint">לחיצה על תא עורכת אותו, והשורה הריקה בתחתית מוסיפה חייל. '
             'פורמט השעות מוסבר בלשונית "מדריך".</p>')
        s_edit = st.data_editor(
            s_df, num_rows='dynamic', hide_index=True, key=f'ed_s_{ss.ed_ver}', **FULL,
            column_config={
                S_COLS['personal_id']: _col(st.column_config.TextColumn, width='small', pinned=True),
                S_COLS['full_name']: _col(st.column_config.TextColumn, pinned=True),
                S_COLS['exemptions']: _col(st.column_config.TextColumn, help='מס"ד של משימות שהחייל פטור מהן, למשל 104, 105'),
                S_COLS['roles']: _col(st.column_config.TextColumn, help='למשל: נהג, חובש'),
                S_COLS['blocked_hours']: _col(st.column_config.TextColumn, help='מתי החייל לא זמין, למשל 08:30-12:00'),
            })
    with d_tasks:
        html('<p class="hint">סימון בעמודה "כוננות" מתיר לשבץ את המשימה במקביל לכל משימה אחרת. '
             'העמידו את העכבר על כותרת של עמודה כדי לראות מה כותבים בה.</p>')
        num = st.column_config.NumberColumn
        t_edit = st.data_editor(
            t_df, num_rows='dynamic', hide_index=True, key=f'ed_t_{ss.ed_ver}', **FULL,
            column_config={
                T_COLS['task_no']: _col(num, 'מס"ד', width='small', step=1, format='%d', pinned=True),
                T_COLS['name']: _col(st.column_config.TextColumn, 'משימה', pinned=True),
                T_COLS['required']: _col(num, 'סד"כ', width='small', min_value=1, step=1, format='%d',
                                         help='כמה חיילים בכל משמרת'),
                T_COLS['shift_hours']: _col(num, 'משך', width='small', min_value=0.25, step=0.5, help='משך משמרת, בשעות'),
                T_COLS['rest_hours']: _col(num, 'מנוחה', width='small', min_value=0.0, step=0.5,
                                           help='כמה שעות מנוחה מגיעות אחרי משמרת במשימה הזו'),
                T_COLS['shared']: _col(st.column_config.CheckboxColumn, 'כוננות', width='small', default=False,
                                       help='משימה שמותרת במקביל לכל משימה אחרת'),
                T_COLS['active_hours']: _col(st.column_config.TextColumn, 'שעות פעילות',
                                             help='all לכל היממה, או טווח כמו 22-6'),
                T_COLS['required_roles']: _col(st.column_config.TextColumn, 'הסמכה נדרשת',
                                               help='הסמכה לכל תקן, למשל: מפקד, נהג. שאר התקנים פתוחים לכולם'),
                T_COLS['intensity']: _col(num, 'עצימות', width='small', min_value=1, max_value=10, step=1, format='%d',
                                          help='מ-1 עד 10. משפיע על חישוב ההוגנות'),
                T_COLS['blocked_roles']: _col(st.column_config.TextColumn, 'תפקידים חסומים',
                                              help='מי לא משובץ למשימה הזו'),
                T_COLS['priority']: _col(num, 'קריטיות', width='small', min_value=1, max_value=10, step=1, format='%d',
                                         help='מ-1 עד 10. כשחסרים חיילים, משימה קריטית מאוישת קודם'),
            })
    with d_files:
        html('<p class="hint">העלאה מחליפה את כל הנתונים השמורים. הקבצים צריכים להיות במבנה של התבניות.</p>')
        f1, f2 = st.columns(2)
        with f1:
            sf = st.file_uploader('קובץ חיילים', type='xlsx', key=f'sf_{ss.ed_ver}')
        with f2:
            tf = st.file_uploader('קובץ משימות', type='xlsx', key=f'tf_{ss.ed_ver}')
        if sf is not None and tf is not None:
            if st.button('החלפת הנתונים בקבצים שהועלו', type='primary'):
                try:
                    up_s, up_t = sc.rows_from_excel(pd.read_excel(sf), pd.read_excel(tf))
                    set_roster(up_s, up_t, 'upload', f'{len(up_s)} חיילים, {len(up_t)} משימות (Excel)')
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
        st.markdown('### הורדות')
        b1, b2, b3 = st.columns(3)
        ex_s, ex_t = rows_to_frames(*demo_rows())
        with b1:
            st.download_button('הנתונים שלי', data=to_excel({'חיילים': s_df, 'משימות': t_df}),
                               file_name='shavtzak_data.xlsx', disabled=not (s_rows or t_rows), key='dl_mine', **FULL)
        with b2:
            st.download_button('תבנית חיילים', data=to_excel({'Soldiers': ex_s.head(6)}),
                               file_name='Soldiers_template.xlsx', key='tpl_s', **FULL)
        with b3:
            st.download_button('תבנית משימות', data=to_excel({'Tasks': ex_t}),
                               file_name='Tasks_template.xlsx', key='tpl_t', **FULL)
        st.markdown('### נתוני דוגמה')
        html('<p class="hint">24 חיילים ו-6 משימות, כדי לנסות את המערכת. הטעינה מחליפה את הנתונים השמורים.</p>')
        if st.button('טעינת נתוני דוגמה', key='demo_data'):
            set_roster(*demo_rows(), 'demo')
            st.rerun()

    changed = not (s_edit.astype(str).fillna('').equals(s_df.astype(str).fillna(''))
                   and t_edit.astype(str).equals(t_df.astype(str)))
    if changed:
        note('יש שינויים שעוד לא נשמרו.', 'warn')
        if st.button('שמירת השינויים', type='primary'):
            try:
                new_s, new_t = sc.rows_from_excel(_clean(s_edit), _clean(t_edit))
                _, _, _, warns = sc.build_objects(new_s, new_t)
                set_roster(new_s, new_t, 'upload', f'{len(new_s)} חיילים, {len(new_t)} משימות')
                ss.save_warnings = warns
                st.rerun()
            except ValueError as e:
                st.error(str(e))
    for w in ss.pop('save_warnings', None) or []:
        note(esc(w), 'warn')

# ── היסטוריה ────────────────────────────────────────────────────
with tab_history:
    runs = get_runs()
    if runs.empty:
        html('<p class="lead">עוד אין שיבוצים שמורים.</p>'
             '<p class="hint">כל שבצ"ק שתיצרו נשמר כאן לבד, ואפשר לחזור אליו ולהוריד אותו שוב.</p>')
    else:
        options = {int(r['id']): f"{r['label']}, {r['num_soldiers']} חיילים, כיסוי {r['coverage_pct']:.0f}%"
                   for _, r in runs.iterrows()}
        chosen = st.selectbox('שיבוץ', list(options), format_func=lambda i: options[i])
        run, rows = get_run(chosen)
        if run:
            h_board, h_task = st.tabs(['לוח', 'לפי משימה'])
            with h_board:
                html(board_html(rows, run['roster'], run['day_start']))
            with h_task:
                html(by_task_html(rows, run['day_start']))
            if st.download_button('הורדה ל-Excel',
                                  data=schedule_excel(rows, run['roster'], run['day_start'],
                                                      {'השוואה': pd.DataFrame(run['compare'])}),
                                  file_name=f'shavtzak_{run["id"]}.xlsx', key=f'dl_hist{run["id"]}'):
                db.log_event(USER['id'], 'download', f'run {run["id"]}')
        with st.expander('מחיקת ההיסטוריה'):
            html('<p class="hint">המחיקה מסירה את כל השיבוצים השמורים שלכם, ואי אפשר לשחזר אותם.</p>')
            if st.button('מחיקת כל השיבוצים שלי'):
                db.delete_runs(USER['id'])
                ss.runs, ss.run_cache, ss.view = None, None, None
                st.rerun()

# ── מדריך ───────────────────────────────────────────────────────
with tab_guide:
    st.markdown("""
### איך המערכת מחליטה

יש סדר עדיפות קבוע, ושלב מאוחר לא בא על חשבון שלב שקדם לו.

1. **כל עמדה מאוישת.** תקן נשאר ריק רק כשאין שום דרך לאייש אותו בלי לשבור אילוץ קשיח. כשאין ברירה, מוותרים קודם על המשימות שהקריטיות שלהן נמוכה.
2. **מנוחה.** מנוחה אחרי כל משמרת לפי מה שהוגדר למשימה, ובלוק מנוחה רצוף לכל חייל, עדיף בלילה.
3. **הוגנות.** העומס מתחלק בין החיילים. משימה קשה שוקלת יותר, ושעת לילה נספרת כשעה וחצי.

**מה לא נשבר אף פעם:** הסמכה נדרשת, פטורים וחסימות אישיות, חייל בשתי משימות בלעדיות באותו זמן, חייל בשני תקנים באותה משמרת, והנחיות המפקד.

המערכת בונה את עיקר השבצ"ק. כל פשרה שנדרשה מופיעה בלשונית "לאישור המפקד", ואת התיקונים הקטנים עושים על הקובץ שמורידים.

### איך כותבים שעות
""")
    html(html_table(pd.DataFrame([
        ['8-20', 'מ-08:00 עד 20:00. שעת הסיום אינה כלולה'],
        ['08:30-16:00', 'בדיוק מ-08:30 עד 16:00'],
        ['22-6', 'מ-22:00 עד 06:00 למחרת'],
        ['8-12, 16-20', 'שני חלונות נפרדים'],
        ['all', 'כל היממה. במשימה גם תא ריק פירושו כל היממה; בחסימה של חייל, תא ריק פירושו שאין חסימה'],
    ], columns=['כותבים', 'המשמעות']), narrow=True, code_key=True))
    st.markdown('### העמודות של החיילים')
    html(html_table(pd.DataFrame([
        ['מספר אישי', 'מזהה ייחודי לכל חייל'],
        ['שם מלא', 'השם שיופיע בשבצ"ק'],
        ['פטורים', 'מס"ד של משימות שהחייל פטור מהן, למשל 104, 105'],
        ['הסמכות', 'תפקידים והסמכות, למשל נהג, חובש'],
        ['שעות חסימה', 'מתי החייל לא זמין, למשל 08:30-12:00'],
    ], columns=['עמודה', 'מה כותבים']), narrow=True))
    st.markdown('### העמודות של המשימות')
    html(html_table(pd.DataFrame([
        ['מס"ד משימה', 'מספר מזהה, למשל 101'],
        ['סד"כ נדרש למשימה', 'כמה חיילים בכל משמרת'],
        ['משך משמרת', 'בשעות. אפשר גם 1.5'],
        ['שעות מנוחה בין משימות', 'כמה מנוחה מגיעה אחרי משמרת במשימה הזו'],
        ['אישור חפיפה בין משימות', 'מסומן למשימה כמו כוננות, שמותרת במקביל לכל משימה אחרת'],
        ['שעות פעילות', 'מתי המשימה פעילה'],
        ['הסמכה נדרשת', 'הסמכה לכל תקן. "מפקד, נהג" במשימה של 3 פירושו מפקד, נהג ועוד חייל כלשהו'],
        ['דירוג עצימות משימה', 'מ-1 עד 10. משפיע על חישוב ההוגנות'],
        ['תפקידים חסומים', 'מי לא משובץ למשימה הזו'],
        ['קריטיות', 'מ-1 עד 10. כשחסרים חיילים, משימה קריטית מאוישת קודם'],
    ], columns=['עמודה', 'מה כותבים']), narrow=True))

# ── ניהול (למנהל המערכת בלבד) ───────────────────────────────────
if USER['is_admin']:
    with tabs[4]:
        import plotly.express as px

        info = db.status()
        if info['backend'] == 'postgres':
            note('הנתונים נשמרים בבסיס נתונים חיצוני, ונשארים גם כשהאתר מאותחל.')
        else:
            note('הנתונים נשמרים כרגע בקובץ מקומי על השרת. באתר החי הקובץ נמחק בכל אתחול ובכל עדכון קוד. '
                 'כדי שהנתונים יישמרו, הוסיפו ב-Secrets את השורה DATABASE_URL.', 'warn')

        if ss.get('admin_ov') is None or st.button('רענון הנתונים'):
            ss.admin_ov = db.admin_overview(days=30)
        ov = ss.admin_ov
        t = ov['totals']
        html('<div class="facts">'
             f'<div class="fact"><b>{t["users"]}</b><span>משתמשים רשומים</span></div>'
             f'<div class="fact"><b>{t["active_week"]}</b><span>משתמשים פעילים בשבוע האחרון</span></div>'
             f'<div class="fact"><b>{t["logins"]}</b><span>כניסות למערכת</span></div>'
             f'<div class="fact"><b>{t["runs"]}</b><span>שיבוצים שנוצרו, בכיסוי ממוצע של {t["avg_coverage"]:.0f}%</span></div>'
             '</div>')

        st.markdown('### פעילות יומית, 30 הימים האחרונים')
        daily = ov['daily'][ov['daily']['event'].isin(EVENT_COLORS)]
        if daily.empty:
            html('<p class="hint">עוד אין פעילות של משתמשים.</p>')
        else:
            fig = px.line(daily, x='day', y='n', color='event', markers=True,
                          color_discrete_map=EVENT_COLORS, category_orders={'event': list(EVENT_COLORS)})
            fig.update_traces(line_width=2, marker_size=8,
                              hovertemplate='%{x}<br>%{fullData.name}: %{y}<extra></extra>')
            fig.update_layout(height=320, plot_bgcolor='white', paper_bgcolor='rgba(0,0,0,0)',
                              font=dict(family='IBM Plex Sans Hebrew, Arial', size=13, color=INK),
                              margin=dict(t=40, b=30, l=20, r=20),
                              legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0, title_text=''))
            fig.update_xaxes(title_text='', type='category', gridcolor='#ECECE8', linecolor=RULE)
            fig.update_yaxes(title_text='פעולות ביום', rangemode='tozero', gridcolor='#ECECE8', linecolor=RULE,
                             tickformat=',d', dtick=1 if daily['n'].max() <= 10 else None)
            st.plotly_chart(fig, key='usage', **FULL)

        st.markdown('### משתמשים')
        if ov['users'].empty:
            html('<p class="hint">עוד אין משתמשים רשומים.</p>')
        else:
            html(html_table(ov['users'].fillna('')))
        with st.expander('100 הפעולות האחרונות'):
            html(html_table(ov['recent']))
        st.download_button('הורדת נתוני השימוש ל-Excel',
                           data=to_excel({'משתמשים': ov['users'],
                                          'פעילות יומית': ov['daily'].rename(columns={
                                              'day': 'תאריך', 'event': 'פעולה', 'n': 'כמות', 'users': 'משתמשים שונים'})
                                          [['תאריך', 'פעולה', 'כמות', 'משתמשים שונים']],
                                          'פעולות אחרונות': ov['recent']}),
                           file_name='usage_report.xlsx', key='usage_dl')

# עוגיית ההתחברות נכתבת בסוף הדף, כדי שלא תזיז רכיבים אחרים (הזזה מאפסת את הלשונית הפתוחה)
flush_cookie_job()
