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
import io
import os
import traceback

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit import config as _st_config

import database as db
import scheduler as sc

# המסך נשאר בהיר גם כשהדפדפן של המשתמש במצב כהה (אחרת הצבעים נשברים)
try:
    _changed = False
    for _k, _v in {'theme.base': 'light', 'theme.primaryColor': '#2d5a27'}.items():
        if _st_config.get_option(_k) != _v:
            _st_config.set_option(_k, _v)
            _changed = True
    if _changed:
        st.rerun()
except Exception:
    pass

st.set_page_config(page_title='שבצ"ק חכם', page_icon="🪖", layout="wide")

# רוחב מלא — תואם גם לגרסאות ישנות וגם לחדשות של Streamlit
_VER = tuple(int(x) for x in st.__version__.split('.')[:2])
FULL = {'width': 'stretch'} if _VER >= (1, 52) else {'use_container_width': True}

# צבעי המשימות בתרשים (סדר קבוע; משימה תשיעית ומעלה מקבלת אפור)
TASK_COLORS = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948']
OTHER_COLOR, SHORT_COLOR, BRAND = '#8a8a86', '#d03b3b', '#2d5a27'
EVENT_COLORS = {'כניסה': '#2a78d6', 'הרצת שיבוץ': '#eb6834', 'הרשמה': '#1baf7a'}
COOKIE = 'shavtzak_session'


# ══════════════════════════════════════════════════════════════════
# עיצוב
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');
html,body{direction:rtl}
.stApp,[data-testid="stAppViewContainer"],.main,.block-container,
.stMarkdown,p,span,li,label,div,
[data-testid="stText"],[data-testid="stMarkdownContainer"],
[data-testid="stAlert"],[data-testid="stExpander"] summary,
[data-testid="stFileUploader"] label,[data-testid="stFileUploader"] div,
[data-testid="stSlider"] label,[data-testid="stSelectbox"] label{
  font-family:'Heebo',sans-serif!important;direction:rtl!important;text-align:right!important}
/* האייקונים של Streamlit הם גופן; בלי השורה הזו הם מוצגים כטקסט ("upload") */
[data-testid="stIconMaterial"],.material-symbols-rounded{font-family:'Material Symbols Rounded'!important;direction:ltr!important}
.stApp{background:#f2f5f2}.block-container{padding:3.6rem 2.5rem 3rem;max-width:1400px}
button[kind="primary"],[data-testid="stBaseButton-primary"]{background:#2d5a27!important;
  border-color:#2d5a27!important;color:white!important;font-weight:700!important}
button[kind="primary"] p,[data-testid="stBaseButton-primary"] p{color:white!important;font-size:17px}
button[kind="primary"]:hover,[data-testid="stBaseButton-primary"]:hover{background:#1a3d17!important}
.app-header{background:linear-gradient(135deg,#1a3d17 0%,#2d5a27 60%,#3d7a35 100%);
  border-radius:16px;padding:26px 32px;margin-bottom:24px;
  box-shadow:0 4px 20px rgba(45,90,39,.25);text-align:right}
.app-header h1{font-size:clamp(22px,4vw,34px);font-weight:900;color:white;margin:0 0 6px;letter-spacing:-.5px}
.app-header p{font-size:15px;color:rgba(255,255,255,.88);margin:0}
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:10px}
[data-testid="stTabs"] [data-baseweb="tab"]{font-weight:600;padding:8px 14px;height:auto;white-space:nowrap}
[data-testid="stTabs"] [data-baseweb="tab"] p{font-size:15px;color:#5a7a57}
[data-testid="stTabs"] [aria-selected="true"] p{color:#1a3d17!important;font-weight:800}
/* תרשימים וטבלאות-נתונים מצוירים על canvas/SVG ואינם תומכים בכיוון RTL: בלי זה מספרים נחתכים */
[data-testid="stDataFrame"],[data-testid="stDataFrame"] *,
[data-testid="stPlotlyChart"],[data-testid="stPlotlyChart"] *,.js-plotly-plot,.js-plotly-plot *{
  direction:ltr!important;text-align:left!important}
[data-testid="stTable"] table{width:100%;border-collapse:collapse;font-size:14px;background:white;direction:rtl}
[data-testid="stTable"] th{background:#2d5a27!important;color:white!important;padding:9px 12px!important;
  font-weight:600!important;text-align:right!important}
[data-testid="stTable"] tbody th{background:#f4f8f3!important;color:#1a3d17!important;
  border-bottom:1px solid #eef1ee!important}
[data-testid="stTable"] td{padding:8px 12px!important;border-bottom:1px solid #eef1ee!important;text-align:right!important}
.metric-row{display:flex;flex-direction:row;gap:16px;margin:18px 0;flex-wrap:wrap}
.metric-card{flex:1;min-width:150px;background:white;border-radius:14px;padding:20px;
  border:1px solid #dde8dc;box-shadow:0 2px 8px rgba(45,90,39,.07);text-align:right;direction:rtl}
.mc-label{font-size:12px;color:#5f7d5c;font-weight:700;margin-bottom:6px}
.mc-value{font-size:32px;font-weight:900;color:#1a3d17;line-height:1}
.mc-sub{font-size:12px;color:#7d957a;margin-top:6px}
[data-testid="stFileUploader"]{background:white;border-radius:12px;padding:14px 16px;
  border:2px dashed #c0d8bc;direction:rtl;text-align:right}
.info-box{background:#edf5ec;border-right:5px solid #2d5a27;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#1a3d17;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.warn-box{background:#fff8e6;border-right:5px solid #e67e22;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#7a4500;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.error-box{background:#fdecea;border-right:5px solid #c0392b;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#7a0010;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
</style>
""", unsafe_allow_html=True)


def small_table(df: pd.DataFrame, limit: int = 40):
    """טבלה קצרה מוצגת כטבלת HTML (מימין לשמאל); ארוכה — כטבלת נתונים נגללת"""
    if len(df) <= limit:
        st.table(df.reset_index(drop=True).set_index(df.columns[0]))
    else:
        st.dataframe(df, hide_index=True, **FULL)


def box(kind: str, html: str):
    st.markdown(f'<div class="{kind}-box">{html}</div>', unsafe_allow_html=True)


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
def _boot(database_url, admin_pw, schema_version):
    # הפרמטרים הם חלק מהמפתח: שינוי ב-Secrets מפעיל את האתחול מחדש
    db.configure(database_url)
    db.init_db()
    db.ensure_admin(admin_pw)
    return True


try:
    _boot(_secret('DATABASE_URL', 'DATABASE_URL'), _secret('ADMIN_PASSWORD', 'SHAVTZAK_ADMIN_PASSWORD'),
          db.SCHEMA_VERSION)
except Exception as e:
    st.error('לא הצלחתי להתחבר לבסיס הנתונים. בדקו ששורת החיבור DATABASE_URL ב-Secrets הועתקה במלואה, '
             'בתוך מירכאות.')
    st.code(str(e))
    st.stop()

ss = st.session_state
ss.setdefault('user', None)
ss.setdefault('view', None)


# ══════════════════════════════════════════════════════════════════
# תבניות ונתוני דוגמה
# ══════════════════════════════════════════════════════════════════
S_COLS = {'personal_id': 'מספר אישי', 'full_name': 'שם מלא', 'exemptions': 'פטורים',
          'roles': 'הסמכות', 'blocked_hours': 'שעות חסימה', 'absences': 'ימי היעדרות'}
T_COLS = {'task_no': 'מס"ד משימה', 'name': 'שם המשימה', 'required': 'סד"כ נדרש למשימה',
          'shift_hours': 'משך משמרת', 'rest_hours': 'שעות מנוחה בין משימות',
          'shared': 'אישור חפיפה בין משימות', 'active_hours': 'שעות פעילות',
          'required_roles': 'הסמכה נדרשת', 'intensity': 'דירוג עצימות משימה (1-10)',
          'blocked_roles': 'תפקידים חסומים', 'priority': 'קריטיות (1-10)'}


def demo_rows():
    names = ['נועם לוי', 'איתי כהן', 'יובל מזרחי', 'עומר פרץ', 'דניאל ביטון', 'רועי אברהם',
             'אורי פרידמן', 'עידו דהן', 'תומר אזולאי', 'גיא שפירא', 'אלון חדד', 'יונתן גבאי',
             'ניר סויסה', 'שחר אוחיון', 'עמית קליין', 'מתן רוזן', 'ליאור נחום', 'אביב שלום',
             'בן וקנין', 'אסף מלכה', 'טל גולן', 'דור ברק', 'יאיר סגל', 'רן אשכנזי']
    roles = {0: 'מפקד', 6: 'מפקד', 12: 'מפקד', 18: 'מפקד',
             1: 'נהג', 7: 'נהג', 13: 'נהג', 19: 'נהג', 4: 'נהג,חובש', 10: 'חובש'}
    blocked = {3: '08:30-12:00', 9: '14:00-18:00', 15: '06:00-10:30'}
    exempt = {5: '104', 20: '105'}
    # ימי היעדרות משפיעים רק בשבצ"ק של כמה ימים: יום שלם, טווח ימים, או שעות ביום מסוים
    absences = {2: '3-4', 5: '2(8-12)', 14: '6-7', 20: '5', 22: '1(14-20), 7'}
    s_rows = [{'personal_id': str(7001 + i), 'full_name': n, 'exemptions': exempt.get(i, ''),
               'roles': roles.get(i, ''), 'blocked_hours': blocked.get(i, ''),
               'absences': absences.get(i, '')}
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


def rows_to_frames(s_rows, t_rows):
    s_df = pd.DataFrame([{S_COLS[k]: r.get(k, '') for k in S_COLS} for r in s_rows])
    t_df = pd.DataFrame([{T_COLS[k]: r.get(k, '') for k in T_COLS} for r in t_rows])
    if not t_df.empty:
        t_df[T_COLS['shared']] = t_df[T_COLS['shared']].map(lambda v: bool(v) and str(v) not in ('0', 'False'))
    return s_df, t_df


def to_excel(sheets: dict) -> bytes:
    """כמה טבלאות לקובץ Excel אחד, מימין לשמאל, עם כותרות מעוצבות"""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        for name, df in sheets.items():
            df.to_excel(w, index=False, sheet_name=name)
            wb, ws = w.book, w.sheets[name]
            ws.right_to_left()
            ws.freeze_panes(1, 1)
            head = wb.add_format({'bold': True, 'fg_color': BRAND, 'font_color': 'white',
                                  'border': 1, 'align': 'right', 'valign': 'vcenter'})
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, head)
                longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)])
                ws.set_column(ci, ci, min(longest + 3, 45))
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════
# תרשימים
# ══════════════════════════════════════════════════════════════════
def _plot_layout(fig, height):
    fig.update_layout(
        height=height, plot_bgcolor='white', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Heebo', size=13, color='#2b2b28'),
        margin=dict(t=40, b=40, l=20, r=20),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1, title_text=''),
        hoverlabel=dict(font_family='Heebo', align='right'))
    fig.update_xaxes(gridcolor='#ecece8', linecolor='#d9d9d4', zeroline=False)
    fig.update_yaxes(gridcolor='#ecece8', linecolor='#d9d9d4', zeroline=False)
    return fig


def timeline_chart(rows, roster, day_start):
    """ציר זמן: שורה לכל חייל, מקטע לכל משמרת. כוננות מוצגת כפס שקוף ברקע."""
    base = datetime.datetime(2000, 1, 1) + datetime.timedelta(minutes=day_start)
    seen, label = {}, {}
    for sid, name in roster:
        seen[name] = seen.get(name, 0) + 1
    for sid, name in roster:
        label[str(sid)] = f'{name} ({sid})' if seen[name] > 1 else name
    short_label = '⚠️ חוסר כוח אדם'

    recs = []
    for r in rows:
        who = label.get(str(r['personal_id']), r['soldier_name']) if r['personal_id'] else short_label
        recs.append({
            'חייל': who,
            'משימה': r['task_name'] if r['personal_id'] else 'חוסר',
            'התחלה': base + datetime.timedelta(minutes=r['start_min']),
            'סיום': base + datetime.timedelta(minutes=r['end_min']),
            'שעות': f"{sc.fmt_clock(r['start_min'], day_start)}–{sc.fmt_clock(r['end_min'], day_start)}",
            'פירוט': r['task_name'] + (f" · {r['role']}" if r['role'] else ''),
        })
    df = pd.DataFrame(recs)
    if df.empty:
        return None

    tasks_order, shared_names = [], set()
    for r in sorted(rows, key=lambda x: (not x['shared'], x['task_no'])):
        if r['task_name'] not in tasks_order:
            tasks_order.append(r['task_name'])
        if r['shared']:
            shared_names.add(r['task_name'])
    by_number = sorted({(r['task_no'], r['task_name']) for r in rows})
    colors = {name: (TASK_COLORS[i] if i < len(TASK_COLORS) else OTHER_COLOR)
              for i, (_, name) in enumerate(by_number)}
    colors['חוסר'] = SHORT_COLOR

    y_order = [label[str(sid)] for sid, _ in roster]
    if (df['חייל'] == short_label).any():
        y_order.append(short_label)

    fig = px.timeline(
        df, x_start='התחלה', x_end='סיום', y='חייל', color='משימה',
        color_discrete_map=colors,
        category_orders={'חייל': y_order, 'משימה': tasks_order + ['חוסר']},
        custom_data=['פירוט', 'שעות'])
    for tr in fig.data:
        tr.hovertemplate = '<b>%{y}</b><br>%{customdata[0]}<br>%{customdata[1]}<extra></extra>'
        tr.marker.line = dict(color='white', width=2)
        if tr.name in shared_names:
            tr.opacity, tr.width = 0.28, 0.9
        else:
            tr.width = 0.62
    _plot_layout(fig, max(400, 26 * len(y_order) + 150))
    fig.update_layout(barmode='overlay', margin=dict(t=95, l=130),
                      legend=dict(orientation='h', yanchor='bottom', y=1.07, xanchor='left', x=0))
    fig.update_xaxes(tickformat='%H:%M', dtick=2 * 3600 * 1000, side='top', title_text='',
                     range=[base, base + datetime.timedelta(hours=24)])
    fig.update_yaxes(title_text='', categoryorder='array', categoryarray=y_order[::-1],
                     autorange=True, automargin=True, ticksuffix='  ')
    return fig


def load_chart(summary_df):
    df = summary_df.sort_values('עומס משוקלל', ascending=False)
    fig = px.bar(df, x='שם', y='עומס משוקלל', custom_data=['שעות במשימות', 'מתוכן בלילה', 'שעות כוננות'])
    fig.update_traces(
        marker_color=BRAND, width=0.55,
        hovertemplate='<b>%{x}</b><br>עומס משוקלל: %{y}<br>שעות במשימות: %{customdata[0]}'
                      '<br>מתוכן בלילה: %{customdata[1]}<br>כוננות: %{customdata[2]}<extra></extra>')
    avg = df['עומס משוקלל'].mean()
    fig.add_hline(y=avg, line_width=1, line_color='#6b6b66',
                  annotation_text=f'ממוצע {avg:.1f}', annotation_position='top right',
                  annotation_font_color='#4a4a45')
    _plot_layout(fig, 380)
    fig.update_xaxes(title_text='', tickangle=-40)
    fig.update_yaxes(title_text='עומס משוקלל (שעות × עצימות)')
    return fig


def usage_chart(daily):
    df = daily[daily['event'].isin(EVENT_COLORS)]
    if df.empty:
        return None
    fig = px.line(df, x='day', y='n', color='event', markers=True, color_discrete_map=EVENT_COLORS,
                  category_orders={'event': list(EVENT_COLORS)})
    fig.update_traces(line_width=2, marker_size=8,
                      hovertemplate='%{x}<br>%{fullData.name}: %{y}<extra></extra>')
    _plot_layout(fig, 340)
    fig.update_xaxes(title_text='', type='category')
    fig.update_yaxes(title_text='מספר פעולות ביום', rangemode='tozero', tickformat=',d',
                     dtick=1 if df['n'].max() <= 10 else None)
    return fig


# ══════════════════════════════════════════════════════════════════
# כניסה, הרשמה והתחברות שנשמרת אחרי רענון
# ══════════════════════════════════════════════════════════════════
def sign_in(user: dict):
    ss.user = user
    try:
        ss.cookie_job = ('set', db.create_session(user['id']))
    except Exception:
        ss.cookie_job = None
    st.rerun()


def _cookie_token():
    try:
        return st.context.cookies.get(COOKIE)
    except Exception:
        return None


def sign_out():
    try:
        db.delete_session(ss.get('token') or _cookie_token())
    except Exception:
        pass
    for k in ('user', 'view', 'roster', 'runs', 'run_cache', 'token', 'admin_ov'):
        ss[k] = None
    ss.cookie_job = ('clear', '')
    st.rerun()


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
    st.markdown("""
    <div class="app-header">
      <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
      <p>בונה שבצ"ק ליום אחד או לשבוע שלם: כל עמדה מאוישת, מנוחה רצופה לחיילים, וחלוקת עומס הוגנת</p>
    </div>""", unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1.4, 1])
    with mid:
        t_in, t_new = st.tabs(["כניסה", "הרשמה"])
        with t_in:
            with st.form('login'):
                u = st.text_input('שם משתמש')
                p = st.text_input('סיסמה', type='password')
                if st.form_submit_button('כניסה', **FULL):
                    user = db.authenticate(u, p)
                    if user:
                        sign_in(user)
                    st.error('שם משתמש או סיסמה שגויים.')
        with t_new:
            with st.form('register'):
                u = st.text_input('שם משתמש (3 תווים לפחות)')
                unit = st.text_input('שם היחידה / המסגרת')
                p = st.text_input('סיסמה (6 תווים לפחות)', type='password')
                p2 = st.text_input('אימות סיסמה', type='password')
                if st.form_submit_button('יצירת חשבון', **FULL):
                    if p != p2:
                        st.error('הסיסמאות אינן זהות.')
                    else:
                        user, err = db.create_user(u, p, unit)
                        if err:
                            st.error(err)
                        else:
                            sign_in(user)
        box('info', 'הנתונים שתזינו נשמרים בחשבון שלכם בלבד. '
                    'מומלץ לעבוד עם שמות ומספרים שאינם מזהים חיילים אמיתיים.')
    flush_cookie_job()


if ss.user is None:
    auth_screen()
    st.stop()

USER = ss.user


# ── נתונים שמורים: נשמרים גם בזיכרון הגלישה, כדי לא לפנות לבסיס הנתונים בכל לחיצה ──
def get_roster():
    if ss.get('roster') is None:
        ss.roster = db.load_roster(USER['id'])
    return ss.roster


def set_roster(s_rows, t_rows, event: str, details: str = ''):
    db.save_roster(USER['id'], s_rows, t_rows)
    db.log_event(USER['id'], event, details)
    ss.roster = None


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
# הרצת שיבוץ
# ══════════════════════════════════════════════════════════════════
def run_schedule(s_rows, t_rows, d_rows, cfg: sc.Config, days: int = 1, progress=None):
    soldiers, tasks, directives, warns = sc.build_objects(s_rows, t_rows, d_rows)
    res = sc.solve_multi(soldiers, tasks, cfg, directives, days=days, progress=progress)
    m = res['metrics']
    rows = res['rows']
    roster = [(s.sid, s.name) for s in soldiers]
    compare = sc.comparison_table(res['baseline_metrics'], m, len(soldiers))
    stamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
    label = f'שיבוץ {stamp}' if days == 1 else f'שיבוץ ל-{days} ימים {stamp}'
    run_id = db.save_run(USER['id'], label, cfg.day_start, res['engine'], len(tasks), m,
                         res['seconds'], roster, compare.to_dict('records'), rows,
                         extra={'num_days': days, 'absent': res['absent']})
    db.log_event(USER['id'], 'run',
                 f"{len(soldiers)} חיילים, {len(tasks)} משימות, {days} ימים, כיסוי {m['coverage_pct']:.0f}%")
    ss.runs = None
    return {
        'run_id': run_id, 'label': label, 'rows': rows, 'roster': roster,
        'num_days': days, 'absent': res['absent'],
        'day_start': cfg.day_start, 'sleep_target': cfg.sleep_target,
        'num_tasks': len(tasks), 'engine': res['engine'], 'seconds': res['seconds'],
        'warnings': warns + res['warnings'],
        'stages': [{k: v for k, v in s.items() if not k.startswith('_')} for s in res['stages']],
        'compare': compare,
        'summary': sc.summary_table(m),
        'violations': sc.violations_table_multi(res['day_metrics'], cfg),
        'shortages': pd.DataFrame(res['shortage_reasons']),
        'hard': m['hard_violations'],
        'm': {k: m[k] for k in ('short_minutes', 'coverage_pct', 'full_rest', 'full_night_rest',
                                'load_gap', 'hours_gap', 'slots_total', 'short_slots', 'soldier_days')},
        'n_rest_viol': len(m['rest_violations']),
    }


def render_schedule(rows, roster, day_start, key: str, summary=None, violations=None,
                    shortages=None, compare=None, absent=None):
    """הצגת שבצ"ק: ציר זמן, טבלאות, והורדה. משמש גם את מסך התוצאה וגם את ההיסטוריה.
    בשבצ"ק של כמה ימים בוחרים איזה יום להציג; קובץ ה-Excel כולל תמיד את כל הימים."""
    absent = {int(d): [str(x) for x in sids] for d, sids in (absent or {}).items()}
    n_days = sc.num_days_of(rows)
    day = 1
    if n_days > 1:
        day = st.radio('איזה יום להציג', list(range(1, n_days + 1)), horizontal=True,
                       format_func=lambda d: f'יום {d}', key=f'day_{key}')
    day_rows = sc.rows_of_day(rows, day)
    gone = set(absent.get(day, []))
    day_roster = [(sid, name) for sid, name in roster if str(sid) not in gone]
    task_df = sc.task_table(day_rows, day_start)
    grid_df = sc.soldier_grid(day_rows, roster, day_start, absent=gone)

    fig = timeline_chart(day_rows, day_roster, day_start)
    if fig is not None:
        st.plotly_chart(fig, key=f'tl_{key}', **FULL)
        st.caption('כל שורה היא חייל וכל מקטע הוא משמרת. פס שקוף ברקע הוא כוננות (משימה שיתופית). '
                   'העמידו את העכבר על מקטע כדי לראות פרטים.')
    if gone:
        st.caption('לא נמצאים ביום הזה: ' + ', '.join(name for sid, name in roster if str(sid) in gone))

    v_grid, v_task = st.tabs(['טבלה לפי חייל', 'טבלה לפי משימה'])
    with v_grid:
        st.dataframe(grid_df, hide_index=True, **FULL)
    with v_task:
        st.dataframe(task_df, hide_index=True, **FULL)

    if n_days == 1:
        sheets = {'לפי משימה': task_df, 'לפי חייל': grid_df}
    else:
        sheets = {'לפי משימה': sc.task_table(rows, day_start)}
        for d in range(1, n_days + 1):
            sheets[f'לפי חייל - יום {d}'] = sc.soldier_grid(
                sc.rows_of_day(rows, d), roster, day_start, absent=absent.get(d, []))
    if summary is not None and not summary.empty:
        sheets['סיכום לחייל'] = summary
    if violations is not None and not violations.empty:
        sheets['חריגות לאישור'] = violations
    if shortages is not None and not shortages.empty:
        sheets['חוסרים'] = shortages
    if compare is not None and not compare.empty:
        sheets['השוואה'] = compare
    if st.download_button('📥 הורדת השבצ"ק (Excel)' + (' — כל הימים' if n_days > 1 else ''),
                          data=to_excel(sheets),
                          file_name=f'shavtzak_{key}.xlsx', key=f'dl_{key}', **FULL):
        db.log_event(USER['id'], 'download', key)


def render_result(v):
    m = v['m']
    if v['hard']:
        box('error', '🚨 נמצאה הפרה של אילוץ קשיח — אל תשתמשו בשבצ"ק הזה:<br>' + '<br>'.join(v['hard']))
    for w in v['warnings']:
        box('warn', '⚠️ ' + w)

    if m['short_minutes'] == 0:
        box('info', f"✅ <b>כיסוי מלא</b> — כל {m['slots_total']} התקנים מאוישים.")
    else:
        box('warn', f"⚠️ <b>{m['short_slots']} תקנים לא אוישו ({m['short_minutes'] / 60:.1f} שעות-עמדה).</b> "
                    f"אין דרך לאייש אותם בלי להפר אילוץ קשיח (הסמכה, פטור, חסימה אישית או "
                    f"שתי משימות בלעדיות באותו זמן). הפירוט והסיבה לכל חוסר מופיעים למטה.")

    n = len(v['roster'])
    days = v.get('num_days', 1)
    multi = days > 1
    rest_of = m.get('soldier_days', n) if multi else n
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="mc-label">כיסוי עמדות</div>
        <div class="mc-value">{m['coverage_pct']:.0f}%</div>
        <div class="mc-sub">{n} חיילים · {v['num_tasks']} משימות{f' · {days} ימים' if multi else ''}</div></div>
      <div class="metric-card"><div class="mc-label">מנוחה רצופה מלאה</div>
        <div class="mc-value">{m['full_rest']}/{rest_of}</div>
        <div class="mc-sub">{'ימי-חייל · ' if multi else ''}{m['full_night_rest']} מהם בלילה · יעד {sc.fmt_dur(v['sleep_target'])} שעות</div></div>
      <div class="metric-card"><div class="mc-label">מנוחה קצרה מהנדרש</div>
        <div class="mc-value">{v['n_rest_viol']}</div>
        <div class="mc-sub">מקרים בין משמרות</div></div>
      <div class="metric-card"><div class="mc-label">פער עומס</div>
        <div class="mc-value">{m['load_gap']:.0f}</div>
        <div class="mc-sub">{'ממוצע ליום, ' if multi else ''}בין העמוס ביותר לפנוי ביותר · {m['hours_gap']:.0f} שעות</div></div>
      <div class="metric-card"><div class="mc-label">זמן חישוב</div>
        <div class="mc-value">{v['seconds']:.0f}״</div>
        <div class="mc-sub">מנוע: {v['engine']}</div></div>
    </div>""", unsafe_allow_html=True)

    if not v['shortages'].empty:
        st.subheader('❌ תקנים שלא אוישו, והסיבה')
        small_table(v['shortages'])

    st.subheader('📅 השבצ"ק')
    render_schedule(v['rows'], v['roster'], v['day_start'], key=f"run{v['run_id']}",
                    summary=v['summary'], violations=v['violations'],
                    shortages=v['shortages'], compare=v['compare'], absent=v.get('absent'))

    st.subheader('🔎 חריגות לאישור המפקד')
    if v['violations'].empty:
        box('info', 'אין חריגות: כל חייל מקבל את המנוחה הנדרשת אחרי כל משמרת ומנוחה רצופה מלאה.')
    else:
        st.caption('המקומות שבהם המערכת ויתרה על מנוחה כדי שכל עמדה תאויש. '
                   'כיסוי העמדות תמיד קודם; אלה הפשרות שנדרשו בשבילו.')
        small_table(v['violations'])

    st.subheader('⚖️ חלוקת העומס')
    if multi:
        st.caption('העומס מוצג כממוצע ליום נוכחות, כדי שחייל שהיה חלק מהימים בבית לא ייראה כמי שעבד פחות.')
    st.plotly_chart(load_chart(v['summary']), key=f"load{v['run_id']}", **FULL)
    with st.expander('טבלת סיכום לכל חייל'):
        st.dataframe(v['summary'], hide_index=True, **FULL)

    st.subheader('📊 לפני ואחרי: שיבוץ בסיסי מול שיבוץ ממוטב')
    st.caption('"שיבוץ בסיסי" מחקה את הדרך הידנית: עוברים על המשמרות לפי סדר הזמן ומשבצים בכל פעם '
               'את החייל הפנוי שעבד הכי פחות. "שיבוץ ממוטב" הוא תוצאת האלגוריתם.')
    small_table(v['compare'])

    with st.expander('איך האלגוריתם הגיע לתוצאה'):
        st.markdown('האופטימיזציה נעשית בשלושה שלבים, לפי סדר עדיפות קשיח. '
                    'כל שלב ננעל לפני שעוברים לשלב הבא, כך ששלב מאוחר לא יכול לפגוע בשלב שקדם לו.')
        if v['stages']:
            small_table(pd.DataFrame(v['stages']))


def collect_inputs(sf, tf, df_, saved_s, saved_t):
    """מחזיר (חיילים, משימות, הנחיות), או None אם חסר משהו (ההודעה כבר הוצגה)."""
    if sf is not None and tf is not None:
        up_s, up_t = sc.rows_from_excel(pd.read_excel(sf), pd.read_excel(tf))
        set_roster(up_s, up_t, 'upload', f'{len(up_s)} חיילים, {len(up_t)} משימות')
        s_rows, t_rows = get_roster()
    elif sf is not None or tf is not None:
        st.error('כדי להחליף את הנתונים השמורים צריך להעלות גם קובץ חיילים וגם קובץ משימות.')
        return None
    else:
        s_rows, t_rows = saved_s, saved_t
    if not s_rows or not t_rows:
        st.error('אין נתונים להרצה. העלו קבצים או טענו נתוני דוגמה.')
        return None
    d_rows = sc.directives_from_excel(pd.read_excel(df_)) if df_ is not None else []
    return s_rows, t_rows, d_rows


# ══════════════════════════════════════════════════════════════════
# מסך ראשי
# ══════════════════════════════════════════════════════════════════
head_l, head_r = st.columns([5, 1])
with head_l:
    st.markdown(f"""
    <div class="app-header">
      <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
      <p>שלום {USER['username']}{' · ' + USER['unit_name'] if USER['unit_name'] else ''}</p>
    </div>""", unsafe_allow_html=True)
with head_r:
    st.write('')
    if st.button('יציאה', **FULL):
        sign_out()

tab_names = ["🚀  שיבוץ", "📂  היסטוריה", "📖  מדריך", "📥  תבניות"]
if USER['is_admin']:
    tab_names.append("🔐  ניהול")
tabs = st.tabs(tab_names)
tab_run, tab_history, tab_guide, tab_templates = tabs[:4]

# ── שיבוץ ───────────────────────────────────────────────────────
with tab_run:
    saved_s, saved_t = get_roster()

    st.markdown('#### 1. נתונים')
    if saved_s and saved_t:
        box('info', f'בחשבון שלכם שמורים <b>{len(saved_s)} חיילים</b> ו-<b>{len(saved_t)} משימות</b>. '
                    f'אפשר להריץ שיבוץ מיד, או להעלות קבצים חדשים שיחליפו אותם.')
    else:
        box('info', 'כדי להתחיל, העלו קובץ חיילים וקובץ משימות (יש תבניות להורדה בלשונית "תבניות"), '
                    'או טענו נתוני דוגמה כדי לראות איך המערכת עובדת.')

    c1, c2, c3 = st.columns(3)
    with c1:
        sf = st.file_uploader('קובץ חיילים (xlsx)', type='xlsx', key='sf')
    with c2:
        tf = st.file_uploader('קובץ משימות (xlsx)', type='xlsx', key='tf')
    with c3:
        df_ = st.file_uploader('הנחיות מפקד (xlsx, לא חובה)', type='xlsx', key='df')

    if st.button('טעינת נתוני דוגמה (24 חיילים, 6 משימות)'):
        set_roster(*demo_rows(), 'demo')
        st.rerun()

    st.markdown('#### 2. הגדרות')
    g0, g1, g2, g3 = st.columns(4)
    with g0:
        num_days = st.selectbox(
            'משך השבצ"ק', list(range(1, sc.MAX_DAYS + 1)), index=0,
            format_func=lambda d: 'יום אחד' if d == 1 else ('יומיים' if d == 2 else ('שבוע (7 ימים)' if d == 7 else f'{d} ימים')),
            help='בשבצ"ק של כמה ימים המשימות חוזרות בכל יום. המערכת זוכרת מיום ליום כמה כל חייל כבר '
                 'עבד ומתי סיים את המשמרת האחרונה, ומתחשבת בעמודה "ימי היעדרות" שבקובץ החיילים.')
    with g1:
        day_start_h = st.selectbox(
            'שעת תחילת היממה', list(range(24)), index=6, format_func=lambda h: f'{h:02d}:00',
            help='השבצ"ק מכסה 24 שעות מהשעה הזו. מומלץ 06:00 או 08:00, כדי שהלילה יישאר רצף אחד '
                 'ומשימות לילה לא ייחתכו באמצע.')
    with g2:
        sleep_h = st.selectbox('יעד מנוחה רצופה לחייל (שעות)', [5, 6, 7, 8], index=2)
    with g3:
        time_limit = st.slider('זמן חישוב מרבי ליום (שניות)', 5, 120, 20, 5,
                               help='החישוב נעצר מוקדם יותר כשאין עוד שיפור. במחלקה רגילה 20 שניות מספיקות. '
                                    'בשבצ"ק של כמה ימים הזמן הזה הוא לכל יום בנפרד.')

    st.markdown('#### 3. הרצה')
    if st.button('⚙️ יצירת שבצ"ק', type='primary', **FULL):
        done = False
        try:
            inputs = collect_inputs(sf, tf, df_, saved_s, saved_t)
            if inputs is not None:
                cfg = sc.Config(day_start=day_start_h * 60, sleep_target=sleep_h * 60,
                                sleep_min=min(4, sleep_h - 1) * 60, time_limit=float(time_limit))
                bar = st.progress(0.0, text='מתחיל...') if num_days > 1 else None
                step = (lambda d, n: bar.progress((d - 1) / n, text=f'מחשב את יום {d} מתוך {n}...')) if bar else None
                with st.spinner('מחשב שיבוץ: כיסוי עמדות ← מנוחה ושינה ← הוגנות...'):
                    ss.view = run_schedule(*inputs, cfg, days=num_days, progress=step)
                done = True
        except ValueError as e:
            ss.view = None
            box('error', f'❌ {e}')
        except Exception:
            ss.view = None
            box('error', '🚨 שגיאה טכנית:')
            st.code(traceback.format_exc())
        if done:
            st.rerun()      # ציור נקי של התוצאה, כדי שהלשוניות לא יתאפסו בלחיצה הבאה

    if ss.view:
        st.markdown('---')
        render_result(ss.view)

# ── היסטוריה ────────────────────────────────────────────────────
with tab_history:
    st.markdown('### 📂 השיבוצים שלכם')
    hist = get_runs()
    if hist.empty:
        box('warn', 'עדיין אין שיבוצים שמורים. כל שיבוץ שתיצרו נשמר כאן אוטומטית.')
    else:
        st.dataframe(hist.rename(columns={
            'id': 'מספר', 'created_at': 'נוצר', 'label': 'שם', 'num_soldiers': 'חיילים',
            'num_tasks': 'משימות', 'short_hours': 'שעות לא מאוישות', 'coverage_pct': 'כיסוי %',
            'rest_violations': 'חריגות מנוחה', 'full_rest': 'מנוחה מלאה',
            'load_gap': 'פער עומס', 'solve_seconds': 'זמן חישוב (שנ׳)'}), hide_index=True, **FULL)
        options = {f"#{r['id']} · {r['label']}": int(r['id']) for _, r in hist.iterrows()}
        chosen = st.selectbox('הצגת שיבוץ', list(options))
        run, rows = get_run(options[chosen])
        if run:
            render_schedule(rows, run['roster'], run['day_start'], key=f"hist{run['id']}",
                            compare=pd.DataFrame(run['compare']),
                            absent=run.get('extra', {}).get('absent'))
        st.markdown('---')
        if st.button('🗑️ מחיקת כל ההיסטוריה שלי'):
            db.delete_runs(USER['id'])
            ss.runs, ss.run_cache, ss.view = None, None, None
            st.rerun()

# ── מדריך ───────────────────────────────────────────────────────
with tab_guide:
    st.markdown("""
### איך המערכת מחליטה

המערכת פועלת לפי סדר עדיפות קבוע. שלב מאוחר לעולם לא בא על חשבון שלב שקדם לו.

1. **כיסוי עמדות — מעל הכול.** קודם כול מאיישים כל תקן שאפשר. תקן נשאר ריק רק כשאין שום
   דרך לאייש אותו בלי להפר אילוץ קשיח. כשאין ברירה, מוותרים קודם על המשימות שהקריטיות שלהן נמוכה.
2. **מנוחה.** מנוחה אחרי כל משמרת לפי מה שהוגדר למשימה, ובלוק מנוחה רצוף לכל חייל, עדיף בלילה (22:00–06:00).
3. **הוגנות.** איזון העומס בין החיילים. העומס משוקלל: שעות × עצימות המשימה, ושעת לילה נחשבת 1.5 שעות.

**אילוצים קשיחים (לא מופרים אף פעם):** הסמכה נדרשת · פטורים, חסימות אישיות ותפקידים חסומים ·
חייל לא נמצא בשתי משימות בלעדיות באותו זמן · חייל לא תופס שני תקנים באותה משמרת · הנחיות מפקד.

**כוננות (משימה שיתופית):** מותרת במקביל לכל משימה אחרת, ואינה נחשבת הפרעה למנוחה.

### איך כותבים שעות

| כותבים | המשמעות |
|---|---|
| `8-20` | מ-08:00 עד 20:00 (12 שעות; שעת הסיום אינה כלולה) |
| `08:30-16:00` | בדיוק מ-08:30 עד 16:00 |
| `22-6` | מ-22:00 עד 06:00 למחרת (חוצה חצות) |
| `8-12, 16-20` | שני חלונות נפרדים |
| `all` או תא ריק | במשימה: כל היממה. בחסימת חייל: תא ריק = אין חסימה |

### קובץ חיילים

| עמודה | הסבר | דוגמה |
|---|---|---|
| מספר אישי | מזהה ייחודי | 7001 |
| שם מלא | שם להצגה | נועם לוי |
| פטורים | מס"ד של משימות שהחייל פטור מהן | 104, 105 |
| הסמכות | תפקידים והסמכות | נהג, חובש |
| שעות חסימה | מתי החייל לא זמין, בכל יום | 08:30-12:00 |
| ימי היעדרות | לא חובה. באילו ימים של השבצ"ק החייל לא נמצא | 3-4 |

### שבצ"ק של כמה ימים

בלשונית "שיבוץ" בוחרים את משך השבצ"ק: מיום אחד ועד שבוע. המשימות חוזרות על עצמן בכל יום.
היום הראשון של השבצ"ק הוא יום 1, זה שאחריו יום 2, וכן הלאה.

המערכת משבצת יום אחרי יום, וכל יום "זוכר" את הימים שלפניו:

- **מנוחה.** מי שסיים משמרת בסוף היום לא יקבל משמרת בתחילת היום שאחריו, אם יש מי שיחליף אותו.
- **הוגנות.** מי שעבד הרבה אתמול יעבוד פחות היום, כך שהעומס מתאזן לאורך כל התקופה.
- **היעדרות.** חייל שלא נמצא לא משובץ, וכשהוא חוזר הוא לא נדרש "להשלים" את מה שהחסיר.

בעמודה **ימי היעדרות** כותבים:

| כותבים | המשמעות |
|---|---|
| `3` | לא נמצא ביום 3 כולו |
| `3-5` | לא נמצא בימים 3 עד 5 |
| `2(8-12)` | ביום 2 בלבד לא זמין מ-08:00 עד 12:00 |
| `1, 4(14-18), 6-7` | אפשר לשלב, עם פסיק בין הפריטים |

ההבדל בין שתי העמודות: "שעות חסימה" חוזרת בכל יום, ו"ימי היעדרות" היא ליום מסוים.

### קובץ משימות

| עמודה | הסבר | דוגמה |
|---|---|---|
| מס"ד משימה | מזהה | 101 |
| שם המשימה | שם קצר | ש"ג |
| סד"כ נדרש למשימה | כמה חיילים בכל משמרת | 2 |
| משך משמרת | בשעות (אפשר גם 1.5) | 4 |
| שעות מנוחה בין משימות | מנוחה נדרשת אחרי משמרת | 8 |
| אישור חפיפה בין משימות | TRUE = כוננות / משימה שיתופית | FALSE |
| שעות פעילות | מתי המשימה פעילה | all |
| הסמכה נדרשת | הסמכה לכל תקן; תקנים בלי הסמכה נשארים פתוחים לכולם | מפקד, נהג |
| דירוג עצימות משימה (1-10) | כמה המשימה קשה; משפיע על חישוב ההוגנות | 7 |
| תפקידים חסומים | מי לא משובץ למשימה | מפקד |
| קריטיות (1-10) | לא חובה. כשיש חוסר, משימה קריטית מאוישת קודם | 9 |

### הנחיות מפקד (לא חובה)

קובץ שלישי עם העמודות: **מספר אישי · מס"ד משימה · סוג · שעות**.
בעמודה "סוג" כותבים `חובה` (החייל חייב לקבל משמרת במשימה הזו) או `אסור` (החייל לא ישובץ אליה).
העמודה "שעות" מצמצמת את ההנחיה לחלון זמן, למשל `22-2`.
בשבצ"ק של כמה ימים ההנחיות חלות על כל יום.
""")

# ── תבניות ──────────────────────────────────────────────────────
with tab_templates:
    box('info', 'הורידו את שלושת הקבצים, מלאו אותם בנתונים שלכם, והעלו אותם בלשונית "שיבוץ". '
                'כדאי לשמור על שמות העמודות כמו שהם.')
    d_s, d_t = demo_rows()
    s_ex, t_ex = rows_to_frames(d_s[:6], d_t)
    dir_ex = pd.DataFrame([{'מספר אישי': 7003, 'מס"ד משימה': 101, 'סוג': 'אסור', 'שעות': ''},
                           {'מספר אישי': 7002, 'מס"ד משימה': 105, 'סוג': 'חובה', 'שעות': '22-2'}])
    st.markdown('**👥 תבנית חיילים**')
    st.dataframe(s_ex, hide_index=True, **FULL)
    st.download_button('⬇️ הורדת תבנית חיילים', data=to_excel({'Soldiers': s_ex}),
                       file_name='Soldiers_template.xlsx', key='tpl_s')
    st.markdown('**📋 תבנית משימות**')
    st.dataframe(t_ex, hide_index=True, **FULL)
    st.download_button('⬇️ הורדת תבנית משימות', data=to_excel({'Tasks': t_ex}),
                       file_name='Tasks_template.xlsx', key='tpl_t')
    st.markdown('**📌 תבנית הנחיות מפקד (לא חובה)**')
    st.dataframe(dir_ex, hide_index=True, **FULL)
    st.download_button('⬇️ הורדת תבנית הנחיות', data=to_excel({'Directives': dir_ex}),
                       file_name='Directives_template.xlsx', key='tpl_d')
    st.markdown('---')
    st.markdown('**🗂️ קובצי דוגמה מלאים** — 24 חיילים ו-6 משימות, מוכנים להעלאה כמו שהם')
    full_s, full_t = rows_to_frames(d_s, d_t)
    e1, e2 = st.columns(2)
    with e1:
        st.download_button('⬇️ קובץ חיילים לדוגמה', data=to_excel({'Soldiers': full_s}),
                           file_name='Soldiers_example.xlsx', key='ex_s', **FULL)
    with e2:
        st.download_button('⬇️ קובץ משימות לדוגמה', data=to_excel({'Tasks': full_t}),
                           file_name='Tasks_example.xlsx', key='ex_t', **FULL)

# ── ניהול (למנהל המערכת בלבד) ───────────────────────────────────
if USER['is_admin']:
    with tabs[4]:
        st.markdown('### 🔐 נתוני שימוש במערכת')
        st.caption('המסך הזה מוצג רק למנהל המערכת. משתמשים רגילים רואים רק את הנתונים של עצמם.')
        if db.status()['backend'] == 'postgres':
            box('info', '✅ הנתונים נשמרים בבסיס נתונים חיצוני, ונשארים גם כשהאתר מאותחל.')
        else:
            box('warn', '⚠️ הנתונים נשמרים כרגע בקובץ מקומי על השרת. באתר החי הקובץ נמחק בכל אתחול ובכל '
                        'עדכון קוד. כדי שהנתונים יישמרו, הוסיפו ב-Secrets את השורה DATABASE_URL.')
        if ss.get('admin_ov') is None or st.button('🔄 רענון הנתונים'):
            ss.admin_ov = db.admin_overview(days=30)
        ov = ss.admin_ov
        t = ov['totals']
        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-card"><div class="mc-label">משתמשים רשומים</div>
            <div class="mc-value">{t['users']}</div><div class="mc-sub">מאז ההשקה</div></div>
          <div class="metric-card"><div class="mc-label">פעילים בשבוע האחרון</div>
            <div class="mc-value">{t['active_week']}</div><div class="mc-sub">משתמשים שונים</div></div>
          <div class="metric-card"><div class="mc-label">כניסות למערכת</div>
            <div class="mc-value">{t['logins']}</div><div class="mc-sub">סה"כ</div></div>
          <div class="metric-card"><div class="mc-label">שיבוצים שנוצרו</div>
            <div class="mc-value">{t['runs']}</div>
            <div class="mc-sub">כיסוי ממוצע {t['avg_coverage']:.0f}% · {t['avg_seconds']:.0f} שנ׳ לחישוב</div></div>
        </div>""", unsafe_allow_html=True)

        st.markdown('#### פעילות יומית ב-30 הימים האחרונים')
        fig = usage_chart(ov['daily'])
        if fig is None:
            box('warn', 'עדיין אין פעילות של משתמשים.')
        else:
            st.plotly_chart(fig, key='usage', **FULL)
            with st.expander('הנתונים כטבלה'):
                st.dataframe(ov['daily'].rename(columns={'day': 'תאריך', 'event': 'פעולה', 'n': 'כמות',
                                                         'users': 'משתמשים שונים'})
                             [['תאריך', 'פעולה', 'כמות', 'משתמשים שונים']], hide_index=True, **FULL)

        st.markdown('#### משתמשים')
        st.dataframe(ov['users'], hide_index=True, **FULL)
        st.markdown('#### 100 הפעולות האחרונות')
        st.dataframe(ov['recent'], hide_index=True, **FULL)
        st.download_button('📥 הורדת נתוני השימוש (Excel)',
                           data=to_excel({'משתמשים': ov['users'], 'פעילות יומית': ov['daily'],
                                          'פעולות אחרונות': ov['recent']}),
                           file_name='usage_report.xlsx', key='usage_dl')

# עוגיית ההתחברות נכתבת בסוף הדף, כדי שלא תזיז רכיבים אחרים (הזזה מאפסת את הלשונית הפתוחה)
flush_cookie_job()
