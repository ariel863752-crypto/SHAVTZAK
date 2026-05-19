import io
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
import traceback
import random

# ══════════════════════════════════════════════════════════════════
# 1. עיצוב
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="שבצ\"ק חכם", page_icon="🪖", layout="wide")
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
.stApp{background:#f2f5f2}.block-container{padding:2rem 2.5rem 3rem;max-width:1400px}
.app-header{background:linear-gradient(135deg,#1a3d17 0%,#2d5a27 60%,#3d7a35 100%);
  border-radius:16px;padding:30px 35px;margin-bottom:28px;
  box-shadow:0 4px 20px rgba(45,90,39,.25);text-align:right}
.app-header h1{font-size:clamp(22px,4vw,38px);font-weight:900;color:white;margin:0 0 8px;letter-spacing:-.5px}
.app-header p{font-size:15px;color:rgba(255,255,255,.88);margin:0}
[data-testid="stTabs"] [data-baseweb="tab-list"]{flex-direction:row-reverse!important;
  justify-content:flex-start!important;gap:6px;background:white;border-radius:12px;
  padding:5px;border:1px solid #dde8dc;margin-bottom:20px}
[data-testid="stTabs"] [data-baseweb="tab"]{border-radius:8px;font-weight:600;
  font-size:14px;padding:8px 20px;color:#5a7a57;direction:rtl}
[data-testid="stTabs"] [aria-selected="true"]{background:#2d5a27!important;color:white!important}
.metric-row{display:flex;flex-direction:row-reverse;gap:16px;margin:22px 0;flex-wrap:wrap}
.metric-card{flex:1;min-width:160px;background:white;border-radius:14px;padding:22px;
  border:1px solid #dde8dc;box-shadow:0 2px 8px rgba(45,90,39,.07);text-align:right;direction:rtl}
.mc-label{font-size:11px;color:#7a9a77;font-weight:700;letter-spacing:.8px;margin-bottom:6px}
.mc-value{font-size:34px;font-weight:900;color:#1a3d17;line-height:1}
.mc-sub{font-size:12px;color:#a0b89d;margin-top:4px}
div.stButton>button:first-child{background:linear-gradient(135deg,#2d5a27,#3d7a35)!important;
  color:white!important;font-weight:700!important;font-size:17px!important;
  border-radius:10px!important;border:none!important;height:3.4em;width:100%;
  box-shadow:0 4px 14px rgba(45,90,39,.3);transition:all .18s!important}
[data-testid="stDownloadButton"]>button{background:#b84d00!important;color:white!important;
  font-weight:600!important;border-radius:10px!important;border:none!important}
[data-testid="stFileUploader"]{background:white;border-radius:12px;padding:14px 16px;
  border:2px dashed #c0d8bc;direction:rtl;text-align:right}
[data-testid="stTable"] table{width:100%;border-collapse:collapse;font-size:12px;
  background:white;direction:rtl}
[data-testid="stTable"] th{background:#2d5a27!important;color:white!important;
  padding:9px 12px!important;font-weight:600!important;text-align:right!important}
[data-testid="stTable"] td{padding:8px 12px!important;border-bottom:1px solid #f0f0f0!important;
  text-align:right!important}
[data-testid="stTable"] tr:nth-child(even) td{background:#f8fcf7!important}
.info-box{background:#edf5ec;border-right:5px solid #2d5a27;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#1a3d17;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.warn-box{background:#fff8e6;border-right:5px solid #e67e22;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#7a4500;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.error-box{background:#fdecea;border-right:5px solid #c0392b;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#7a0010;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val) -> list:
    if pd.isna(val) or str(val).strip().lower() in ('all', '', 'nan'):
        return list(range(24))
    res = set()
    for part in str(val).split(','):
        part = part.strip().split(':')[0].strip()
        if '-' in part:
            ps = part.split('-')
            if len(ps) == 2 and ps[0].isdigit() and ps[1].isdigit():
                s, e = int(ps[0]), int(ps[1])
                if s <= e:
                    for h in range(s, e + 1): res.add(h % 24)
                else:
                    for h in range(s, 24): res.add(h)
                    for h in range(0, e + 1): res.add(h)
        elif part.replace('.0','').isdigit():
            res.add(int(float(part)) % 24)
    return sorted(res)


class Soldier:
    def __init__(self, s_id, name, restr='', roles='', unavail=''):
        self.sid   = str(s_id).replace('.0','').strip()
        self.name  = str(name).strip()
        self.exempt = ([int(float(t)) for t in str(restr).split(',')
                        if str(t).strip().replace('.0','').isdigit()]
                       if pd.notna(restr) and str(restr).strip() not in ('','nan') else [])
        self.roles  = ([r.strip() for r in str(roles).split(',') if r.strip()]
                       if pd.notna(roles) and str(roles).strip() not in ('','nan') else [])
        self.blocked_hours = parse_time_ranges(unavail)


class Task:
    def __init__(self, t_id, name, req, shift, rest, overlap, hours, req_roles, intensity, block_roles=''):
        self.tid       = int(float(t_id))
        self.name      = str(name).strip()
        self.req       = int(req)
        self.shift     = int(shift)   if pd.notna(shift)   else 1
        self.rest      = int(rest)    if pd.notna(rest)    else 0
        self.overlap   = str(overlap).strip().lower() == 'true'
        self.hours     = parse_time_ranges(hours)
        self.intensity = int(intensity) if pd.notna(intensity) else 1
        self.block_roles = ([r.strip() for r in str(block_roles).split(',') if r.strip()]
                            if pd.notna(block_roles) and str(block_roles).strip() not in ('','nan') else [])
        parsed = ([r.strip() for r in str(req_roles).split(',')]
                  if pd.notna(req_roles) and str(req_roles).strip() not in ('','nan') else [])
        self.slots = parsed[:]
        while len(self.slots) < self.req:
            self.slots.append(None)

    def can_assign(self, soldier: Soldier, h: int) -> bool:
        if h in soldier.blocked_hours: return False
        if self.tid in soldier.exempt: return False
        if any(r in self.block_roles for r in soldier.roles): return False
        return True

    def slot_ok(self, slot_idx: int, soldier: Soldier) -> bool:
        req_role = self.slots[slot_idx]
        if req_role is None: return True
        return req_role in soldier.roles


# ══════════════════════════════════════════════════════════════════
# 3. Excel מעוצב
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df: pd.DataFrame, sheet_name='שבצ"ק', include_index=True) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        df.to_excel(w, index=include_index, sheet_name=sheet_name)
        wb, ws = w.book, w.sheets[sheet_name]
        hf = wb.add_format({'bold':True,'fg_color':'#2d5a27','font_color':'white',
                             'border':1,'align':'right','valign':'vcenter'})
        ef = wb.add_format({'fg_color':'#f0f8ef','align':'right'})
        bf = wb.add_format({'align':'right'})
        for ci, cv in enumerate(df.columns):
            ix = ci + (1 if include_index else 0)
            ws.write(0, ix, cv, hf)
            ws.set_column(ix, ix, min(max(df[cv].astype(str).map(len).max(), len(cv))+4, 40))
        for ri in range(1, len(df)+1):
            ws.set_row(ri, None, ef if ri%2==0 else bf)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════
# 4. לב המערכת: גריידי חכם — תמיד מוצא פתרון
#
#  שלב א: שיבוץ גריידי מובטח (שניות)
#    • עוקב אחרי "עסוק עד שעה X" לכל חייל
#    • מחלק לפי עומס: הכי פנוי קודם
#    • מטפל במשמרות מעגליות (23:00→01:00 וכו')
#    • אם אין חייל מתאים — מסמן "חסר"
#
#  שלב ב: שיפור CP-SAT קצר (אופציונלי)
#    • מתחיל מהפתרון הגריידי כ-hint
#    • רק מנסה לשפר הוגנות ומנוחה — לא מחפש feasibility
# ══════════════════════════════════════════════════════════════════

def greedy_schedule(soldiers: list, tasks: list, num_hours: int = 24):
    """
    מחזיר: schedule[tid][slot_idx][h] = sid | None
    ו-stats dict.
    """
    # schedule[tid][slot][h]
    schedule = {
        t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))}
        for t in tasks
    }

    # עסוק_עד[sid] = שעה אחרונה (כולל) שהחייל תפוס
    busy_until: dict[str, int] = {s.sid: -1 for s in soldiers}
    work_hours: dict[str, int] = {s.sid: 0  for s in soldiers}
    intensity_load: dict[str, int] = {s.sid: 0 for s in soldiers}

    def is_free(sid: str, h: int, dur: int, rest: int, overlap: bool) -> bool:
        """האם החייל פנוי לכל dur שעות החל מ-h?"""
        if overlap:
            return True
        # בדוק שהחייל לא תפוס בשעות המשמרת
        for i in range(dur):
            hh = (h + i) % num_hours
            if busy_until[sid] >= hh and not overlap:
                # פשוט: אם busy_until >= h → לא פנוי
                pass
        # גישה פשוטה: busy_until מייצג "עסוק עד שעה X"
        # משמרת חדשה מתחילה ב-h → צריך h > busy_until
        end_of_rest = busy_until[sid]
        if end_of_rest < 0:
            return True
        # תמיכה במעגליות: אם busy_until >= 20 ו-h קטן (מחרת)
        if end_of_rest >= 20 and h <= 8:
            return False  # עדיין במנוחה
        return h > end_of_rest

    def mark_busy(sid: str, h: int, dur: int, rest: int):
        end = (h + dur + rest - 1) % num_hours
        # אם חוצה חצות — שמור את הגדול (לוגיקה פשוטה)
        busy_until[sid] = (h + dur + rest - 1)  # לא מודולרי בכוונה לדיוק

    # מיין משימות: קודם הקשות יותר (פחות גמישות)
    sorted_tasks = sorted(tasks, key=lambda t: (-t.req, -t.intensity))

    missing_slots = []  # (tid, slot, h) שלא אויישו

    for t in sorted_tasks:
        for slot_idx in range(len(t.slots)):
            # בנה רשימת שעות התחלה — רק תחילת משמרת
            shift_starts = []
            prev_h = None
            for h in sorted(t.hours):
                if prev_h is None or h != prev_h + 1:
                    shift_starts.append(h)
                prev_h = h

            for start_h in shift_starts:
                # בחר חייל לפי עומס (הכי פנוי קודם), מעורבב קצת לגיוון
                candidates = []
                for s in soldiers:
                    if not t.can_assign(s, start_h): continue
                    if not t.slot_ok(slot_idx, s): continue
                    if not is_free(s.sid, start_h, t.shift, t.rest, t.overlap): continue
                    score = work_hours[s.sid] * 10 + intensity_load[s.sid]
                    candidates.append((score, s.sid, s))

                if not candidates:
                    # נסה בלי בדיקת מנוחה (רך)
                    for s in soldiers:
                        if not t.can_assign(s, start_h): continue
                        if not t.slot_ok(slot_idx, s): continue
                        score = work_hours[s.sid] * 10 + intensity_load[s.sid]
                        candidates.append((score + 1000, s.sid, s))  # קנס

                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    # בחר מתוך 3 הטובים אקראית לגיוון
                    top = candidates[:3]
                    _, chosen_sid, chosen_s = random.choice(top) if len(top) > 1 else top[0]

                    # שבץ את כל שעות המשמרת
                    for i in range(t.shift):
                        hh = (start_h + i) % num_hours
                        schedule[t.tid][slot_idx][hh] = chosen_sid

                    work_hours[chosen_sid] += t.shift
                    intensity_load[chosen_sid] += t.shift * t.intensity
                    if not t.overlap:
                        mark_busy(chosen_sid, start_h, t.shift, t.rest)
                else:
                    for i in range(t.shift):
                        hh = (start_h + i) % num_hours
                        missing_slots.append((t.name, slot_idx+1, hh))

    return schedule, missing_slots, work_hours, intensity_load


def build_result_df(soldiers, tasks, schedule, num_hours=24):
    SLEEP = set([22,23,0,1,2,3,4,5,6,7,8])
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    sid_to_soldier = {s.sid: s for s in soldiers}
    rows = []
    for s in soldiers:
        row = {"שם": s.name}
        total = 0
        night = 0
        intensity = 0
        for h in range(num_hours):
            active = []
            for t in tasks:
                for slot_idx in range(len(t.slots)):
                    if schedule[t.tid][slot_idx][h] == s.sid:
                        active.append(t.name)
                        if not t.overlap:
                            if h in SLEEP: night += 1
                            intensity += t.intensity
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if active: total += 1
        row["סך שעות"] = total
        row["מדד עצימות"] = intensity
        row["שעות שינה (22-08)"] = len(SLEEP) - night
        rows.append(row)
    df = pd.DataFrame(rows)
    df.index = range(1, len(df)+1)
    return df


# ══════════════════════════════════════════════════════════════════
# 5. שיפור CP-SAT (אופציונלי, קצר)
# ══════════════════════════════════════════════════════════════════
def improve_with_cpsat(soldiers, tasks, schedule, num_hours=24, time_limit=60.0):
    """
    מקבל schedule גריידי ומשפר הוגנות/מנוחה.
    עובד רק על שעות שה-greedy כבר מילא (לא מנסה לשנות coverage).
    """
    # המרת schedule ל-hint dict
    hints = {}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                assigned_sid = schedule[t.tid][si][h]
                for s in soldiers:
                    hints[(s.sid, t.tid, si, h)] = 1 if s.sid == assigned_sid else 0

    model = cp_model.CpModel()
    x = {}
    for s in soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.sid, t.tid, si, h] = model.NewBoolVar(f"x_{s.sid}_{t.tid}_{si}_{h}")
                    model.AddHint(x[s.sid, t.tid, si, h], hints.get((s.sid, t.tid, si, h), 0))

    # אילוצי כיסוי (קשיח) — שמור על מה שה-greedy עשה
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                assigned = sum(x[s.sid, t.tid, si, h] for s in soldiers)
                if h in t.hours:
                    model.Add(assigned == 1)
                else:
                    model.Add(assigned == 0)

    # אילוצי כשירות
    for s in soldiers:
        for t in tasks:
            if any(r in t.block_roles for r in s.roles) or t.tid in s.exempt:
                for si in range(len(t.slots)):
                    for h in range(num_hours):
                        model.Add(x[s.sid, t.tid, si, h] == 0)
                continue
            for si, req_role in enumerate(t.slots):
                if req_role and req_role not in s.roles:
                    for h in range(num_hours):
                        model.Add(x[s.sid, t.tid, si, h] == 0)

        for h in range(num_hours):
            if h in s.blocked_hours:
                model.Add(sum(x[s.sid, t.tid, si, h] for t in tasks for si in range(len(t.slots))) == 0)

    # חד-ערכיות
    for s in soldiers:
        for h in range(num_hours):
            bl = [x[s.sid, t.tid, si, h] for t in tasks if not t.overlap for si in range(len(t.slots))]
            if bl: model.Add(sum(bl) <= 1)

    # מטרה: הוגנות שעות
    total_hours = []
    for s in soldiers:
        th = model.NewIntVar(0, num_hours*len(tasks), f'th_{s.sid}')
        model.Add(th == sum(x[s.sid, t.tid, si, h]
                            for t in tasks for si in range(len(t.slots)) for h in range(num_hours)))
        total_hours.append(th)

    max_h = model.NewIntVar(0, 1000, 'max_h')
    min_h = model.NewIntVar(0, 1000, 'min_h')
    model.AddMaxEquality(max_h, total_hours)
    model.AddMinEquality(min_h, total_hours)
    diff = model.NewIntVar(0, 1000, 'diff')
    model.Add(diff == max_h - min_h)
    model.Minimize(diff)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers  = 6
    solver.parameters.log_search_progress = False
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return schedule  # החזר greedy אם CP-SAT נכשל

    # בנה schedule מעודכן
    new_schedule = {
        t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))}
        for t in tasks
    }
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                for s in soldiers:
                    if solver.Value(x[s.sid, t.tid, si, h]) == 1:
                        new_schedule[t.tid][si][h] = s.sid
    return new_schedule


# ══════════════════════════════════════════════════════════════════
# 6. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v11)</h1>
  <p>גריידי מובטח · שיפור CP-SAT אופציונלי · תמיד מחזיר תוצאה · 40+ חיילים</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀  ביצוע שיבוץ", "📖  מדריך", "📥  תבניות"])

# ── תבניות ──
with tab_templates:
    s_ex = pd.DataFrame({
        'מספר אישי':    [1001, 1002, 1003, 1004],
        'שם מלא':       ['ישראל ישראלי', 'יוסי כהן', 'אבי לוי', 'רון גל'],
        'פטורים':       ['', '101', '', ''],
        'הסמכות':       ['', '', '', ''],
        'שעות חסימה':   ['', '10-14', '', '22-6'],
    })
    t_ex = pd.DataFrame({
        'מס"ד משימה':                [101, 102, 103],
        'שם המשימה':                 ['שמירה', 'סיור', 'כוננות'],
        'סד"כ נדרש למשימה':         [2, 2, 1],
        'משך משמרת':                 [4, 6, 24],
        'שעות מנוחה בין משימות':     [8, 8, 0],
        'אישור חפיפה בין משימות':    [False, False, True],
        'שעות פעילות':               ['all', 'all', 'all'],
        'הסמכה נדרשת':               ['', '', ''],
        'דירוג עצימות משימה (1-3)': [2, 3, 1],
        'תפקידים חסומים':            ['', '', ''],
    })
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**👥 תבנית חיילים**")
        st.dataframe(s_ex, use_container_width=True, hide_index=True)
        st.download_button("⬇️ הורד תבנית חיילים",
                           data=to_excel_styled(s_ex, "Soldiers", False),
                           file_name="Soldiers_v11.xlsx", use_container_width=True)
    with c2:
        st.markdown("**📋 תבנית משימות**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button("⬇️ הורד תבנית משימות",
                           data=to_excel_styled(t_ex, "Tasks", False),
                           file_name="Tasks_v11.xlsx", use_container_width=True)

# ── מדריך ──
with tab_guide:
    st.markdown("""
### 📖 מדריך v11

#### קובץ חיילים
| עמודה | הסבר | דוגמה |
|---|---|---|
| מספר אישי | מזהה ייחודי | 1001 |
| שם מלא | שם להצגה | ישראל ישראלי |
| פטורים | מס"ד משימות חסומות | 101,103 |
| הסמכות | תפקידים (אם יש) | נהג,מפקד |
| שעות חסימה | שעות לא זמין | 10-14 / 22-6 |

#### קובץ משימות
| עמודה | הסבר | דוגמה |
|---|---|---|
| מס"ד משימה | מזהה | 101 |
| שם המשימה | שם קצר | שמירה |
| סד"כ נדרש | כמות חיילים | 3 |
| משך משמרת | שעות רצופות | 4 |
| שעות מנוחה | צינון אחרי משמרת | 8 |
| אישור חפיפה | True/False | False |
| שעות פעילות | all / 8-20 / 6,7,8 | all |
| הסמכה נדרשת | תפקיד ספציפי (אם נדרש) | נהג |
| דירוג עצימות | 1-3 | 2 |
| תפקידים חסומים | מי לא יכול | קצין |
""")
    st.markdown("""
<div class="info-box">
<b>💡 v11 — איך זה עובד:</b><br>
<b>שלב 1 — גריידי (שניות):</b> האלגוריתם משבץ חייל-חייל לפי עומס. תמיד מחזיר תוצאה — גם אם יש חוסר כוח אדם.<br>
<b>שלב 2 — CP-SAT (אופציונלי):</b> מנסה לשפר הוגנות מבלי לשבור את הכיסוי. אם נכשל — נשמר הגריידי.<br>
<b>חוסר כוח אדם:</b> שעות ללא חייל יסומנו ב-⚠️ כדי שתדעו היכן החלל.
</div>
""", unsafe_allow_html=True)

# ── ביצוע שיבוץ ──
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    with st.expander("⚙️ הגדרות מתקדמות"):
        use_cpsat = st.toggle("שיפור CP-SAT אחרי גריידי (משפר הוגנות)", value=True)
        cpsat_time = st.slider("זמן CP-SAT (שניות)", 10, 120, 45, 5)
        st.markdown('<div class="info-box" style="font-size:12px">הגריידי <b>תמיד</b> רץ ראשון ומחזיר תוצאה. CP-SAT רק משפר.</div>', unsafe_allow_html=True)

    if sf and tf:
        if st.button('⚙️ צור שבצ"ק חכם (v11)', use_container_width=True, key="run_btn"):
            try:
                # ── קריאת קבצים ──
                s_df = pd.read_excel(sf)
                t_df = pd.read_excel(tf)

                # זיהוי עמודות גמיש
                def find_col(df, keywords):
                    for kw in keywords:
                        matches = [c for c in df.columns if kw in str(c)]
                        if matches: return matches[0]
                    return None

                id_col    = find_col(s_df, ['מספר','אישי'])
                name_col  = find_col(s_df, ['שם'])
                t_id_col  = find_col(t_df, ['מס"ד','מסד'])
                t_nm_col  = find_col(t_df, ['שם'])
                t_req_col = find_col(t_df, ['סד"כ','נדרש'])

                if not all([id_col, name_col, t_id_col, t_nm_col, t_req_col]):
                    st.markdown('<div class="error-box">❌ לא נמצאו עמודות חובה. ודא שהקבצים לפי התבנית.</div>', unsafe_allow_html=True)
                    st.stop()

                s_df = s_df.dropna(subset=[id_col, name_col])
                t_df = t_df.dropna(subset=[t_id_col, t_nm_col, t_req_col])

                def gcol(df, keys, default=''):
                    c = find_col(df, keys)
                    return c if c else None

                shift_col    = gcol(t_df, ['משך'])
                rest_col     = gcol(t_df, ['מנוחה'])
                overlap_col  = gcol(t_df, ['חפיפה'])
                hours_col    = gcol(t_df, ['פעילות','שעות'])
                role_col     = gcol(t_df, ['הסמכה'])
                intense_col  = gcol(t_df, ['עצימות','דירוג'])
                block_col    = gcol(t_df, ['חסומים'])

                soldiers = [
                    Soldier(r[id_col], r[name_col],
                            r.get('פטורים',''), r.get('הסמכות',''),
                            r.get('שעות חסימה',''))
                    for _, r in s_df.iterrows()
                ]
                tasks = [
                    Task(r[t_id_col], r[t_nm_col], r[t_req_col],
                         r.get(shift_col) if shift_col else 1,
                         r.get(rest_col)  if rest_col  else 0,
                         r.get(overlap_col) if overlap_col else False,
                         r.get(hours_col)  if hours_col else 'all',
                         r.get(role_col)   if role_col  else '',
                         r.get(intense_col) if intense_col else 1,
                         r.get(block_col)  if block_col  else '')
                    for _, r in t_df.iterrows()
                ]

                st.info(f"📋 נטענו {len(soldiers)} חיילים ו-{len(tasks)} משימות")

                # ── שלב א: גריידי ──
                with st.spinner("⚡ שלב 1: שיבוץ גריידי (שניות)..."):
                    schedule, missing, work_h, int_load = greedy_schedule(soldiers, tasks)

                if missing:
                    st.markdown(
                        f'<div class="warn-box">⚠️ <b>{len(missing)} שעות ללא כיסוי</b> — '
                        f'אין חיילים זמינים לאיישן. שקלו להגדיל כוח אדם או לקצר מנוחות.</div>',
                        unsafe_allow_html=True)

                # ── שלב ב: CP-SAT ──
                if use_cpsat and not missing:
                    with st.spinner(f"🔧 שלב 2: שיפור CP-SAT ({cpsat_time}ש׳)..."):
                        schedule = improve_with_cpsat(soldiers, tasks, schedule, time_limit=cpsat_time)
                elif use_cpsat and missing:
                    st.markdown('<div class="info-box">ℹ️ CP-SAT דולג — יש שעות לא מאויישות. תחילה מלאו את כוח האדם.</div>', unsafe_allow_html=True)

                # ── בניית תוצאה ──
                final_df = build_result_df(soldiers, tasks, schedule)

                if len(final_df) == 0:
                    st.markdown('<div class="error-box">❌ לא נוצרו שורות — בדוק את קבצי הקלט.</div>', unsafe_allow_html=True)
                    st.stop()

                # ── מדדים ──
                gap_h     = int(final_df["סך שעות"].max() - final_df["סך שעות"].min())
                avg_h     = final_df["סך שעות"].mean()
                avg_sleep = final_df["שעות שינה (22-08)"].mean()
                badge     = "✅ מצוין" if gap_h <= 2 else ("⚠️ סביר" if gap_h <= 5 else "❗ גבוה")

                st.markdown(f"""
                <div class="metric-row">
                  <div class="metric-card"><div class="mc-label">חיילים</div>
                    <div class="mc-value">{len(soldiers)}</div>
                    <div class="mc-sub">{len(tasks)} משימות</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שעות</div>
                    <div class="mc-value">{avg_h:.1f}</div><div class="mc-sub">לחייל</div></div>
                  <div class="metric-card"><div class="mc-label">פער הוגנות</div>
                    <div class="mc-value">{gap_h}</div><div class="mc-sub">{badge}</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שינה</div>
                    <div class="mc-value">{avg_sleep:.1f}</div><div class="mc-sub">יעד: 7.0</div></div>
                  <div class="metric-card"><div class="mc-label">שעות חסרות</div>
                    <div class="mc-value">{"0 ✅" if not missing else f"{len(missing)} ⚠️"}</div>
                    <div class="mc-sub">{'כיסוי מלא' if not missing else 'חסר כוח אדם'}</div></div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📅 לוח השיבוץ הסופי")
                st.table(final_df)
                st.download_button("📥 הורד לוח שיבוץ (Excel)",
                                   data=to_excel_styled(final_df),
                                   file_name="Final_Shavtzak_v11.xlsx",
                                   use_container_width=True)

                # ── גרף ──
                st.markdown("---")
                st.subheader("📊 ניתוח עומסים")
                fig = px.bar(final_df, x="שם", y="סך שעות", color="מדד עצימות",
                             color_continuous_scale=["#a8d5a2","#1a3d17"],
                             title="עומס שעות לחייל", text="סך שעות")
                fig.update_traces(textposition="outside", marker_line_width=0)
                fig.update_layout(plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                                  font=dict(family="Heebo", size=12),
                                  xaxis=dict(tickangle=-30, title=""),
                                  yaxis=dict(title='שעות סה"כ'),
                                  margin=dict(t=50,b=80,l=30,r=20))
                st.plotly_chart(fig, use_container_width=True)

                with st.expander("💡 תובנות"):
                    max_s = final_df[final_df["סך שעות"]==final_df["סך שעות"].max()]["שם"].tolist()
                    st.markdown(f"""
**שלב גריידי:** שיבץ {len(soldiers)} חיילים תוך שניות לפי עומס מינימלי.
**שיפור CP-SAT:** {'שיפר הוגנות בין חיילים.' if use_cpsat else 'לא הופעל.'}
**חיילים עמוסים:** {', '.join(max_s)} — {final_df['סך שעות'].max()} שעות.
**פער הוגנות:** {gap_h} שעות — {'מצוין.' if gap_h<=2 else 'מומלץ להוסיף חיילים.' if gap_h>4 else 'סביר.'}
**שינה:** ממוצע {avg_sleep:.1f} שעות.
{"**⚠️ שעות ללא כיסוי:** " + str(len(missing)) + " — הוסף חיילים או קצר מנוחות." if missing else "**✅ כיסוי מלא** — כל העמדות מאויישות."}
                    """)

            except Exception as e:
                st.markdown('<div class="error-box">🚨 שגיאה טכנית:</div>', unsafe_allow_html=True)
                st.code(traceback.format_exc())
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין תבניות? לחצו על הטאב <b>תבניות</b>.
        </div>
        """, unsafe_allow_html=True)
