
import io
import re
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
import traceback

# ══════════════════════════════════════════════════════════════════
# 1. עיצוב ומשתני Session
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title='שבצ"ק חכם', page_icon="🪖", layout="wide")

if "directives" not in st.session_state:
    st.session_state.directives = []
if "adhoc_tasks" not in st.session_state:
    st.session_state.adhoc_tasks = []

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
.adhoc-box{background:#f0f4ff;border-right:5px solid #3a5bc7;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#1a2d7a;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.directive-box{background:#fdf6ff;border-right:5px solid #8b5cf6;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#3b0764;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.rec-box{background:#fffbeb;border-right:5px solid #f59e0b;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#78350f;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
.infeasible-box{background:#fff0f0;border-right:5px solid #e53e3e;padding:14px 18px;margin:14px 0;
  font-size:14px;color:#63171b;line-height:1.8;direction:rtl;text-align:right;border-radius:0 10px 10px 0}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val, is_task=True):
    if pd.isna(val) or str(val).strip().lower() in ('all', '', 'nan'):
        return list(range(24)) if is_task else []
    def to_hour(s: str) -> int:
        s = s.strip()
        if ':' in s: return int(s.split(':')[0])
        return int(float(s.replace('.0', '') or '0'))
    res = set()
    for part in str(val).split(','):
        part = part.strip()
        if not part or part.lower() == 'nan': continue
        if '-' in part:
            ps = part.split('-')
            start_s, end_s = ps[0].strip(), ps[-1].strip()
            if start_s and end_s:
                s, e = to_hour(start_s), to_hour(end_s)
                if s <= e:
                    for h in range(s, e + 1): res.add(h % 24)
                else:
                    for h in range(s, 24): res.add(h)
                    for h in range(0, e + 1): res.add(h)
        else:
            try: res.add(to_hour(part) % 24)
            except (ValueError, IndexError): pass
    return sorted(res)

DUMMY_SID        = "__DUMMY__"
_ADHOC_TID_START = 9000

class Soldier:
    def __init__(self, s_id, name, restr='', roles='', unavail='', is_dummy=False):
        self.sid      = str(s_id).replace('.0', '').strip()
        self.name     = str(name).strip()
        self.is_dummy = is_dummy
        self.exempt   = ([int(float(t)) for t in str(restr).split(',') if str(t).strip().replace('.0', '').isdigit()] if pd.notna(restr) and str(restr).strip() not in ('', 'nan') else [])
        self.roles = ([r.strip() for r in str(roles).split(',') if r.strip()] if pd.notna(roles) and str(roles).strip() not in ('', 'nan') else [])
        self.blocked_hours = [] if is_dummy else parse_time_ranges(unavail, is_task=False)

    def can_assign(self, task, shift_hours):
        if self.is_dummy: return True
        if any(h in self.blocked_hours for h in shift_hours): return False
        if task.tid in self.exempt: return False
        if any(r in task.block_roles for r in self.roles): return False
        return True

class Task:
    def __init__(self, t_id, name, req, shift, rest, overlap, hours, req_roles, intensity, block_roles='', is_adhoc=False):
        self.tid       = int(float(t_id))
        self.name      = str(name).strip()
        self.req       = int(req)
        self.shift     = int(shift)   if pd.notna(shift)   else 1
        self.rest      = int(rest)    if pd.notna(rest)    else 0
        self.overlap   = str(overlap).strip().lower() == 'true'
        self.hours     = parse_time_ranges(hours)
        self.intensity = int(intensity) if pd.notna(intensity) else 1
        self.is_adhoc  = is_adhoc
        self.block_roles = ([r.strip() for r in str(block_roles).split(',') if r.strip()] if pd.notna(block_roles) and str(block_roles).strip() not in ('', 'nan') else [])
        parsed = ([r.strip() for r in str(req_roles).split(',')] if pd.notna(req_roles) and str(req_roles).strip() not in ('', 'nan') else [])
        self.slots = parsed[:]
        while len(self.slots) < self.req: self.slots.append(None)

    def slot_ok(self, slot_idx, soldier):
        if soldier.is_dummy: return True
        req_role = self.slots[slot_idx]
        if req_role is None: return True
        return req_role in soldier.roles

    def get_shift_starts(self):
        if not self.hours: return []
        hours_set, sorted_hours = set(self.hours), sorted(self.hours)
        starts, i = [], 0
        while i < len(sorted_hours):
            start = sorted_hours[i]
            if all((start + k) % 24 in hours_set for k in range(self.shift)):
                starts.append(start)
                i += self.shift
            else: i += 1
        return starts

# ══════════════════════════════════════════════════════════════════
# 3. הנחיות מפקד — מבנה נתונים ופרסור
# ══════════════════════════════════════════════════════════════════
class Directive:
    def __init__(self, soldier_name, directive_type, task_name, hours):
        self.soldier_name   = soldier_name.strip()
        self.directive_type = directive_type
        self.task_name      = task_name.strip()
        self.hours          = hours

    def __repr__(self):
        h_str = f" בשעות {self.hours}" if self.hours else " (כל היום)"
        arrow = "✅ חייב" if self.directive_type == 'must' else "❌ אסור"
        return f"{arrow}: {self.soldier_name} ← {self.task_name}{h_str}"

def parse_free_text_directive(text, soldiers, tasks):
    text = text.strip()
    soldier_obj = None
    for s in soldiers:
        if s.is_dummy: continue
        first_name = s.name.split()[0] if s.name.split() else s.name
        if (first_name in text or s.name in text):
            soldier_obj = s
            break
    if soldier_obj is None:
        names = [s.name for s in soldiers if not s.is_dummy]
        return None, f"לא זוהה שם חייל בטקסט. חיילים קיימים: {', '.join(names[:8])}"

    must_keywords     = ['חייב', 'חייבת', 'יבצע', 'תבצע', 'ישרת', 'תשרת', 'ישמור', 'תשמור', 'must', 'חייב לשמור', 'חייב לבצע']
    must_not_keywords = ['אסור', 'לא יכול', 'לא יעבוד', 'לא תעבוד', 'לא ישרת', 'לא ישמור', 'must not', 'cannot', 'לא יבצע']
    directive_type = None
    for kw in must_not_keywords:
        if kw in text: directive_type = 'must_not'; break
    if directive_type is None:
        for kw in must_keywords:
            if kw in text: directive_type = 'must'; break
    if directive_type is None:
        return None, "לא זוהתה הנחיה. השתמש במילים: חייב / אסור / לא יכול"

    task_obj = None
    for t in tasks:
        if t.name in text:
            task_obj = t
            break
    if task_obj is None:
        task_names = [t.name for t in tasks]
        return None, f"לא זוהתה משימה. משימות קיימות: {', '.join(task_names)}"

    hours = []
    m = re.search(r'בין\s+(\d{1,2})(?::\d{2})?\s+ל[-–]\s*(\d{1,2})(?::\d{2})?', text)
    if m:
        sh, eh = int(m.group(1)), int(m.group(2))
        hours  = list(range(sh, eh + 1)) if sh <= eh else (list(range(sh, 24)) + list(range(0, eh + 1)))
    else:
        m2 = re.search(r'בשעה\s+(\d{1,2})(?::\d{2})?', text)
        if m2: hours = [int(m2.group(1)) % 24]
        else:
            m3 = re.search(r'(\d{1,2})(?::\d{2})?\s*[-–]\s*(\d{1,2})(?::\d{2})?', text)
            if m3:
                sh, eh = int(m3.group(1)), int(m3.group(2))
                hours  = list(range(sh, eh + 1)) if sh <= eh else (list(range(sh, 24)) + list(range(0, eh + 1)))

    return Directive(soldier_obj.name, directive_type, task_obj.name, hours), ""

# ══════════════════════════════════════════════════════════════════
# 4. Pre-flight Validation
# ══════════════════════════════════════════════════════════════════
def validate_inputs_and_directives(soldiers, tasks, directives):
    errors = []
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    name_to_soldier = {s.name: s for s in real_soldiers}
    name_to_task    = {t.name: t for t in tasks}

    for d in directives:
        if d.soldier_name == "__RAW__": continue
        if d.soldier_name not in name_to_soldier:
            errors.append(f'🔸 שם החייל <b>"{d.soldier_name}"</b> בהנחיית המפקד לא נמצא בקובץ החיילים.')
        if d.task_name not in name_to_task:
            errors.append(f'🔸 שם המשימה <b>"{d.task_name}"</b> בהנחיית המפקד לא נמצא בקובץ המשימות.')

    must_set, must_not_set = {}, {}
    for d in directives:
        if d.soldier_name == "__RAW__": continue
        t_obj = name_to_task.get(d.task_name)
        if t_obj is None: continue
        hours_to_check = d.hours if d.hours else t_obj.hours
        for h in hours_to_check:
            key = (d.soldier_name, d.task_name, h)
            if d.directive_type == 'must': must_set[key] = d
            else: must_not_set[key] = d

    for key in must_set:
        if key in must_not_set:
            s_name, t_name, h = key
            errors.append(f'🔸 סתירה: <b>"{s_name}"</b> מסומן גם כחייב וגם כאסור למשימה "{t_name}" בשעה {h:02d}:00.')

    return len(errors) == 0, errors

# ══════════════════════════════════════════════════════════════════
# 5. אבחון כשל מודל
# ══════════════════════════════════════════════════════════════════
def diagnose_infeasible_model(soldiers, tasks, directives):
    real_soldiers   = [s for s in soldiers if not s.is_dummy]
    name_to_soldier = {s.name: s for s in real_soldiers}
    name_to_task    = {t.name: t for t in tasks}
    insights        = []

    for d in directives:
        if d.soldier_name == "__RAW__" or d.directive_type != 'must': continue
        s_obj, t_obj = name_to_soldier.get(d.soldier_name), name_to_task.get(d.task_name)
        if s_obj is None or t_obj is None: continue

        hours_to_check = d.hours if d.hours else t_obj.hours
        for h in hours_to_check:
            if h in s_obj.blocked_hours:
                insights.append(f'🔸 <b>"{s_obj.name}"</b> חייב "{t_obj.name}" ב-{h:02d}:00 אך <b>השעה חסומה לו אישית</b>.')
            if t_obj.tid in s_obj.exempt:
                insights.append(f'🔸 <b>"{s_obj.name}"</b> חייב "{t_obj.name}" אך הוא <b>פטור ממשימה זו</b>.')
            if any(r in t_obj.block_roles for r in s_obj.roles):
                insights.append(f'🔸 <b>"{s_obj.name}"</b> חייב "{t_obj.name}" אך תפקידו <b>חסום למשימה זו</b>.')
            if t_obj.slots and t_obj.slots[0] and t_obj.slots[0] not in s_obj.roles:
                insights.append(f'🔸 <b>"{s_obj.name}"</b> חייב "{t_obj.name}" אך <b>אין לו את ההסמכה הנדרשת</b>.')
            if h not in t_obj.hours:
                insights.append(f'🔸 <b>"{s_obj.name}"</b> חייב "{t_obj.name}" ב-{h:02d}:00 אך המשימה <b>לא פעילה בשעה זו</b>.')

    if not insights:
        insights.append('🔸 המערכת אינה מסוגלת לכסות את כל המשימות עם האילוצים הקיימים. נסה להפחית חסימות או להוסיף חיילים זמינים.')
    return insights

# ══════════════════════════════════════════════════════════════════
# 6. Excel 
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df, sheet_name='שבצ"ק', include_index=True):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        df.to_excel(w, index=include_index, sheet_name=sheet_name)
        wb, ws = w.book, w.sheets[sheet_name]
        hf = wb.add_format({'bold':True,'fg_color':'#2d5a27','font_color':'white','border':1,'align':'right'})
        ef = wb.add_format({'fg_color':'#f0f8ef','align':'right'})
        bf = wb.add_format({'align':'right'})
        for ci, cv in enumerate(df.columns):
            ix = ci + (1 if include_index else 0)
            ws.write(0, ix, cv, hf)
            ws.set_column(ix, ix, min(max(df[cv].astype(str).map(len).max(), len(cv))+4, 40))
        for ri in range(1, len(df)+1):
            ws.set_row(ri, None, ef if ri%2==0 else bf)
    return out.getvalue()

def to_excel_task_view(soldiers, tasks, schedule):
    df = build_task_df(soldiers, tasks, schedule)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        df.to_excel(w, index=True, sheet_name='לפי משימות')
        wb, ws = w.book, w.sheets['לפי משימות']
        hf   = wb.add_format({'bold':True,'fg_color':'#1a3d17','font_color':'white','border':1,'align':'center'})
        idx_f= wb.add_format({'bold':True,'fg_color':'#2d5a27','font_color':'white','align':'center'})
        ok_f = wb.add_format({'fg_color':'#edf5ec','align':'right'})
        wf   = wb.add_format({'fg_color':'#fdecea','font_color':'#c0392b','bold':True,'align':'right'})
        ef   = wb.add_format({'fg_color':'#f9f9f9','font_color':'#bbbbbb','align':'center'})
        ws.write(0, 0, 'שעה', idx_f)
        ws.set_column(0, 0, 8)
        for ci, col in enumerate(df.columns):
            ws.write(0, ci+1, col, hf)
            ws.set_column(ci+1, ci+1, 22)
        for ri, (idx_val, row) in enumerate(df.iterrows()):
            ws.write(ri+1, 0, idx_val, idx_f)
            for ci, col in enumerate(df.columns):
                val = row[col]
                fmt = wf if '⚠️' in str(val) else (ef if val == '—' else ok_f)
                ws.write(ri+1, ci+1, val, fmt)
            ws.set_row(ri+1, 18)
    return out.getvalue()

# ══════════════════════════════════════════════════════════════════
# 7. אלגוריתמיקה
# ══════════════════════════════════════════════════════════════════
def greedy_schedule(soldiers, tasks, num_hours=24):
    schedule = {t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))} for t in tasks}
    assignments = {s.sid: {h: [] for h in range(num_hours)} for s in soldiers}
    assignments[DUMMY_SID] = {h: [] for h in range(num_hours)}
    work_hours = {s.sid: 0 for s in soldiers}; work_hours[DUMMY_SID] = 0
    intensity_load = {s.sid: 0 for s in soldiers}; intensity_load[DUMMY_SID] = 0
    dummy_slots = []

    def is_free(sid, shift_hours, new_task):
        for h in shift_hours:
            for ex in assignments[sid][h]:
                if not new_task.overlap: return False
                if not ex.overlap: return False
        return True

    sorted_tasks = sorted(tasks, key=lambda t: (1 if t.overlap else 0, -t.req, -t.intensity, -t.shift))

    for t in sorted_tasks:
        for slot_idx in range(len(t.slots)):
            for start_h in t.get_shift_starts():
                shift_hours = [(start_h + i) % num_hours for i in range(t.shift)]
                candidates  = []
                for s in soldiers:
                    if s.is_dummy: continue
                    if not s.can_assign(t, shift_hours): continue
                    if not t.slot_ok(slot_idx, s): continue
                    if not is_free(s.sid, shift_hours, t): continue
                    candidates.append((work_hours[s.sid]*10 + intensity_load[s.sid], s.sid, s))
                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    chosen_sid, chosen_dummy = candidates[0][1], False
                else:
                    chosen_sid, chosen_dummy = DUMMY_SID, True
                for h in shift_hours: schedule[t.tid][slot_idx][h] = chosen_sid
                if chosen_dummy: dummy_slots.append((t.name, slot_idx+1, start_h, shift_hours[-1]))
                else:
                    work_hours[chosen_sid] += t.shift
                    intensity_load[chosen_sid] += t.shift * t.intensity
                    for h in shift_hours: assignments[chosen_sid][h].append(t)
    return schedule, dummy_slots, work_hours, intensity_load


def improve_with_cpsat(soldiers, tasks, schedule, directives, num_hours=24, time_limit=60.0):
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    model = cp_model.CpModel()
    name_to_soldier, name_to_task = {s.name: s for s in real_soldiers}, {t.name: t for t in tasks}

    x = {}
    for s in real_soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.sid, t.tid, si, h] = model.NewBoolVar(f"x_{s.sid}_{t.tid}_{si}_{h}")
                    model.AddHint(x[s.sid, t.tid, si, h], 1 if schedule[t.tid][si][h] == s.sid else 0)

    dummy_vars = {}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in t.hours:
                dv = model.NewBoolVar(f"d_{t.tid}_{si}_{h}")
                dummy_vars[t.tid, si, h] = dv
                model.AddHint(dv, 1 if schedule[t.tid][si][h] == DUMMY_SID else 0)

    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h in t.hours: model.Add(sum(x[s.sid, t.tid, si, h] for s in real_soldiers) + dummy_vars[t.tid, si, h] == 1)
                else:
                    for s in real_soldiers: model.Add(x[s.sid, t.tid, si, h] == 0)

    for s in real_soldiers:
        for t in tasks:
            task_blocked = t.tid in s.exempt or any(r in t.block_roles for r in s.roles)
            for si in range(len(t.slots)):
                role_blocked = (t.slots[si] is not None and t.slots[si] not in s.roles)
                for h in range(num_hours):
                    if task_blocked or role_blocked or h in s.blocked_hours:
                        model.Add(x[s.sid, t.tid, si, h] == 0)

    for s in real_soldiers:
        for h in range(num_hours):
            for t1 in tasks:
                if t1.overlap: continue
                for si1 in range(len(t1.slots)):
                    for t2 in tasks:
                        if t2.tid == t1.tid: continue
                        for si2 in range(len(t2.slots)):
                            model.Add(x[s.sid,t1.tid,si1,h] + x[s.sid,t2.tid,si2,h] <= 1)

    for s in real_soldiers:
        for t in tasks:
            for h in t.hours:
                if len(t.slots) > 1: model.Add(sum(x[s.sid,t.tid,si,h] for si in range(len(t.slots))) <= 1)

    directive_errors = []
    for d in directives:
        s_obj, t_obj = name_to_soldier.get(d.soldier_name), name_to_task.get(d.task_name)
        if not s_obj or not t_obj:
            directive_errors.append(f"הנחיה לא הוחלה: {d}")
            continue
        hours_to_apply = d.hours if d.hours else t_obj.hours
        for h in hours_to_apply:
            if h not in t_obj.hours: continue
            for si in range(len(t_obj.slots)):
                if d.directive_type == 'must':
                    if si == 0: model.Add(x[s_obj.sid, t_obj.tid, si, h] == 1)
                elif d.directive_type == 'must_not': model.Add(x[s_obj.sid, t_obj.tid, si, h] == 0)

    penalties = []
    for t in tasks:
        for si in range(len(t.slots)):
            for h in t.hours: penalties.append(dummy_vars[t.tid, si, h] * 100_000)

    NIGHT_HOURS = sorted(set(range(22, 24)) | set(range(0, 9)))
    for s in real_soldiers:
        nv = [x[s.sid,t.tid,si,h] for t in tasks if not t.overlap for si in range(len(t.slots)) for h in NIGHT_HOURS]
        if nv:
            nt, ne = model.NewIntVar(0, 500, f'nt_{s.sid}'), model.NewIntVar(0, 500, f'ne_{s.sid}')
            model.Add(nt == sum(nv)); model.Add(ne >= nt - 4); model.Add(ne >= 0)
            penalties.append(ne * 800)

    for t in tasks:
        for start_h in t.get_shift_starts():
            shift_hours = [(start_h+k)%num_hours for k in range(t.shift)]
            if len(shift_hours) < 2: continue
            for si in range(len(t.slots)):
                for s in real_soldiers:
                    for k in range(len(shift_hours)-1):
                        h_now, h_next = shift_hours[k], shift_hours[k+1]
                        sv = model.NewBoolVar(f'sp_{s.sid}_{t.tid}_{si}_{h_now}')
                        model.Add(x[s.sid,t.tid,si,h_now]-x[s.sid,t.tid,si,h_next]>=1).OnlyEnforceIf(sv)
                        model.Add(x[s.sid,t.tid,si,h_now]-x[s.sid,t.tid,si,h_next]<=0).OnlyEnforceIf(sv.Not())
                        penalties.append(sv * 400)

    th_vars = []
    for s in real_soldiers:
        th = model.NewIntVar(0, num_hours*len(tasks), f'th_{s.sid}')
        model.Add(th == sum(x[s.sid,t.tid,si,h] for t in tasks for si in range(len(t.slots)) for h in range(num_hours)))
        th_vars.append(th)
    max_h, min_h, diff = model.NewIntVar(0, 1000, 'max_h'), model.NewIntVar(0, 1000, 'min_h'), model.NewIntVar(0, 1000, 'diff')
    model.AddMaxEquality(max_h, th_vars); model.AddMinEquality(min_h, th_vars); model.Add(diff == max_h - min_h)
    penalties.append(diff * 100)

    for s in real_soldiers:
        for t in tasks:
            if t.overlap or t.rest == 0: continue
            for start_h in t.get_shift_starts():
                end_h = (start_h + t.shift - 1) % num_hours
                in_s = model.NewBoolVar(f'ins_{s.sid}_{t.tid}_{start_h}')
                model.Add(sum(x[s.sid,t.tid,si,start_h] for si in range(len(t.slots)))>=1).OnlyEnforceIf(in_s)
                model.Add(sum(x[s.sid,t.tid,si,start_h] for si in range(len(t.slots)))==0).OnlyEnforceIf(in_s.Not())
                for r in range(1, t.rest+1):
                    rest_h = (end_h+r) % num_hours
                    rw = sum(x[s.sid,t2.tid,si2,rest_h] for t2 in tasks if not t2.overlap for si2 in range(len(t2.slots)))
                    ir, viol = model.NewBoolVar(f'ir_{s.sid}_{t.tid}_{start_h}_{r}'), model.NewBoolVar(f'rv_{s.sid}_{t.tid}_{start_h}_{r}')
                    model.Add(rw>=1).OnlyEnforceIf(ir); model.Add(rw==0).OnlyEnforceIf(ir.Not())
                    model.AddBoolAnd([in_s, ir]).OnlyEnforceIf(viol); model.AddBoolOr([in_s.Not(), ir.Not()]).OnlyEnforceIf(viol.Not())
                    penalties.append(viol * 50)

    tp = model.NewIntVar(0, 100_000_000, 'tp')
    model.Add(tp == sum(penalties))
    model.Minimize(tp)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = 8
    solver.parameters.log_search_progress = False
    solver.parameters.relative_gap_limit = 0.03
    status = solver.Solve(model)

    if status == cp_model.INFEASIBLE: return schedule, directive_errors, True
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return schedule, directive_errors, False

    new_sch = {t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))} for t in tasks}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in t.hours:
                if solver.Value(dummy_vars[t.tid, si, h]) == 1: new_sch[t.tid][si][h] = DUMMY_SID
                else:
                    for s in real_soldiers:
                        if solver.Value(x[s.sid, t.tid, si, h]) == 1:
                            new_sch[t.tid][si][h] = s.sid
                            break
    return new_sch, directive_errors, False

# ══════════════════════════════════════════════════════════════════
# 8. אבחון חוסרים והמלצות מפקד
# ══════════════════════════════════════════════════════════════════
def diagnose_dummy_slots(soldiers, tasks, schedule, num_hours=24):
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    recommendations, seen = [], set()

    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if schedule[t.tid][si][h] != DUMMY_SID: continue
                if (t.tid, si, h) in seen: continue
                seen.add((t.tid, si, h))

                blocked_by_role, blocked_by_exemption, blocked_by_hours, wrong_role, potentially_free = [], [], [], [], []
                for s in real_soldiers:
                    if t.tid in s.exempt: blocked_by_exemption.append(s.name)
                    elif any(r in t.block_roles for r in s.roles): blocked_by_role.append(s.name)
                    elif h in s.blocked_hours: blocked_by_hours.append(s.name)
                    elif t.slots[si] and t.slots[si] not in s.roles: wrong_role.append(s.name)
                    else: potentially_free.append(s.name)

                req_role_name = t.slots[si]
                if len(real_soldiers) == 0: reason, rec = "אין חיילים במערכת", "הוסף חיילים לקובץ"
                elif req_role_name and not [s for s in real_soldiers if req_role_name in s.roles]:
                    reason, rec = f"אין חייל עם הסמכה '{req_role_name}'", f"הוסף חייל עם הסמכה: {req_role_name}"
                elif req_role_name and len(wrong_role) == len(real_soldiers) - len(blocked_by_exemption) - len(blocked_by_role):
                    reason, rec = f"כל בעלי ההסמכה '{req_role_name}' עסוקים", "הוסף עוד חייל בעל הסמכה זו"
                elif len(blocked_by_hours) == len(real_soldiers):
                    reason, rec = f"כל החיילים חסומים ב-{h:02d}:00", "בדוק שעות חסימה אישיות"
                elif potentially_free:
                    reason, rec = f"חיילים זמינים תיאורטית אך חסומים עקב מנוחה/עומס", "קצר שעות מנוחה במשימות אחרות"
                else:
                    reason, rec = "כל החיילים עסוקים/חסומים", "הוסף חייל נוסף למצבה"

                recommendations.append({"משימה": t.name, "שעה": f"{h:02d}:00", "סלוט": si+1, "סיבה": reason, "המלצה": rec})

    merged = {}
    for r in recommendations:
        key = (r["משימה"], r["סיבה"])
        if key not in merged: merged[key] = {**r, "_hours": [r["שעה"]]}
        else: merged[key]["_hours"].append(r["שעה"])

    final = []
    for rec in merged.values():
        hours_list = rec.pop("_hours")
        rec["שעות"] = ", ".join(hours_list) if len(hours_list)<=3 else f"{hours_list[0]}–{hours_list[-1]} ({len(hours_list)} שעות)"
        del rec["שעה"]
        final.append(rec)
    return final

# ══════════════════════════════════════════════════════════════════
# 9. DataFrames
# ══════════════════════════════════════════════════════════════════
def build_result_df(soldiers, tasks, schedule, num_hours=24):
    SLEEP, hour_labels, rows = set(range(22,24)) | set(range(0,9)), [f"{h:02d}:00" for h in range(num_hours)], []
    for s in soldiers:
        if s.is_dummy: continue
        row, total, night, intensity = {"שם": s.name}, 0, 0, 0
        for h in range(num_hours):
            active = []
            for t in tasks:
                for si in range(len(t.slots)):
                    if schedule[t.tid][si][h] == s.sid:
                        active.append(t.name)
                        if not t.overlap:
                            if h in SLEEP: night += 1
                            intensity += t.intensity
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if active: total += 1
        row["סך שעות"], row["מדד עצימות"], row["שעות שינה (22-08)"] = total, intensity, len(SLEEP) - night
        rows.append(row)

    dummy_row = {"שם": "⚠️ חוסר כוח אדם"}
    for h in range(num_hours):
        d = [t.name for t in tasks for si in range(len(t.slots)) if schedule[t.tid][si][h] == DUMMY_SID]
        dummy_row[hour_labels[h]] = "❌ " + " + ".join(d) if d else "—"
    dummy_row["סך שעות"] = sum(1 for t in tasks for si in range(len(t.slots)) for h in range(num_hours) if schedule[t.tid][si][h] == DUMMY_SID)
    dummy_row["מדד עצימות"], dummy_row["שעות שינה (22-08)"] = 0, 0

    df = pd.DataFrame(rows)
    if any(schedule[t.tid][si][h] == DUMMY_SID for t in tasks for si in range(len(t.slots)) for h in range(num_hours)):
        df = pd.concat([df, pd.DataFrame([dummy_row])], ignore_index=True)
    df.index = range(1, len(df)+1)
    return df

def build_task_df(soldiers, tasks, schedule, num_hours=24):
    sid_to_name, hour_labels, rows = {s.sid: s.name for s in soldiers if not s.is_dummy}, [f"{h:02d}:00" for h in range(num_hours)], []
    for h in range(num_hours):
        row = {"שעה": hour_labels[h]}
        for t in tasks:
            if h not in t.hours: row[t.name] = "—"; continue
            names, has_dummy = [], False
            for si in range(len(t.slots)):
                val = schedule[t.tid][si][h]
                if val == DUMMY_SID: has_dummy = True
                elif val is not None: names.append(sid_to_name.get(val, val))
            if has_dummy: row[t.name] = "⚠️ חוסר" + (" + " + ", ".join(names) if names else "")
            elif names: row[t.name] = ", ".join(names)
            else: row[t.name] = "—"
        rows.append(row)
    df = pd.DataFrame(rows)
    df.set_index("שעה", inplace=True)
    return df

# ══════════════════════════════════════════════════════════════════
# 10. ממשק משתמש הראשי
# ══════════════════════════════════════════════════════════════════
try:
    st.markdown("""
    <div class="app-header">
      <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה </h1>
      <p> · מהיר · יעיל ·  אפקטיבי  </p>
    </div>
    """, unsafe_allow_html=True)

    tab_run, tab_guide, tab_templates = st.tabs(["🚀  ביצוע שיבוץ", "📖  מדריך", "📥  תבניות"])

    # ── תבניות ──────────────────────────────────────────────────────
    with tab_templates:
        s_ex = pd.DataFrame({
            'מספר אישי': [1001,1002,1003,1004], 'שם מלא': ['ישראל ישראלי','יוסי כהן','אבי לוי','רון גל'],
            'פטורים': ['','101','',''], 'הסמכות': ['','','',''], 'שעות חסימה': ['','10-14','','22-6'],
        })
        t_ex = pd.DataFrame({
            'מס"ד משימה': [101,102,103], 'שם המשימה': ['שמירה','סיור','כוננות'],
            'סד"כ נדרש למשימה': [2,2,1], 'משך משמרת': [4,6,24],
            'שעות מנוחה מינימליות בין משימות': [8,8,0], 'אישור חפיפה בין משימות': [False,False,True],
            'שעות פעילות': ['all','all','all'], 'הסמכה נדרשת': ['','',''],
            'דירוג עצימות משימה (1-3)': [2,3,1], 'תפקידים חסומים': ['','',''],
        })
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**👥 תבנית חיילים**")
            st.dataframe(s_ex, use_container_width=True, hide_index=True)
            st.download_button("⬇️ הורד", data=to_excel_styled(s_ex,"Soldiers",False),
                               file_name="Soldiers_v22.xlsx", use_container_width=True)
        with c2:
            st.markdown("**📋 תבנית משימות**")
            st.dataframe(t_ex, use_container_width=True, hide_index=True)
            st.download_button("⬇️ הורד", data=to_excel_styled(t_ex,"Tasks",False),
                               file_name="Tasks_v22.xlsx", use_container_width=True)

 # ── מדריך ───────────────────────────────────────────────────────
    with tab_guide:
        st.markdown("### 📖 מדריך מקיף למשתמש — שבצ\"ק חכם")
        st.markdown("מערכת זו נועדה לחסוך זמן ולייצר שיבוץ הוגן, מדויק וללא טעויות אנוש. עקבו אחר ההנחיות למילוי הקבצים:")
        
        st.markdown("#### 👥 1. קובץ חיילים (מצבת כוח אדם)")
        st.markdown("""
| עמודה | הסבר | דוגמה למילוי |
| :--- | :--- | :--- |
| **מספר אישי / שם מלא** | שדות חובה לזיהוי החייל במערכת. | `1001`, `אריאל כהן` |
| **פטורים** | מס"ד (מספרי) המשימות שהחייל פטור מהן לחלוטין. | `101`, `101, 103` |
| **הסמכות** | תפקידים מיוחדים שהחייל מוסמך אליהם. | `חובש`, `נהג, מפקד` |
| **שעות חסימה** | שעות שבהן החייל חסום מלעבוד (חופש, לו"ז אישי). | `8-12`, `22-6`, `14` |
        """)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 2. קובץ משימות")
        st.markdown("""
| עמודה | הסבר | דוגמה למילוי |
| :--- | :--- | :--- |
| **מס"ד / שם / סד"כ** | חובה: זיהוי המשימה וכמה חיילים נדרשים לה בכל רגע. | `101`, `שמירה`, `2` |
| **משך משמרת** | כמה שעות ברצף אורכת כל משמרת (משפיע על פיצולים). | `4`, `8` |
| **שעות מנוחה (מינימום)** | חובת מנוחה והפסקה *אחרי* שחייל מסיים את המשימה. | `8`, `0` |
| **אישור חפיפה** | האם ניתן לבצע במקביל למשימה אחרת? (למשל: כוננות). | `True`, `False` |
| **שעות פעילות** | טווח השעות שבהן המשימה קיימת לאורך היממה. | `all`, `10-18` |
| **הסמכה נדרשת** | תפקיד חובה הנדרש כדי לאייש את המשימה. | `נהג`, `חובש` |
| **עצימות (1-3)** | דירוג רמת השחיקה במשימה, לצורך חלוקת נטל הוגנת. | `1`, `2`, `3` |
        """)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 🧠 3. יכולות חכמות וכלים למפקד")
        st.markdown("""
        <div class="directive-box">
        <b>👨‍✈️ הנחיות מפקד מיוחדות (NLP):</b><br>
        הזנת "פקודות ברזל" לאלגוריתם לפני הריצה, ללא צורך לשנות את ה-Excel.<br>
        ניתן לכתוב טקסט חופשי, למשל: <i>"אריאל חייב לשמור בין 14 ל-18"</i> או <i>"תומר אסור לו לעבוד במטבח"</i>.<br>
        <span style="font-size: 0.9em; color: #5a2a82;">* במידה ותוזן הנחיה שסותרת אילוצים אחרים (למשל לשבץ חייל בשעות שהוא חסום), המערכת תתריע על כך.</span>
        </div>
        
        <div class="adhoc-box">
        <b>🔵 משימות אד-הוק (Ad-Hoc):</b><br>
        קפצה משימה פתאומית? אין צורך לערוך את הקבצים. ניתן להוסיף אותה זמנית תחת "משימות אד-הוק" במסך הראשי, והאלגוריתם יתחשב בה בשיבוץ.
        </div>
        
        <div class="rec-box">
        <b>💡 פענוח תוצאות והמלצות לחוסרים:</b><br>
        אם יוצגו "חורים" בלוח והמערכת תתקשה לאייש עמדה (סימון ⚠️), היא תנתח את הסיבה המתמטית (למשל: 'כל הנהגים בשעות מנוחה') ותמליץ לכם קונקרטית כיצד לפתור את הפלונטר.
        </div>
        """, unsafe_allow_html=True)

    # ── ביצוע שיבוץ ─────────────────────────────────────────────────
    with tab_run:
        col_u1, col_u2 = st.columns(2)
        with col_u1: sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
        with col_u2: tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

        # ── הנחיות מפקד ──────────────────────────────────────────────
        st.markdown("---")
        with st.expander("👨‍✈️ הנחיות מפקד מיוחדות — כללים קשיחים לפני הרצה", expanded=False):
            st.markdown("""
            <div class="directive-box">
            <b>שיטה 1:</b> בנאי מובנה — בחר חייל, סוג הנחיה, משימה ושעות.<br>
            <b>שיטה 2:</b> טקסט חופשי — "ארי חייב לשמור בין 14 ל-18" או "תומר אסור לו לעבוד במטבח"
            </div>
            """, unsafe_allow_html=True)

            st.markdown("**➕ הוסף הנחיה מובנית:**")
            with st.form("directive_structured", clear_on_submit=True):
                dc1, dc2, dc3, dc4, dc5 = st.columns([2,1.5,2,1,1])
                with dc1: dir_soldier = st.text_input("שם חייל", placeholder="ארי כהן")
                with dc2: dir_type = st.selectbox("סוג", ["חייב ✅", "אסור ❌"])
                with dc3: dir_task = st.text_input("שם משימה", placeholder="שמירה")
                with dc4: dir_sh = st.number_input("משעה", min_value=0, max_value=23, value=0)
                with dc5: dir_eh = st.number_input("עד שעה", min_value=0, max_value=23, value=23)
                if st.form_submit_button("➕ הוסף") and dir_soldier.strip() and dir_task.strip():
                    dtype = 'must' if 'חייב' in dir_type else 'must_not'
                    hours = list(range(dir_sh, dir_eh+1)) if dir_sh<=dir_eh else (list(range(dir_sh,24)) + list(range(0, dir_eh+1)))
                    hours = [] if (dir_sh == 0 and dir_eh == 23) else hours
                    st.session_state.directives.append(Directive(dir_soldier.strip(), dtype, dir_task.strip(), hours))
                    st.success(f"✅ נוספה הנחיה")

            st.markdown("**✍️ הוסף הנחיה בטקסט חופשי:**")
            with st.form("directive_freetext", clear_on_submit=True):
                free_text = st.text_input("כתוב הנחיה", placeholder='לדוגמה: "ארי חייב לשמור בין 14 ל-18"')
                if st.form_submit_button("🔍 פרסר והוסף") and free_text.strip():
                    st.session_state.directives.append(Directive("__RAW__", "__RAW__", "__RAW__", [free_text.strip()]))
                    st.info(f"📝 נשמר לפרסור בזמן הרצה")

            if st.session_state.directives:
                st.markdown("**📋 הנחיות פעילות:**")
                for i, d in enumerate(st.session_state.directives):
                    lbl = f"📝 טקסט חופשי: «{d.hours[0] if d.hours else '?'}»" if d.soldier_name == "__RAW__" else f"{'✅' if d.directive_type == 'must' else '❌'} {d.soldier_name} ← {d.task_name}"
                    col_lbl, col_del = st.columns([5, 1])
                    with col_lbl: st.markdown(f"  {i+1}. {lbl}")
                    with col_del:
                        if st.button("🗑️", key=f"del_dir_{i}"):
                            st.session_state.directives.pop(i)
                            st.rerun()

        # ── הגדרות מתקדמות ──────────────────────────────────────────
        with st.expander("⚙️ הגדרות מתקדמות"):
            use_cpsat  = st.toggle("שיפור CP-SAT", value=True)
            cpsat_time = st.slider("זמן CP-SAT (שניות)", 10, 180, 60, 5)

        # ── הרצה ─────────────────────────────────────────────────────
        if sf and tf:
            if st.button('⚙️ צור שבצ"ק חכם', use_container_width=True):
                try:
                    s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)

                    def find_col(df, kws):
                        for kw in kws:
                            m = [c for c in df.columns if kw in str(c)]
                            if m: return m[0]
                        return None

                    id_col, name_col = find_col(s_df, ['מספר','אישי']), find_col(s_df, ['שם'])
                    t_id_col, t_nm_col, t_req_col = find_col(t_df, ['מס"ד','מסד']), find_col(t_df, ['שם']), find_col(t_df, ['סד"כ','נדרש'])

                    if not all([id_col, name_col, t_id_col, t_nm_col, t_req_col]):
                        st.markdown('<div class="error-box">❌ לא נמצאו עמודות חובה באקסל.</div>', unsafe_allow_html=True)
                        st.stop()

                    s_df, t_df = s_df.dropna(subset=[id_col, name_col]), t_df.dropna(subset=[t_id_col, t_nm_col, t_req_col])
                    
                    shift_col, rest_col = find_col(t_df, ['משך']), find_col(t_df, ['מנוחה'])
                    overlap_col, hours_col = find_col(t_df, ['חפיפה']), find_col(t_df, ['פעילות','שעות'])
                    role_col, intense_col, block_col = find_col(t_df, ['הסמכה']), find_col(t_df, ['עצימות']), find_col(t_df, ['חסומים'])

                    soldiers = [Soldier(r[id_col], r[name_col], r.get('פטורים',''), r.get('הסמכות',''), r.get('שעות חסימה','')) for _, r in s_df.iterrows()]
                    soldiers.append(Soldier(DUMMY_SID, "⚠️ חוסר", is_dummy=True))

                    tasks = [Task(r[t_id_col], r[t_nm_col], r[t_req_col], r.get(shift_col,1), r.get(rest_col,0), r.get(overlap_col,False), r.get(hours_col,'all'), r.get(role_col,''), r.get(intense_col,1), r.get(block_col,'')) for _, r in t_df.iterrows()]

                    real_soldiers = [s for s in soldiers if not s.is_dummy]

                    # פרסר הנחיות מפקד
                    parsed_directives, raw_parse_errors = [], []
                    for d in st.session_state.directives:
                        if d.soldier_name == "__RAW__":
                            parsed, err = parse_free_text_directive(d.hours[0] if d.hours else "", real_soldiers, tasks)
                            if parsed: parsed_directives.append(parsed)
                            else: raw_parse_errors.append(err)
                        else: parsed_directives.append(d)

                    if raw_parse_errors:
                        st.markdown('<div class="warn-box">⚠️ ' + '<br>'.join(raw_parse_errors) + '</div>', unsafe_allow_html=True)

                    is_valid, validation_errors = validate_inputs_and_directives(soldiers, tasks, parsed_directives)
                    if not is_valid:
                        st.markdown('<div class="error-box">🚨 <b>שגיאות בהנחיות:</b><br>' + '<br>'.join(validation_errors) + '</div>', unsafe_allow_html=True)
                        st.stop()

                    with st.spinner("⚡ שלב 1: שיבוץ גריידי..."):
                        schedule, dummy_slots, work_h, int_load = greedy_schedule(soldiers, tasks)

                    was_infeasible = False
                    if use_cpsat:
                        with st.spinner(f"🔧 שלב 2: אופטימיזציה CP-SAT..."):
                            schedule, directive_errors, was_infeasible = improve_with_cpsat(soldiers, tasks, schedule, directives=parsed_directives, time_limit=cpsat_time)

                        if was_infeasible:
                            st.markdown('<div class="infeasible-box">🚨 <b>האופטימיזציה נכשלה</b> — הנחיות מפקד יצרו אילוץ בלתי אפשרי. הוחזר גיבוי.</div>', unsafe_allow_html=True)
                            insights = diagnose_infeasible_model(soldiers, tasks, parsed_directives)
                            for i in insights: st.markdown(f'<div class="warn-box">{i}</div>', unsafe_allow_html=True)

                    dummy_hours_count = sum(1 for t in tasks for si in range(len(t.slots)) for h in range(24) if schedule[t.tid][si][h] == DUMMY_SID)
                    
                    final_df = build_result_df(soldiers, tasks, schedule)
                    real_df = final_df[~final_df["שם"].str.startswith("⚠️")]
                    
                    st.markdown("---")
                    st.subheader("📅 לוח לפי חייל")
                    st.table(final_df)

                    st.markdown("---")
                    st.subheader("📋 לוח לפי משימה")
                    st.table(build_task_df(soldiers, tasks, schedule))
                    
                    if dummy_hours_count > 0:
                        st.markdown("---")
                        st.subheader("💡 המלצות המערכת לפתרון חוסרים")
                        for r in diagnose_dummy_slots(soldiers, tasks, schedule):
                            st.markdown(f"<div class='rec-box'>🎯 {r['משימה']} | 🕐 {r['שעות']}<br>📋 {r['סיבה']}<br>✅ {r['המלצה']}</div>", unsafe_allow_html=True)

                except Exception as e:
                    st.markdown('<div class="error-box">🚨 <b>שגיאה טכנית בזמן הרצה</b></div>', unsafe_allow_html=True)
                    st.code(traceback.format_exc())
        else:
            st.markdown('<div class="info-box">👆 העלו קבצי אקסל כדי להתחיל.</div>', unsafe_allow_html=True)

except Exception as e:
    st.error("🚨 שגיאה קריטית באפליקציה.")
    st.code(traceback.format_exc())
