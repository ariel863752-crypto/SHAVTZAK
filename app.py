import io
import re
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
import traceback
from collections import defaultdict

# ══════════════════════════════════════════════════════════════════
# 1. עיצוב
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title='שבצ"ק חכם', page_icon="🪖", layout="wide")
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
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val, is_task=True) -> list:
    if pd.isna(val) or str(val).strip().lower() in ('all', '', 'nan'):
        return list(range(24)) if is_task else []

    def to_hour(s: str) -> int:
        s = s.strip()
        if ':' in s:
            return int(s.split(':')[0])
        return int(float(s.replace('.0', '') or '0'))

    res = set()
    for part in str(val).split(','):
        part = part.strip()
        if not part or part.lower() == 'nan':
            continue
        if '-' in part:
            ps = part.split('-')
            start_s, end_s = ps[0].strip(), ps[-1].strip()
            if start_s and end_s:
                s = to_hour(start_s)
                e = to_hour(end_s)
                if s <= e:
                    for h in range(s, e + 1): res.add(h % 24)
                else:
                    for h in range(s, 24): res.add(h)
                    for h in range(0, e + 1): res.add(h)
        else:
            try: res.add(to_hour(part) % 24)
            except (ValueError, IndexError): pass
    return sorted(res)


DUMMY_SID       = "__DUMMY__"
_ADHOC_TID_START = 9000


class Soldier:
    def __init__(self, s_id, name, restr='', roles='', unavail='', is_dummy=False):
        self.sid      = str(s_id).replace('.0', '').strip()
        self.name     = str(name).strip()
        self.is_dummy = is_dummy
        self.exempt   = (
            [int(float(t)) for t in str(restr).split(',')
             if str(t).strip().replace('.0', '').isdigit()]
            if pd.notna(restr) and str(restr).strip() not in ('', 'nan') else [])
        self.roles = (
            [r.strip() for r in str(roles).split(',') if r.strip()]
            if pd.notna(roles) and str(roles).strip() not in ('', 'nan') else [])
        self.blocked_hours = [] if is_dummy else parse_time_ranges(unavail, is_task=False)

    def can_assign(self, task, shift_hours: list) -> bool:
        if self.is_dummy: return True
        if any(h in self.blocked_hours for h in shift_hours): return False
        if task.tid in self.exempt: return False
        if any(r in task.block_roles for r in self.roles): return False
        return True


class Task:
    def __init__(self, t_id, name, req, shift, rest, overlap, hours,
                 req_roles, intensity, block_roles='', is_adhoc=False):
        self.tid       = int(float(t_id))
        self.name      = str(name).strip()
        self.req       = int(req)
        self.shift     = int(shift)   if pd.notna(shift)   else 1
        self.rest      = int(rest)    if pd.notna(rest)    else 0
        self.overlap   = str(overlap).strip().lower() == 'true'
        self.hours     = parse_time_ranges(hours)
        self.intensity = int(intensity) if pd.notna(intensity) else 1
        self.is_adhoc  = is_adhoc
        self.block_roles = (
            [r.strip() for r in str(block_roles).split(',') if r.strip()]
            if pd.notna(block_roles) and str(block_roles).strip() not in ('', 'nan') else [])
        parsed = (
            [r.strip() for r in str(req_roles).split(',')]
            if pd.notna(req_roles) and str(req_roles).strip() not in ('', 'nan') else [])
        self.slots = parsed[:]
        while len(self.slots) < self.req:
            self.slots.append(None)

    def slot_ok(self, slot_idx: int, soldier) -> bool:
        if soldier.is_dummy: return True
        req_role = self.slots[slot_idx]
        if req_role is None: return True
        return req_role in soldier.roles

    def get_shift_starts(self) -> list:
        if not self.hours: return []
        hours_set    = set(self.hours)
        sorted_hours = sorted(self.hours)
        starts, i    = [], 0
        while i < len(sorted_hours):
            start = sorted_hours[i]
            if all((start + k) % 24 in hours_set for k in range(self.shift)):
                starts.append(start)
                i += self.shift
            else:
                i += 1
        return starts


# ══════════════════════════════════════════════════════════════════
# 3. הנחיות מפקד — מבנה נתונים ופרסור
# ══════════════════════════════════════════════════════════════════
class Directive:
    """
    הנחיית מפקד: חייל X חייב/אסור לבצע משימה Y בשעות Z.
    directive_type: 'must' | 'must_not'
    hours: רשימת שעות (ריק = כל שעות המשימה)
    """
    def __init__(self, soldier_name: str, directive_type: str,
                 task_name: str, hours: list):
        self.soldier_name   = soldier_name.strip()
        self.directive_type = directive_type   # 'must' / 'must_not'
        self.task_name      = task_name.strip()
        self.hours          = hours            # [] = כל שעות המשימה

    def __repr__(self):
        h_str = f" בשעות {self.hours}" if self.hours else " (כל היום)"
        arrow = "✅ חייב" if self.directive_type == 'must' else "❌ אסור"
        return f"{arrow}: {self.soldier_name} ← {self.task_name}{h_str}"


def parse_free_text_directive(text: str,
                               soldiers: list,
                               tasks: list) -> tuple[Directive | None, str]:
    """
    מנסה לפרש משפט חופשי לאובייקט Directive.
    מחזיר (Directive, "") אם הצליח, (None, error_msg) אם נכשל.

    תבניות נתמכות:
      "ארי חייב לשמור בין 14:00 ל-18:00"
      "תומר אסור לו לעבוד במטבח"
      "רון חייב לבצע סיור רכוב בשעה 08:00"
      "ניצן לא יכול לשרת בכוננות"
    """
    text = text.strip()

    # ── זיהוי חייל ──────────────────────────────────────────────
    soldier_obj = None
    for s in soldiers:
        if s.is_dummy: continue
        # חיפוש לפי שם ראשון או שם מלא (case-insensitive)
        first_name = s.name.split()[0] if s.name.split() else s.name
        if (first_name in text or s.name in text):
            soldier_obj = s
            break
    if soldier_obj is None:
        names = [s.name for s in soldiers if not s.is_dummy]
        return None, f"לא זוהה שם חייל בטקסט. חיילים קיימים: {', '.join(names[:8])}"

    # ── זיהוי סוג הנחיה ─────────────────────────────────────────
    must_keywords     = ['חייב', 'חייבת', 'יבצע', 'תבצע', 'ישרת', 'תשרת',
                          'ישמור', 'תשמור', 'must', 'חייב לשמור', 'חייב לבצע']
    must_not_keywords = ['אסור', 'לא יכול', 'לא יעבוד', 'לא תעבוד', 'לא ישרת',
                          'לא ישמור', 'must not', 'cannot', 'לא יבצע']
    directive_type = None
    for kw in must_not_keywords:
        if kw in text:
            directive_type = 'must_not'
            break
    if directive_type is None:
        for kw in must_keywords:
            if kw in text:
                directive_type = 'must'
                break
    if directive_type is None:
        return None, "לא זוהתה הנחיה. השתמש במילים: חייב / אסור / לא יכול"

    # ── זיהוי משימה ─────────────────────────────────────────────
    task_obj = None
    for t in tasks:
        if t.name in text:
            task_obj = t
            break
    if task_obj is None:
        task_names = [t.name for t in tasks]
        return None, f"לא זוהתה משימה. משימות קיימות: {', '.join(task_names)}"

    # ── זיהוי שעות (אופציונלי) ──────────────────────────────────
    hours = []
    # תבנית: "בין 14 ל-18" / "בין 14:00 ל-18:00"
    m = re.search(r'בין\s+(\d{1,2})(?::\d{2})?\s+ל[-–]\s*(\d{1,2})(?::\d{2})?', text)
    if m:
        sh, eh = int(m.group(1)), int(m.group(2))
        hours  = list(range(sh, eh + 1)) if sh <= eh else (
            list(range(sh, 24)) + list(range(0, eh + 1)))
    else:
        # תבנית: "בשעה 14" / "בשעה 14:00"
        m2 = re.search(r'בשעה\s+(\d{1,2})(?::\d{2})?', text)
        if m2:
            hours = [int(m2.group(1)) % 24]
        else:
            # תבנית: "14:00-18:00" ישיר
            m3 = re.search(r'(\d{1,2})(?::\d{2})?\s*[-–]\s*(\d{1,2})(?::\d{2})?', text)
            if m3:
                sh, eh = int(m3.group(1)), int(m3.group(2))
                hours  = list(range(sh, eh + 1)) if sh <= eh else (
                    list(range(sh, 24)) + list(range(0, eh + 1)))

    return Directive(soldier_obj.name, directive_type, task_obj.name, hours), ""


# ══════════════════════════════════════════════════════════════════
# 4. Excel מעוצב
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


def to_excel_task_view(soldiers, tasks, schedule) -> bytes:
    df = build_task_df(soldiers, tasks, schedule)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        df.to_excel(w, index=True, sheet_name='לפי משימות')
        wb, ws = w.book, w.sheets['לפי משימות']
        hf   = wb.add_format({'bold':True,'fg_color':'#1a3d17','font_color':'white',
                               'border':1,'align':'center','valign':'vcenter'})
        idx_f= wb.add_format({'bold':True,'fg_color':'#2d5a27','font_color':'white',
                               'align':'center','valign':'vcenter'})
        ok_f = wb.add_format({'fg_color':'#edf5ec','align':'right','text_wrap':True})
        wf   = wb.add_format({'fg_color':'#fdecea','font_color':'#c0392b','bold':True,
                               'align':'right','text_wrap':True})
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
# 5. גריידי v20
# ══════════════════════════════════════════════════════════════════
def greedy_schedule(soldiers: list, tasks: list, num_hours: int = 24):
    schedule = {
        t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))}
        for t in tasks
    }
    assignments: dict[str, dict[int, list]] = {
        s.sid: {h: [] for h in range(num_hours)} for s in soldiers
    }
    assignments[DUMMY_SID] = {h: [] for h in range(num_hours)}
    work_hours:     dict[str, int] = {s.sid: 0 for s in soldiers}
    work_hours[DUMMY_SID]          = 0
    intensity_load: dict[str, int] = {s.sid: 0 for s in soldiers}
    intensity_load[DUMMY_SID]      = 0
    dummy_slots = []

    def is_free(sid: str, shift_hours: list, new_task: Task) -> bool:
        for h in shift_hours:
            for ex in assignments[sid][h]:
                if not new_task.overlap: return False
                if not ex.overlap:       return False
        return True

    def mark_busy(sid: str, shift_hours: list, task: Task):
        for h in shift_hours:
            assignments[sid][h].append(task)

    sorted_tasks = sorted(
        tasks,
        key=lambda t: (1 if t.overlap else 0, -t.req, -t.intensity, -t.shift)
    )

    for t in sorted_tasks:
        for slot_idx in range(len(t.slots)):
            for start_h in t.get_shift_starts():
                shift_hours = [(start_h + i) % num_hours for i in range(t.shift)]
                end_h       = shift_hours[-1]
                candidates  = []
                for s in soldiers:
                    if s.is_dummy: continue
                    if not s.can_assign(t, shift_hours): continue
                    if not t.slot_ok(slot_idx, s): continue
                    if not is_free(s.sid, shift_hours, t): continue
                    candidates.append((work_hours[s.sid]*10 + intensity_load[s.sid], s.sid, s))
                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    _, chosen_sid, _ = candidates[0]
                    chosen_dummy = False
                else:
                    chosen_sid   = DUMMY_SID
                    chosen_dummy = True
                for h in shift_hours:
                    schedule[t.tid][slot_idx][h] = chosen_sid
                if chosen_dummy:
                    dummy_slots.append((t.name, slot_idx+1, start_h, end_h))
                else:
                    work_hours[chosen_sid]     += t.shift
                    intensity_load[chosen_sid] += t.shift * t.intensity
                    mark_busy(chosen_sid, shift_hours, t)

    return schedule, dummy_slots, work_hours, intensity_load


# ══════════════════════════════════════════════════════════════════
# 6. CP-SAT v20 — עם הנחיות מפקד
# ══════════════════════════════════════════════════════════════════
def improve_with_cpsat(soldiers: list, tasks: list, schedule: dict,
                       directives: list,          # ← חדש: רשימת Directive
                       num_hours: int = 24,
                       time_limit: float = 60.0):
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    model         = cp_model.CpModel()

    # ── build lookup maps ────────────────────────────────────────
    name_to_soldier = {s.name: s for s in real_soldiers}
    name_to_task    = {t.name: t for t in tasks}

    # ── משתני החלטה ──────────────────────────────────────────────
    x = {}
    for s in real_soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.sid, t.tid, si, h] = model.NewBoolVar(f"x_{s.sid}_{t.tid}_{si}_{h}")

    for s in real_soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    model.AddHint(x[s.sid, t.tid, si, h],
                                  1 if schedule[t.tid][si][h] == s.sid else 0)

    dummy_vars = {}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in t.hours:
                dv = model.NewBoolVar(f"d_{t.tid}_{si}_{h}")
                dummy_vars[t.tid, si, h] = dv
                model.AddHint(dv, 1 if schedule[t.tid][si][h] == DUMMY_SID else 0)

    # כיסוי
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h in t.hours:
                    model.Add(sum(x[s.sid, t.tid, si, h] for s in real_soldiers)
                              + dummy_vars[t.tid, si, h] == 1)
                else:
                    for s in real_soldiers:
                        model.Add(x[s.sid, t.tid, si, h] == 0)

    # כשירות
    for s in real_soldiers:
        for t in tasks:
            task_blocked = t.tid in s.exempt or any(r in t.block_roles for r in s.roles)
            for si in range(len(t.slots)):
                role_blocked = (t.slots[si] is not None and t.slots[si] not in s.roles)
                for h in range(num_hours):
                    if task_blocked or role_blocked or h in s.blocked_hours:
                        model.Add(x[s.sid, t.tid, si, h] == 0)

    # חפיפה סימטרית
    for s in real_soldiers:
        for h in range(num_hours):
            for t1 in tasks:
                if t1.overlap: continue
                for si1 in range(len(t1.slots)):
                    for t2 in tasks:
                        if t2.tid == t1.tid: continue
                        for si2 in range(len(t2.slots)):
                            model.Add(x[s.sid,t1.tid,si1,h] + x[s.sid,t2.tid,si2,h] <= 1)

    # no-self-duplication
    for s in real_soldiers:
        for t in tasks:
            for h in t.hours:
                if len(t.slots) > 1:
                    model.Add(sum(x[s.sid,t.tid,si,h] for si in range(len(t.slots))) <= 1)

    # ── הנחיות מפקד — HARD constraints ──────────────────────────
    directive_errors = []
    for d in directives:
        s_obj = name_to_soldier.get(d.soldier_name)
        t_obj = name_to_task.get(d.task_name)
        if s_obj is None or t_obj is None:
            directive_errors.append(f"הנחיה לא הוחלה: {d} (חייל/משימה לא נמצאו)")
            continue
        hours_to_apply = d.hours if d.hours else t_obj.hours
        for h in hours_to_apply:
            if h not in t_obj.hours:
                continue  # שעה לא פעילה — דלג
            for si in range(len(t_obj.slots)):
                if d.directive_type == 'must':
                    # חייב — לפחות סלוט אחד פעיל עבורו בשעה זו
                    # (אם יש כמה slots, מספיק שאחד יהיה שלו)
                    if si == 0:   # נאכוף רק על סלוט 0 למניעת conflic
                        model.Add(x[s_obj.sid, t_obj.tid, si, h] == 1)
                elif d.directive_type == 'must_not':
                    # אסור בכל הסלוטים
                    model.Add(x[s_obj.sid, t_obj.tid, si, h] == 0)

    # ── פונקציית מטרה ────────────────────────────────────────────
    PENALTY_DUMMY    = 100_000
    PENALTY_NIGHT    =     800
    PENALTY_SPLIT    =     400
    PENALTY_FAIRNESS =     100
    PENALTY_REST     =      50
    penalties = []

    for t in tasks:
        for si in range(len(t.slots)):
            for h in t.hours:
                penalties.append(dummy_vars[t.tid, si, h] * PENALTY_DUMMY)

    NIGHT_HOURS = sorted(set(range(22, 24)) | set(range(0, 9)))
    for s in real_soldiers:
        nv = [x[s.sid,t.tid,si,h] for t in tasks if not t.overlap
              for si in range(len(t.slots)) for h in NIGHT_HOURS]
        if nv:
            nt = model.NewIntVar(0, len(NIGHT_HOURS)*len(tasks), f'nt_{s.sid}')
            model.Add(nt == sum(nv))
            ne = model.NewIntVar(0, 200, f'ne_{s.sid}')
            model.Add(ne >= nt - 4)
            model.Add(ne >= 0)
            penalties.append(ne * PENALTY_NIGHT)

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
                        penalties.append(sv * PENALTY_SPLIT)

    th_vars = []
    for s in real_soldiers:
        th = model.NewIntVar(0, num_hours*len(tasks), f'th_{s.sid}')
        model.Add(th == sum(x[s.sid,t.tid,si,h]
                            for t in tasks for si in range(len(t.slots))
                            for h in range(num_hours)))
        th_vars.append(th)
    max_h = model.NewIntVar(0, 1000, 'max_h')
    min_h = model.NewIntVar(0, 1000, 'min_h')
    model.AddMaxEquality(max_h, th_vars)
    model.AddMinEquality(min_h, th_vars)
    diff = model.NewIntVar(0, 1000, 'diff')
    model.Add(diff == max_h - min_h)
    penalties.append(diff * PENALTY_FAIRNESS)

    for s in real_soldiers:
        for t in tasks:
            if t.overlap or t.rest == 0: continue
            for start_h in t.get_shift_starts():
                end_h   = (start_h + t.shift - 1) % num_hours
                in_s    = model.NewBoolVar(f'ins_{s.sid}_{t.tid}_{start_h}')
                model.Add(sum(x[s.sid,t.tid,si,start_h] for si in range(len(t.slots)))>=1).OnlyEnforceIf(in_s)
                model.Add(sum(x[s.sid,t.tid,si,start_h] for si in range(len(t.slots)))==0).OnlyEnforceIf(in_s.Not())
                for r in range(1, t.rest+1):
                    rest_h = (end_h+r) % num_hours
                    rw     = sum(x[s.sid,t2.tid,si2,rest_h]
                                 for t2 in tasks if not t2.overlap
                                 for si2 in range(len(t2.slots)))
                    ir = model.NewBoolVar(f'ir_{s.sid}_{t.tid}_{start_h}_{r}')
                    model.Add(rw>=1).OnlyEnforceIf(ir)
                    model.Add(rw==0).OnlyEnforceIf(ir.Not())
                    viol = model.NewBoolVar(f'rv_{s.sid}_{t.tid}_{start_h}_{r}')
                    model.AddBoolAnd([in_s, ir]).OnlyEnforceIf(viol)
                    model.AddBoolOr([in_s.Not(), ir.Not()]).OnlyEnforceIf(viol.Not())
                    penalties.append(viol * PENALTY_REST)

    max_p = (num_hours * sum(len(t.slots) for t in tasks) * PENALTY_DUMMY
             + len(real_soldiers) * 200 * PENALTY_NIGHT + 1_000_000)
    tp = model.NewIntVar(0, max_p, 'tp')
    model.Add(tp == sum(penalties))
    model.Minimize(tp)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers  = 8
    solver.parameters.log_search_progress = False
    solver.parameters.relative_gap_limit  = 0.03
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return schedule, directive_errors

    new_sch = {t.tid: {si: {h: None for h in range(num_hours)}
                       for si in range(len(t.slots))} for t in tasks}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h not in t.hours: continue
                if solver.Value(dummy_vars[t.tid, si, h]) == 1:
                    new_sch[t.tid][si][h] = DUMMY_SID
                else:
                    for s in real_soldiers:
                        if solver.Value(x[s.sid, t.tid, si, h]) == 1:
                            new_sch[t.tid][si][h] = s.sid
                            break
    return new_sch, directive_errors


# ══════════════════════════════════════════════════════════════════
# 7. אבחון חוסרים — המלצות מפקד
# ══════════════════════════════════════════════════════════════════
def diagnose_dummy_slots(soldiers: list, tasks: list,
                         schedule: dict, num_hours: int = 24) -> list[dict]:
    """
    לכל שעת-סלוט עם DUMMY — מנתח מדוע לא נמצא חייל ומציע המלצה.
    מחזיר רשימת dict: {task, hour, reason, recommendation}
    """
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    recommendations = []
    seen = set()  # מניעת כפילויות

    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if schedule[t.tid][si][h] != DUMMY_SID:
                    continue
                key = (t.tid, si, h)
                if key in seen:
                    continue
                seen.add(key)

                # ── נתח מדוע אין חייל ──────────────────────────
                blocked_by_role     = []
                blocked_by_exemption= []
                blocked_by_hours    = []
                blocked_by_busy     = []
                wrong_role          = []
                potentially_free    = []

                for s in real_soldiers:
                    # פטור
                    if t.tid in s.exempt:
                        blocked_by_exemption.append(s.name)
                        continue
                    # block_roles
                    if any(r in t.block_roles for r in s.roles):
                        blocked_by_role.append(s.name)
                        continue
                    # שעות חסימה
                    if h in s.blocked_hours:
                        blocked_by_hours.append(s.name)
                        continue
                    # הסמכה נדרשת
                    req_role = t.slots[si]
                    if req_role and req_role not in s.roles:
                        wrong_role.append(s.name)
                        continue
                    # אחרת — פנוי תיאורטית (אבל עסוק ע"י גריידי)
                    potentially_free.append(s.name)

                # ── קבע סיבה עיקרית ────────────────────────────
                req_role_name = t.slots[si]
                if len(real_soldiers) == 0:
                    reason = "אין חיילים כלל במערכת"
                    rec    = "הוסף חיילים לקובץ החיילים"
                elif req_role_name and len([s for s in real_soldiers
                                            if req_role_name in s.roles]) == 0:
                    reason = f"אין חייל עם הסמכה '{req_role_name}' במערכת"
                    rec    = f"הוסף חייל עם הסמכה: {req_role_name}"
                elif req_role_name and len(wrong_role) == len(real_soldiers) - len(blocked_by_exemption) - len(blocked_by_role):
                    needed = len([s for s in real_soldiers if req_role_name in s.roles])
                    reason = f"כל החיילים המוסמכים ל'{req_role_name}' ({needed} חיילים) כבר עסוקים"
                    rec    = f"הוסף עוד חייל עם הסמכה '{req_role_name}', או הזז את '{t.name}' לשעה אחרת"
                elif len(blocked_by_hours) == len(real_soldiers):
                    reason = f"כל החיילים חסומים בשעה {h:02d}:00"
                    rec    = f"בדוק שעות חסימה אישיות — ייתכן שהגדרת חסימה כוללנית מדי"
                elif len(blocked_by_exemption) + len(blocked_by_role) == len(real_soldiers):
                    reason = f"כל החיילים פטורים ממשימה זו או חסומי תפקיד"
                    rec    = f"הסר פטורים/חסימות, או הוסף חייל שאינו פטור מ'{t.name}'"
                elif len(potentially_free) > 0:
                    # יש חיילים תיאורטית זמינים — אבל עסוקים מחסימת מנוחה/עומס
                    reason = (f"{len(potentially_free)} חיילים זמינים עקרונית "
                              f"({', '.join(potentially_free[:3])}{'...' if len(potentially_free)>3 else ''}) "
                              f"אך כולם עסוקים בשעה {h:02d}:00 עקב משמרת קיימת/מנוחה")
                    rec    = (f"קצר את 'שעות מנוחה' של משימות קודמות, "
                              f"או הזז את '{t.name}' לשעה אחרת (למשל {(h+2)%24:02d}:00)")
                else:
                    total_avail = len(real_soldiers) - len(blocked_by_exemption) - len(blocked_by_role)
                    reason = f"רק {total_avail} חיילים זמינים, כולם עסוקים"
                    rec    = f"הוסף חייל נוסף, או הפחת את 'סד\"כ נדרש' של '{t.name}'"

                recommendations.append({
                    "משימה":   f"{'🔵 ' if t.is_adhoc else ''}{t.name}",
                    "שעה":     f"{h:02d}:00",
                    "סלוט":    si + 1,
                    "סיבה":    reason,
                    "המלצה":   rec,
                })

    # מיזוג המלצות זהות — מקבץ לפי (task, reason)
    merged: dict[tuple, dict] = {}
    for r in recommendations:
        key = (r["משימה"], r["סיבה"])
        if key not in merged:
            merged[key] = {**r, "_hours": [r["שעה"]]}
        else:
            merged[key]["_hours"].append(r["שעה"])

    final = []
    for key, rec in merged.items():
        hours_list = rec.pop("_hours")
        if len(hours_list) <= 3:
            rec["שעות"] = ", ".join(hours_list)
        else:
            rec["שעות"] = f"{hours_list[0]}–{hours_list[-1]} ({len(hours_list)} שעות)"
        del rec["שעה"]
        final.append(rec)

    return final


# ══════════════════════════════════════════════════════════════════
# 8. בניית DataFrames
# ══════════════════════════════════════════════════════════════════
def build_result_df(soldiers, tasks, schedule, num_hours=24):
    SLEEP = set(range(22,24)) | set(range(0,9))
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
    for s in soldiers:
        if s.is_dummy: continue
        row = {"שם": s.name}
        total = night = intensity = 0
        for h in range(num_hours):
            active = []
            for t in tasks:
                for si in range(len(t.slots)):
                    if schedule[t.tid][si][h] == s.sid:
                        active.append(("🔵 " if t.is_adhoc else "") + t.name)
                        if not t.overlap:
                            if h in SLEEP: night += 1
                            intensity += t.intensity
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if active: total += 1
        row["סך שעות"]          = total
        row["מדד עצימות"]       = intensity
        row["שעות שינה (22-08)"] = len(SLEEP) - night
        rows.append(row)

    dummy_row = {"שם": "⚠️ חוסר כוח אדם"}
    for h in range(num_hours):
        d = [t.name for t in tasks for si in range(len(t.slots))
             if schedule[t.tid][si][h] == DUMMY_SID]
        dummy_row[hour_labels[h]] = "❌ " + " + ".join(d) if d else "—"
    dummy_row["סך שעות"]          = sum(1 for t in tasks for si in range(len(t.slots))
                                        for h in range(num_hours)
                                        if schedule[t.tid][si][h] == DUMMY_SID)
    dummy_row["מדד עצימות"]       = "—"
    dummy_row["שעות שינה (22-08)"] = "—"

    has_dummy = any(schedule[t.tid][si][h] == DUMMY_SID
                    for t in tasks for si in range(len(t.slots)) for h in range(num_hours))
    df = pd.DataFrame(rows)
    if has_dummy:
        df = pd.concat([df, pd.DataFrame([dummy_row])], ignore_index=True)
    df.index = range(1, len(df)+1)
    return df


def build_task_df(soldiers, tasks, schedule, num_hours=24):
    sid_to_name = {s.sid: s.name for s in soldiers if not s.is_dummy}
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
    for h in range(num_hours):
        row = {"שעה": hour_labels[h]}
        for t in tasks:
            if h not in t.hours:
                row[t.name] = "—"
                continue
            names, has_dummy = [], False
            for si in range(len(t.slots)):
                val = schedule[t.tid][si][h]
                if val == DUMMY_SID: has_dummy = True
                elif val is not None: names.append(sid_to_name.get(val, val))
            if has_dummy:
                row[t.name] = "⚠️ חוסר" + (" + " + ", ".join(names) if names else "")
            elif names:
                row[t.name] = ", ".join(names)
            else:
                row[t.name] = "—"
        rows.append(row)
    df = pd.DataFrame(rows)
    df.set_index("שעה", inplace=True)
    return df


# ══════════════════════════════════════════════════════════════════
# 9. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v20)</h1>
  <p>הנחיות מפקד · אבחון חוסרים חכם · CP-SAT · split-shifts · תצוגת משימות</p>
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
                           file_name="Soldiers_v20.xlsx", use_container_width=True)
    with c2:
        st.markdown("**📋 תבנית משימות**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button("⬇️ הורד", data=to_excel_styled(t_ex,"Tasks",False),
                           file_name="Tasks_v20.xlsx", use_container_width=True)

# ── מדריך ───────────────────────────────────────────────────────
with tab_guide:
    st.markdown("### 📖 מדריך v20")
    st.markdown("""
<div class="info-box">
<b>✅ חדש: הנחיות מפקד:</b> הכנס משפטים חופשיים כמו:<br>
• "ארי חייב לשמור בין 14:00 ל-18:00"<br>
• "תומר אסור לו לעבוד במטבח"<br>
• "רון חייב לבצע סיור רכוב בשעה 08:00"<br>
<b>✅ חדש: אבחון חוסרים:</b> אם נוצרו רפאים, המערכת מנתחת את הסיבה ומציעה המלצה פעולה.<br>
<b>✅ שמות עמודות:</b> "שעות מנוחה מינימליות בין משימות" — גם השם הישן עובד.
</div>
""", unsafe_allow_html=True)

# ── ביצוע שיבוץ ─────────────────────────────────────────────────
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    # ── הנחיות מפקד ──────────────────────────────────────────────
    st.markdown("---")
    with st.expander("👨‍✈️ הנחיות מפקד מיוחדות — כללים קשיחים לפני הרצה", expanded=False):
        st.markdown("""
        <div class="directive-box">
        <b>שיטה 1 — בנאי מובנה:</b> בחר חייל, סוג הנחיה, משימה ושעות.<br>
        <b>שיטה 2 — טקסט חופשי:</b> כתוב משפט בעברית, המערכת תפרש אותו.<br>
        <b>דוגמאות:</b> "ארי חייב לשמור בין 14 ל-18" · "תומר אסור לו לעבוד במטבח"
        </div>
        """, unsafe_allow_html=True)

        if "directives" not in st.session_state:
            st.session_state.directives = []

        # ── בנאי מובנה ──────────────────────────────────────────
        st.markdown("**➕ הוסף הנחיה מובנית:**")
        with st.form("directive_structured", clear_on_submit=True):
            dc1, dc2, dc3, dc4, dc5 = st.columns([2,1.5,2,1,1])
            with dc1:
                dir_soldier = st.text_input("שם חייל", placeholder="ארי כהן")
            with dc2:
                dir_type = st.selectbox("סוג", ["חייב ✅", "אסור ❌"])
            with dc3:
                dir_task = st.text_input("שם משימה", placeholder="שמירה")
            with dc4:
                dir_sh = st.number_input("משעה", min_value=0, max_value=23, value=0)
            with dc5:
                dir_eh = st.number_input("עד שעה", min_value=0, max_value=23, value=23)
            sub_structured = st.form_submit_button("➕ הוסף")
            if sub_structured and dir_soldier.strip() and dir_task.strip():
                dtype = 'must' if 'חייב' in dir_type else 'must_not'
                hours = list(range(dir_sh, dir_eh+1)) if dir_sh<=dir_eh else (
                    list(range(dir_sh,24)) + list(range(0, dir_eh+1)))
                hours = [] if (dir_sh == 0 and dir_eh == 23) else hours
                st.session_state.directives.append(
                    Directive(dir_soldier.strip(), dtype, dir_task.strip(), hours))
                st.success(f"✅ נוספה הנחיה: {dir_soldier} {'חייב' if dtype=='must' else 'אסור'} ← {dir_task}")

        # ── טקסט חופשי ──────────────────────────────────────────
        st.markdown("**✍️ הוסף הנחיה בטקסט חופשי:**")
        with st.form("directive_freetext", clear_on_submit=True):
            free_text = st.text_input("כתוב הנחיה",
                placeholder='לדוגמה: "ארי חייב לשמור בין 14 ל-18" או "תומר אסור לו לעבוד במטבח"')
            sub_free = st.form_submit_button("🔍 פרסר והוסף")
            if sub_free and free_text.strip():
                # נוכל לפרסר רק אם נטענו חיילים ומשימות
                # אם הקבצים לא נטענו — נשמור כטקסט גולמי ונפרסר בזמן הרצה
                st.session_state.directives.append(
                    Directive("__RAW__", "__RAW__", "__RAW__", [free_text.strip()]))
                st.info(f"📝 נשמר לפרסור בזמן הרצה: «{free_text.strip()}»")

        # ── הצגת הנחיות קיימות ──────────────────────────────────
        if st.session_state.directives:
            st.markdown("**📋 הנחיות פעילות:**")
            for i, d in enumerate(st.session_state.directives):
                if d.soldier_name == "__RAW__":
                    lbl = f"📝 טקסט חופשי: «{d.hours[0] if d.hours else '?'}»"
                else:
                    icon = "✅" if d.directive_type == 'must' else "❌"
                    h_str = f" ({d.hours[0]}:00–{d.hours[-1]}:00)" if d.hours else " (כל היום)"
                    lbl = f"{icon} {d.soldier_name} ← {d.task_name}{h_str}"
                col_lbl, col_del = st.columns([5, 1])
                with col_lbl: st.markdown(f"  {i+1}. {lbl}")
                with col_del:
                    if st.button("🗑️", key=f"del_dir_{i}"):
                        st.session_state.directives.pop(i)
                        st.rerun()

    # ── אד-הוק ──────────────────────────────────────────────────
    with st.expander("🔵 משימות אד-הוק", expanded=False):
        if "adhoc_tasks" not in st.session_state:
            st.session_state.adhoc_tasks = []
        with st.form("adhoc_form", clear_on_submit=True):
            ac1, ac2, ac3 = st.columns(3)
            with ac1:
                ah_name   = st.text_input("שם משימה")
                ah_req    = st.number_input("כמות חיילים", min_value=1, max_value=20, value=2)
            with ac2:
                ah_start  = st.number_input("שעת התחלה", min_value=0, max_value=23, value=8)
                ah_end    = st.number_input("שעת סיום",   min_value=0, max_value=23, value=16)
            with ac3:
                ah_over   = st.selectbox("חפיפה", ["False (בלעדי)", "True (שיתופי)"])
                ah_intens = st.selectbox("עצימות", [1, 2, 3])
            if st.form_submit_button("➕ הוסף") and ah_name.strip():
                st.session_state.adhoc_tasks.append({
                    "name": ah_name.strip(), "req": ah_req,
                    "start": ah_start, "end": ah_end,
                    "overlap": "True" in ah_over, "intensity": ah_intens,
                })
        if st.session_state.adhoc_tasks:
            st.dataframe(pd.DataFrame(st.session_state.adhoc_tasks), use_container_width=True, hide_index=True)
            if st.button("🗑️ נקה אד-הוק"):
                st.session_state.adhoc_tasks = []
                st.rerun()

    # ── הגדרות מתקדמות ──────────────────────────────────────────
    with st.expander("⚙️ הגדרות מתקדמות"):
        use_cpsat  = st.toggle("שיפור CP-SAT", value=True)
        cpsat_time = st.slider("זמן CP-SAT (שניות)", 10, 180, 60, 5)

    # ── הרצה ─────────────────────────────────────────────────────
    if sf and tf:
        if st.button('⚙️ צור שבצ"ק חכם (v20)', use_container_width=True, key="run_btn"):
            try:
                s_df = pd.read_excel(sf)
                t_df = pd.read_excel(tf)

                def find_col(df, kws):
                    for kw in kws:
                        m = [c for c in df.columns if kw in str(c)]
                        if m: return m[0]
                    return None

                id_col    = find_col(s_df, ['מספר','אישי'])
                name_col  = find_col(s_df, ['שם'])
                t_id_col  = find_col(t_df, ['מס"ד','מסד'])
                t_nm_col  = find_col(t_df, ['שם'])
                t_req_col = find_col(t_df, ['סד"כ','נדרש'])

                if not all([id_col, name_col, t_id_col, t_nm_col, t_req_col]):
                    st.markdown('<div class="error-box">❌ לא נמצאו עמודות חובה.</div>',
                                unsafe_allow_html=True)
                    st.stop()

                s_df = s_df.dropna(subset=[id_col, name_col])
                t_df = t_df.dropna(subset=[t_id_col, t_nm_col, t_req_col])

                def gcol(df, kws): return find_col(df, kws)
                def gval(r, col, default=''):
                    return r.get(col, default) if col else default

                shift_col   = gcol(t_df, ['משך'])
                rest_col    = gcol(t_df, ['מנוחה מינימליות','מנוחה'])
                overlap_col = gcol(t_df, ['חפיפה'])
                hours_col   = gcol(t_df, ['פעילות','שעות'])
                role_col    = gcol(t_df, ['הסמכה'])
                intense_col = gcol(t_df, ['עצימות','דירוג'])
                block_col   = gcol(t_df, ['חסומים'])

                soldiers = [
                    Soldier(r[id_col], r[name_col],
                            r.get('פטורים',''), r.get('הסמכות',''),
                            r.get('שעות חסימה',''))
                    for _, r in s_df.iterrows()
                ]
                soldiers.append(Soldier(DUMMY_SID, "⚠️ חוסר", is_dummy=True))

                tasks = [
                    Task(r[t_id_col], r[t_nm_col], r[t_req_col],
                         gval(r,shift_col,1), gval(r,rest_col,0),
                         gval(r,overlap_col,False), gval(r,hours_col,'all'),
                         gval(r,role_col,''), gval(r,intense_col,1),
                         gval(r,block_col,''))
                    for _, r in t_df.iterrows()
                ]

                for idx, ah in enumerate(st.session_state.get("adhoc_tasks",[])):
                    tid = _ADHOC_TID_START + idx
                    sh, eh = ah["start"], ah["end"]
                    hrs_str   = f"{sh}-{eh}" if sh<=eh else f"{sh}-23,0-{eh}"
                    shift_len = (eh-sh+1) if sh<=eh else (24-sh+eh+1)
                    tasks.append(Task(tid, ah["name"], ah["req"],
                                      max(1,shift_len), 0,
                                      "True" if ah["overlap"] else "False",
                                      hrs_str, '', ah["intensity"], '', is_adhoc=True))

                real_soldiers = [s for s in soldiers if not s.is_dummy]

                # ── פרסר הנחיות שנשמרו כטקסט גולמי ─────────────
                parsed_directives = []
                raw_parse_errors  = []
                for d in st.session_state.get("directives", []):
                    if d.soldier_name == "__RAW__":
                        raw_text = d.hours[0] if d.hours else ""
                        parsed, err = parse_free_text_directive(raw_text, real_soldiers, tasks)
                        if parsed:
                            parsed_directives.append(parsed)
                        else:
                            raw_parse_errors.append(f"«{raw_text}» — {err}")
                    else:
                        parsed_directives.append(d)

                if raw_parse_errors:
                    st.markdown(
                        '<div class="warn-box">⚠️ הנחיות שלא פורסרו:<br>' +
                        '<br>'.join(raw_parse_errors) + '</div>',
                        unsafe_allow_html=True)

                if parsed_directives:
                    dir_preview = "<br>".join(str(d) for d in parsed_directives)
                    st.markdown(
                        f'<div class="directive-box">👨‍✈️ <b>{len(parsed_directives)} הנחיות מפקד יוחלו:</b><br>{dir_preview}</div>',
                        unsafe_allow_html=True)

                # ── ניתוח זמינות ─────────────────────────────────
                with st.expander("🔍 ניתוח זמינות", expanded=True):
                    bottlenecks = []
                    for t in tasks:
                        eligible = [
                            s for s in real_soldiers
                            if t.tid not in s.exempt
                            and not any(r in t.block_roles for r in s.roles)
                        ]
                        bottlenecks.append({
                            "": "✅" if len(eligible)>=t.req else "⚠️",
                            "משימה": ("🔵 " if t.is_adhoc else "") + t.name,
                            "נדרש": t.req, "זמינים": len(eligible),
                            "מחסור": max(0, t.req-len(eligible))
                        })
                    st.dataframe(pd.DataFrame(bottlenecks), use_container_width=True, hide_index=True)

                adhoc_count = len(st.session_state.get("adhoc_tasks",[]))
                st.info(f"📋 {len(real_soldiers)} חיילים · {len(tasks)-adhoc_count} משימות"
                        + (f" + {adhoc_count} אד-הוק" if adhoc_count else "")
                        + (f" · {len(parsed_directives)} הנחיות מפקד" if parsed_directives else ""))

                # ── שלב א: גריידי ────────────────────────────────
                with st.spinner("⚡ שלב 1: שיבוץ גריידי..."):
                    schedule, dummy_slots, work_h, int_load = greedy_schedule(soldiers, tasks)

                greedy_dummy = sum(1 for t in tasks for si in range(len(t.slots))
                                   for h in range(24) if schedule[t.tid][si][h]==DUMMY_SID)
                st.markdown(f'<div class="info-box">⚡ גריידי: {greedy_dummy} שעות-סלוט עם רפאים</div>',
                            unsafe_allow_html=True)

                # ── שלב ב: CP-SAT ─────────────────────────────────
                directive_errors = []
                if use_cpsat:
                    with st.spinner(f"🔧 CP-SAT ({cpsat_time}ש׳) + הנחיות מפקד..."):
                        schedule, directive_errors = improve_with_cpsat(
                            soldiers, tasks, schedule,
                            directives=parsed_directives,
                            time_limit=cpsat_time)
                    st.markdown('<div class="info-box">🔧 CP-SAT הסתיים.</div>',
                                unsafe_allow_html=True)
                    if directive_errors:
                        st.markdown(
                            '<div class="warn-box">⚠️ הנחיות שלא הוחלו:<br>' +
                            '<br>'.join(directive_errors) + '</div>',
                            unsafe_allow_html=True)

                # ── ספירה סופית ──────────────────────────────────
                dummy_hours_count = sum(1 for t in tasks for si in range(len(t.slots))
                                        for h in range(24)
                                        if schedule[t.tid][si][h] == DUMMY_SID)
                if dummy_hours_count > 0:
                    st.markdown(
                        f'<div class="warn-box">⚠️ <b>{dummy_hours_count} שעות-סלוט ללא כיסוי</b></div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">✅ <b>כיסוי מלא!</b></div>',
                                unsafe_allow_html=True)

                # ── בנה תוצאה ────────────────────────────────────
                final_df = build_result_df(soldiers, tasks, schedule)
                if len(final_df) == 0:
                    st.markdown('<div class="error-box">❌ לא נוצרו שורות.</div>',
                                unsafe_allow_html=True)
                    st.stop()

                real_df   = final_df[~final_df["שם"].str.startswith("⚠️")]
                gap_h     = int(real_df["סך שעות"].max() - real_df["סך שעות"].min())
                avg_h     = real_df["סך שעות"].mean()
                avg_sleep = real_df["שעות שינה (22-08)"].mean()
                badge     = "✅ מצוין" if gap_h<=2 else ("⚠️ סביר" if gap_h<=5 else "❗ גבוה")

                st.markdown(f"""
                <div class="metric-row">
                  <div class="metric-card"><div class="mc-label">חיילים</div>
                    <div class="mc-value">{len(real_soldiers)}</div>
                    <div class="mc-sub">{len(tasks)} משימות</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שעות</div>
                    <div class="mc-value">{avg_h:.1f}</div><div class="mc-sub">לחייל</div></div>
                  <div class="metric-card"><div class="mc-label">פער הוגנות</div>
                    <div class="mc-value">{gap_h}</div><div class="mc-sub">{badge}</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שינה</div>
                    <div class="mc-value">{avg_sleep:.1f}</div><div class="mc-sub">יעד: 7.0</div></div>
                  <div class="metric-card"><div class="mc-label">שעות חסרות</div>
                    <div class="mc-value">{"0 ✅" if dummy_hours_count==0 else f"{dummy_hours_count} ⚠️"}</div>
                    <div class="mc-sub">{"כיסוי מלא" if dummy_hours_count==0 else "חסר"}</div></div>
                </div>""", unsafe_allow_html=True)

                # ── לוחות ────────────────────────────────────────
                st.markdown("---")
                st.subheader("📅 לוח לפי חייל")
                st.table(final_df)

                st.markdown("---")
                st.subheader("📋 לוח לפי משימה")
                st.table(build_task_df(soldiers, tasks, schedule))

                dl1, dl2 = st.columns(2)
                with dl1:
                    st.download_button("📥 Excel — לפי חיילים",
                                       data=to_excel_styled(final_df),
                                       file_name="Shavtzak_v20_Soldiers.xlsx",
                                       use_container_width=True)
                with dl2:
                    st.download_button("📥 Excel — לפי משימות",
                                       data=to_excel_task_view(soldiers, tasks, schedule),
                                       file_name="Shavtzak_v20_Tasks.xlsx",
                                       use_container_width=True)

                # ── גרף ──────────────────────────────────────────
                st.markdown("---")
                st.subheader("📊 ניתוח עומסים")
                fig = px.bar(real_df, x="שם", y="סך שעות", color="מדד עצימות",
                             color_continuous_scale=["#a8d5a2","#1a3d17"],
                             title="עומס שעות לחייל", text="סך שעות")
                fig.update_traces(textposition="outside", marker_line_width=0)
                fig.update_layout(plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                                  font=dict(family="Heebo",size=12),
                                  xaxis=dict(tickangle=-30,title=""),
                                  yaxis=dict(title='שעות סה"כ'),
                                  margin=dict(t=50,b=80,l=30,r=20))
                st.plotly_chart(fig, use_container_width=True)

                # ── המלצות לפתרון חוסרים ─────────────────────────
                if dummy_hours_count > 0:
                    st.markdown("---")
                    st.subheader("💡 המלצות המערכת לפתרון חוסרים")
                    recs = diagnose_dummy_slots(soldiers, tasks, schedule)
                    if recs:
                        for r in recs:
                            st.markdown(f"""
<div class="rec-box">
<b>🎯 משימה:</b> {r['משימה']} &nbsp;|&nbsp; <b>🕐 שעות:</b> {r['שעות']} &nbsp;|&nbsp; <b>סלוט:</b> {r['סלוט']}<br>
<b>📋 סיבה:</b> {r['סיבה']}<br>
<b>✅ המלצה:</b> {r['המלצה']}
</div>""", unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="info-box">ℹ️ לא נוצרו המלצות ספציפיות.</div>',
                                    unsafe_allow_html=True)

                # ── תובנות ───────────────────────────────────────
                with st.expander("💡 תובנות"):
                    max_s = real_df[real_df["סך שעות"]==real_df["סך שעות"].max()]["שם"].tolist()
                    st.markdown(f"""
**גריידי:** שיבץ {len(real_soldiers)} חיילים. False-tasks קודם.
**CP-SAT:** {"הרץ עם {len(parsed_directives)} הנחיות מפקד." if use_cpsat else "לא הופעל."}
**עמוסים:** {", ".join(max_s)} — {real_df["סך שעות"].max()} שעות.
**פער:** {gap_h} שעות — {"מצוין." if gap_h<=2 else "שקול להוסיף כוח אדם." if gap_h>4 else "סביר."}
**שינה:** ממוצע {avg_sleep:.1f} שעות.
{"**⚠️ חוסרים:** "+str(dummy_hours_count)+" שעות — ראה המלצות למעלה." if dummy_hours_count>0 else "**✅ כיסוי מלא.**"}
                    """)

            except Exception:
                st.markdown('<div class="error-box">🚨 שגיאה טכנית:</div>',
                            unsafe_allow_html=True)
                st.code(traceback.format_exc())
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.
        </div>""", unsafe_allow_html=True)
