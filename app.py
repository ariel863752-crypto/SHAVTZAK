import io
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
import traceback

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
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val, is_task=True) -> list:
    """
    מפרש שעות פעילות — תומך בפורמטים מורכבים.
    is_task=True  → תא ריק = 24 שעות (משימה פעילה כל היום)
    is_task=False → תא ריק = רשימה ריקה  (חייל לא חסום בכלל)
    """
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
            start_s = ps[0].strip()
            end_s   = ps[-1].strip()
            if start_s and end_s:
                s = to_hour(start_s)
                e = to_hour(end_s)
                if s <= e:
                    for h in range(s, e + 1):
                        res.add(h % 24)
                else:  # מעגלי: 22-6
                    for h in range(s, 24):
                        res.add(h)
                    for h in range(0, e + 1):
                        res.add(h)
        else:
            try:
                res.add(to_hour(part) % 24)
            except (ValueError, IndexError):
                pass

    return sorted(res)


# ── חייל רפאים (Dummy) ──────────────────────────────────────────
DUMMY_SID = "__DUMMY__"


class Soldier:
    def __init__(self, s_id, name, restr='', roles='', unavail='', is_dummy=False):
        self.sid      = str(s_id).replace('.0', '').strip()
        self.name     = str(name).strip()
        self.is_dummy = is_dummy
        self.exempt   = (
            [int(float(t)) for t in str(restr).split(',')
             if str(t).strip().replace('.0', '').isdigit()]
            if pd.notna(restr) and str(restr).strip() not in ('', 'nan') else []
        )
        self.roles = (
            [r.strip() for r in str(roles).split(',') if r.strip()]
            if pd.notna(roles) and str(roles).strip() not in ('', 'nan') else []
        )
        self.blocked_hours = [] if is_dummy else parse_time_ranges(unavail, is_task=False)

    def can_assign(self, task, shift_hours: list) -> bool:
        """האם החייל יכול לבצע משימה זו בכל שעות המשמרת?"""
        if self.is_dummy:
            return True
        if any(h in self.blocked_hours for h in shift_hours):
            return False
        if task.tid in self.exempt:
            return False
        if any(r in task.block_roles for r in self.roles):
            return False
        return True


class Task:
    def __init__(self, t_id, name, req, shift, rest, overlap, hours,
                 req_roles, intensity, block_roles=''):
        self.tid       = int(float(t_id))
        self.name      = str(name).strip()
        self.req       = int(req)
        self.shift     = int(shift)   if pd.notna(shift)   else 1
        self.rest      = int(rest)    if pd.notna(rest)    else 0
        self.overlap   = str(overlap).strip().lower() == 'true'
        self.hours     = parse_time_ranges(hours)
        self.intensity = int(intensity) if pd.notna(intensity) else 1
        self.block_roles = (
            [r.strip() for r in str(block_roles).split(',') if r.strip()]
            if pd.notna(block_roles) and str(block_roles).strip() not in ('', 'nan') else []
        )
        parsed = (
            [r.strip() for r in str(req_roles).split(',')]
            if pd.notna(req_roles) and str(req_roles).strip() not in ('', 'nan') else []
        )
        self.slots = parsed[:]
        while len(self.slots) < self.req:
            self.slots.append(None)

    def slot_ok(self, slot_idx: int, soldier) -> bool:
        if soldier.is_dummy:
            return True
        req_role = self.slots[slot_idx]
        if req_role is None:
            return True
        return req_role in soldier.roles

    def get_shift_starts(self) -> list:
        if not self.hours:
            return []
        sorted_hours = sorted(self.hours)
        starts = []
        i = 0
        while i < len(sorted_hours):
            start = sorted_hours[i]
            valid = True
            for k in range(self.shift):
                expected = (start + k) % 24
                if expected not in self.hours:
                    valid = False
                    break
            if valid:
                starts.append(start)
                i += self.shift
            else:
                i += 1
        return starts


# ══════════════════════════════════════════════════════════════════
# 3. Excel מעוצב
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df: pd.DataFrame, sheet_name='שבצ"ק', include_index=True) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        df.to_excel(w, index=include_index, sheet_name=sheet_name)
        wb, ws = w.book, w.sheets[sheet_name]
        hf = wb.add_format({'bold': True, 'fg_color': '#2d5a27', 'font_color': 'white',
                             'border': 1, 'align': 'right', 'valign': 'vcenter'})
        ef = wb.add_format({'fg_color': '#f0f8ef', 'align': 'right'})
        bf = wb.add_format({'align': 'right'})
        for ci, cv in enumerate(df.columns):
            ix = ci + (1 if include_index else 0)
            ws.write(0, ix, cv, hf)
            ws.set_column(ix, ix, min(max(df[cv].astype(str).map(len).max(), len(cv)) + 4, 40))
        for ri in range(1, len(df) + 1):
            row_fmt = ef if ri % 2 == 0 else bf
            ws.set_row(ri, None, row_fmt)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════
# 4. לב המערכת: גריידי חכם v16
#
#  תיקונים לעומת v15:
#  ✅ חפיפה סימטרית: חייל יכול להיות ב-2 משימות רק אם שתיהן overlap=True
#  ✅ is_free בודק גם את המשימות הקיימות של החייל (לא רק overlap של המשימה הנוכחית)
# ══════════════════════════════════════════════════════════════════

def greedy_schedule(soldiers: list, tasks: list, num_hours: int = 24):
    """
    מחזיר:
      schedule[tid][slot_idx][h] = sid | DUMMY_SID | None
      dummy_slots  — רשימת (task_name, slot, sh, eh) שאויישו ע"י רפאים
      work_hours   — {sid: מספר שעות}
      intensity_load — {sid: עומס עצימות}
    """
    schedule = {
        t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))}
        for t in tasks
    }

    # מעקב: לכל חייל ושעה — איזו משימות הוא משובץ בה (ומה overlap שלהן)
    assignments: dict[str, dict[int, list]] = {s.sid: {h: [] for h in range(num_hours)}
                                                for s in soldiers}
    assignments[DUMMY_SID] = {h: [] for h in range(num_hours)}

    busy_hours: dict[str, set] = {s.sid: set() for s in soldiers}
    busy_hours[DUMMY_SID] = set()

    work_hours:     dict[str, int] = {s.sid: 0 for s in soldiers}
    work_hours[DUMMY_SID]          = 0
    intensity_load: dict[str, int] = {s.sid: 0 for s in soldiers}
    intensity_load[DUMMY_SID]      = 0

    dummy_slots = []

    # ── v16: בדיקת חפיפה סימטרית ────────────────────────────────
    def is_free(sid: str, shift_hours: list, new_task: Task) -> bool:
        """
        חייל פנוי לשיבוץ במשמרת חדשה אם ורק אם:
        1. המשימה החדשה לא מאשרת חפיפה → חייב להיות ריק לגמרי בכל שעות המשמרת.
        2. המשימה החדשה מאשרת חפיפה → גם כל המשימות הקיימות של החייל בשעות אלו
           חייבות לאשר חפיפה.
        """
        for h in shift_hours:
            existing = assignments[sid][h]  # רשימת Task-objects בשעה h
            if not new_task.overlap:
                # המשימה החדשה אוסרת חפיפה — חייב להיות ריק לחלוטין
                if existing:
                    return False
            else:
                # המשימה החדשה מאשרת חפיפה — כל המשימות הקיימות חייבות לאשר גם הן
                for existing_task in existing:
                    if not existing_task.overlap:
                        return False
        return True

    def mark_busy(sid: str, shift_hours: list, rest: int, task_end_h: int, task: Task):
        for h in shift_hours:
            assignments[sid][h].append(task)
            if not task.overlap:
                busy_hours[sid].add(h)
        for i in range(1, rest + 1):
            busy_hours[sid].add((task_end_h + i) % num_hours)

    sorted_tasks = sorted(tasks, key=lambda t: (-t.req, -t.intensity, -t.shift))

    for t in sorted_tasks:
        shift_starts = t.get_shift_starts()

        for slot_idx in range(len(t.slots)):
            for start_h in shift_starts:
                shift_hours = [(start_h + i) % num_hours for i in range(t.shift)]
                end_h       = shift_hours[-1]

                real_candidates = []
                for s in soldiers:
                    if s.is_dummy:
                        continue
                    if not s.can_assign(t, shift_hours):
                        continue
                    if not t.slot_ok(slot_idx, s):
                        continue
                    # v16: is_free מקבל את אובייקט המשימה כדי לבדוק חפיפה סימטרית
                    if not is_free(s.sid, shift_hours, t):
                        continue
                    score = work_hours[s.sid] * 10 + intensity_load[s.sid]
                    real_candidates.append((score, s.sid, s))

                if real_candidates:
                    real_candidates.sort(key=lambda x: x[0])
                    _, chosen_sid, chosen_s = real_candidates[0]
                    chosen_is_dummy = False
                else:
                    chosen_sid      = DUMMY_SID
                    chosen_is_dummy = True

                for h in shift_hours:
                    schedule[t.tid][slot_idx][h] = chosen_sid

                if chosen_is_dummy:
                    dummy_slots.append((t.name, slot_idx + 1, start_h, end_h))
                else:
                    work_hours[chosen_sid]     += t.shift
                    intensity_load[chosen_sid] += t.shift * t.intensity
                    mark_busy(chosen_sid, shift_hours, t.rest, end_h, t)

    return schedule, dummy_slots, work_hours, intensity_load


# ══════════════════════════════════════════════════════════════════
# 5. בניית DataFrame תוצאה
# ══════════════════════════════════════════════════════════════════
def build_result_df(soldiers: list, tasks: list, schedule: dict,
                    num_hours: int = 24) -> pd.DataFrame:
    SLEEP_HOURS = set(range(22, 24)) | set(range(0, 9))
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]

    rows = []
    for s in soldiers:
        if s.is_dummy:
            continue
        row      = {"שם": s.name}
        total    = 0
        night    = 0
        intensity = 0

        for h in range(num_hours):
            active = []
            for t in tasks:
                for slot_idx in range(len(t.slots)):
                    assigned = schedule[t.tid][slot_idx][h]
                    if assigned == s.sid:
                        active.append(t.name)
                        if not t.overlap:
                            if h in SLEEP_HOURS:
                                night += 1
                            intensity += t.intensity
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if active:
                total += 1

        row["סך שעות"]          = total
        row["מדד עצימות"]       = intensity
        row["שעות שינה (22-08)"] = len(SLEEP_HOURS) - night
        rows.append(row)

    # שורת "חוסר"
    dummy_row = {"שם": "⚠️ חוסר כוח אדם"}
    for h in range(num_hours):
        dummy_slots_h = []
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                if schedule[t.tid][slot_idx][h] == DUMMY_SID:
                    dummy_slots_h.append(t.name)
        dummy_row[hour_labels[h]] = "❌ " + " + ".join(dummy_slots_h) if dummy_slots_h else "—"
    dummy_row["סך שעות"] = sum(
        1 for h in range(num_hours)
        for t in tasks for si in range(len(t.slots))
        if schedule[t.tid][si][h] == DUMMY_SID
    )
    dummy_row["מדד עצימות"]       = "—"
    dummy_row["שעות שינה (22-08)"] = "—"

    has_dummy = any(
        schedule[t.tid][si][h] == DUMMY_SID
        for t in tasks for si in range(len(t.slots)) for h in range(num_hours)
    )

    df = pd.DataFrame(rows)
    if has_dummy:
        df = pd.concat([df, pd.DataFrame([dummy_row])], ignore_index=True)

    df.index = range(1, len(df) + 1)
    return df


# ══════════════════════════════════════════════════════════════════
# 6. שיפור CP-SAT (אופציונלי) — v16
#
#  תיקון עיקרי: אילוץ חפיפה סימטרי
#  אם חייל משובץ למשימה non-overlap בשעה h → לא יכול להיות בשום משימה אחרת
#  אם חייל משובץ למשימות overlap בלבד → יכול להצטבר
# ══════════════════════════════════════════════════════════════════

def improve_with_cpsat(soldiers: list, tasks: list, schedule: dict,
                       num_hours: int = 24, time_limit: float = 60.0):
    real_soldiers = [s for s in soldiers if not s.is_dummy]
    model         = cp_model.CpModel()

    x = {}
    for s in real_soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.sid, t.tid, si, h] = model.NewBoolVar(
                        f"x_{s.sid}_{t.tid}_{si}_{h}")

    for s in real_soldiers:
        for t in tasks:
            for si in range(len(t.slots)):
                for h in range(num_hours):
                    hint = 1 if schedule[t.tid][si][h] == s.sid else 0
                    model.AddHint(x[s.sid, t.tid, si, h], hint)

    dummy_vars = {}
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h in t.hours:
                    dv = model.NewBoolVar(f"dummy_{t.tid}_{si}_{h}")
                    dummy_vars[t.tid, si, h] = dv
                    model.AddHint(dv, 1 if schedule[t.tid][si][h] == DUMMY_SID else 0)

    # אילוץ כיסוי
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h in t.hours:
                    model.Add(
                        sum(x[s.sid, t.tid, si, h] for s in real_soldiers)
                        + dummy_vars[t.tid, si, h] == 1
                    )
                else:
                    for s in real_soldiers:
                        model.Add(x[s.sid, t.tid, si, h] == 0)

    # אילוצי כשירות
    for s in real_soldiers:
        for t in tasks:
            if t.tid in s.exempt or any(r in t.block_roles for r in s.roles):
                for si in range(len(t.slots)):
                    for h in range(num_hours):
                        model.Add(x[s.sid, t.tid, si, h] == 0)
                continue
            for si, req_role in enumerate(t.slots):
                if req_role and req_role not in s.roles:
                    for h in range(num_hours):
                        model.Add(x[s.sid, t.tid, si, h] == 0)
        for h in s.blocked_hours:
            for t in tasks:
                for si in range(len(t.slots)):
                    model.Add(x[s.sid, t.tid, si, h] == 0)

    # ── v16: אילוץ חפיפה סימטרי ──────────────────────────────────
    # כלל: אם חייל משובץ לאיזשהי משימה non-overlap בשעה h,
    #       הוא לא יכול להיות בשום משימה אחרת באותה שעה.
    # כלל שקול: סכום כל המשימות non-overlap ≤ 1,
    #            ואם יש כזו → סכום כל המשימות (כולל overlap) ≤ 1.
    for s in real_soldiers:
        for h in range(num_hours):
            non_ov_vars = [
                x[s.sid, t.tid, si, h]
                for t in tasks if not t.overlap
                for si in range(len(t.slots))
            ]
            all_vars_h = [
                x[s.sid, t.tid, si, h]
                for t in tasks
                for si in range(len(t.slots))
            ]

            if not non_ov_vars:
                continue

            # אסור ל-non-overlap להצטבר בין עצמן
            model.Add(sum(non_ov_vars) <= 1)

            # אם יש non-overlap → אסור overlap נוספות
            # מומש ע"י: אם non-overlap_i = 1, כל שאר המשימות = 0
            # נשתמש ב: sum(all_vars) <= 1 בתנאי שיש non-overlap
            # שקול ל: sum(all_vars) * 1 <= 1 + (num_overlap_tasks)*(1 - sum(non_ov))
            # גישה פשוטה וחזקה:
            # לכל משימה non-overlap (t1) ולכל משימה אחרת (t2) —
            # לא יכולים שניהם להיות 1 בו-זמנית
            for t1 in tasks:
                if t1.overlap:
                    continue
                for si1 in range(len(t1.slots)):
                    for t2 in tasks:
                        if t2.tid == t1.tid:
                            continue
                        for si2 in range(len(t2.slots)):
                            # t1 non-overlap + t2 כלשהי ← אסור
                            model.Add(
                                x[s.sid, t1.tid, si1, h] + x[s.sid, t2.tid, si2, h] <= 1
                            )

    # פונקציית מטרה
    PENALTY_DUMMY    = 100_000
    PENALTY_REST     =     500
    PENALTY_NIGHT    =     200
    PENALTY_FAIRNESS =     100

    penalties = []

    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h in t.hours:
                    penalties.append(dummy_vars[t.tid, si, h] * PENALTY_DUMMY)

    total_hours_vars = []
    for s in real_soldiers:
        th = model.NewIntVar(0, num_hours * len(tasks), f'th_{s.sid}')
        model.Add(th == sum(
            x[s.sid, t.tid, si, h]
            for t in tasks for si in range(len(t.slots)) for h in range(num_hours)
        ))
        total_hours_vars.append(th)

    max_h = model.NewIntVar(0, 1000, 'max_h')
    min_h = model.NewIntVar(0, 1000, 'min_h')
    model.AddMaxEquality(max_h, total_hours_vars)
    model.AddMinEquality(min_h, total_hours_vars)
    diff = model.NewIntVar(0, 1000, 'diff')
    model.Add(diff == max_h - min_h)
    penalties.append(diff * PENALTY_FAIRNESS)

    for s in real_soldiers:
        for t in tasks:
            if t.overlap or t.rest == 0:
                continue
            shift_starts = t.get_shift_starts()
            for start_h in shift_starts:
                worked_in_shift = sum(
                    x[s.sid, t.tid, si, start_h] for si in range(len(t.slots))
                )
                is_working_shift = model.NewBoolVar(
                    f'is_working_{s.sid}_{t.tid}_{start_h}')
                model.Add(worked_in_shift >= 1).OnlyEnforceIf(is_working_shift)
                model.Add(worked_in_shift == 0).OnlyEnforceIf(is_working_shift.Not())

                end_h = (start_h + t.shift - 1) % num_hours

                for r in range(1, t.rest + 1):
                    rest_h = (end_h + r) % num_hours
                    rest_work = sum(
                        x[s.sid, t2.tid, si2, rest_h]
                        for t2 in tasks if not t2.overlap
                        for si2 in range(len(t2.slots))
                    )
                    is_working_in_rest = model.NewBoolVar(
                        f'work_rest_{s.sid}_{t.tid}_{start_h}_{r}')
                    model.Add(rest_work >= 1).OnlyEnforceIf(is_working_in_rest)
                    model.Add(rest_work == 0).OnlyEnforceIf(is_working_in_rest.Not())

                    viol = model.NewBoolVar(
                        f'rest_viol_{s.sid}_{t.tid}_{start_h}_{r}')
                    model.AddBoolAnd(
                        [is_working_shift, is_working_in_rest]
                    ).OnlyEnforceIf(viol)
                    model.AddBoolOr(
                        [is_working_shift.Not(), is_working_in_rest.Not()]
                    ).OnlyEnforceIf(viol.Not())
                    penalties.append(viol * PENALTY_REST)

    NIGHT_HOURS = list(set(range(22, 24)) | set(range(0, 9)))
    NIGHT_LIMIT = 4
    for s in real_soldiers:
        night_work_vars = [
            x[s.sid, t.tid, si, h]
            for t in tasks if not t.overlap
            for si in range(len(t.slots))
            for h in NIGHT_HOURS
        ]
        if night_work_vars:
            night_total = model.NewIntVar(0, len(NIGHT_HOURS) * len(tasks), f'night_{s.sid}')
            model.Add(night_total == sum(night_work_vars))
            night_excess = model.NewIntVar(0, 100, f'night_exc_{s.sid}')
            model.Add(night_excess >= night_total - NIGHT_LIMIT)
            model.Add(night_excess >= 0)
            penalties.append(night_excess * PENALTY_NIGHT)

    max_penalty = 24 * sum(len(t.slots) for t in tasks) * PENALTY_DUMMY + 500_000
    total_penalty = model.NewIntVar(0, max_penalty, 'total_penalty')
    model.Add(total_penalty == sum(penalties))
    model.Minimize(total_penalty)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds  = time_limit
    solver.parameters.num_search_workers   = 6
    solver.parameters.log_search_progress  = False
    solver.parameters.relative_gap_limit   = 0.05
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return schedule

    new_schedule = {
        t.tid: {si: {h: None for h in range(num_hours)} for si in range(len(t.slots))}
        for t in tasks
    }
    for t in tasks:
        for si in range(len(t.slots)):
            for h in range(num_hours):
                if h not in t.hours:
                    continue
                if solver.Value(dummy_vars[t.tid, si, h]) == 1:
                    new_schedule[t.tid][si][h] = DUMMY_SID
                else:
                    assigned = None
                    for s in real_soldiers:
                        if solver.Value(x[s.sid, t.tid, si, h]) == 1:
                            assigned = s.sid
                            break
                    new_schedule[t.tid][si][h] = assigned

    return new_schedule


# ══════════════════════════════════════════════════════════════════
# 7. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v16)</h1>
  <p>גריידי מהיר + CP-SAT חכם · חפיפה סימטרית · בדיקת כל שעות המשמרת · תמיד מחזיר לוח מלא</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀  ביצוע שיבוץ", "📖  מדריך", "📥  תבניות"])

# ── תבניות ──────────────────────────────────────────────────────
with tab_templates:
    s_ex = pd.DataFrame({
        'מספר אישי':  [1001, 1002, 1003, 1004],
        'שם מלא':     ['ישראל ישראלי', 'יוסי כהן', 'אבי לוי', 'רון גל'],
        'פטורים':     ['', '101', '', ''],
        'הסמכות':     ['', '', '', ''],
        'שעות חסימה': ['', '10-14', '', '22-6'],
    })
    t_ex = pd.DataFrame({
        'מס"ד משימה':               [101, 102, 103],
        'שם המשימה':                ['שמירה', 'סיור', 'כוננות'],
        'סד"כ נדרש למשימה':        [2, 2, 1],
        'משך משמרת':                [4, 6, 24],
        'שעות מנוחה בין משימות':    [8, 8, 0],
        'אישור חפיפה בין משימות':   [False, False, True],
        'שעות פעילות':              ['all', 'all', 'all'],
        'הסמכה נדרשת':              ['', '', ''],
        'דירוג עצימות משימה (1-3)': [2, 3, 1],
        'תפקידים חסומים':           ['', '', ''],
    })
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**👥 תבנית חיילים**")
        st.dataframe(s_ex, use_container_width=True, hide_index=True)
        st.download_button("⬇️ הורד תבנית חיילים",
                           data=to_excel_styled(s_ex, "Soldiers", False),
                           file_name="Soldiers_v16.xlsx", use_container_width=True)
    with c2:
        st.markdown("**📋 תבנית משימות**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button("⬇️ הורד תבנית משימות",
                           data=to_excel_styled(t_ex, "Tasks", False),
                           file_name="Tasks_v16.xlsx", use_container_width=True)

# ── מדריך ───────────────────────────────────────────────────────
with tab_guide:
    st.markdown("""
### 📖 מדריך v16

#### קובץ חיילים
| עמודה | הסבר | דוגמה |
|---|---|---|
| מספר אישי | מזהה ייחודי | 1001 |
| שם מלא | שם להצגה | ישראל ישראלי |
| פטורים | מס"ד משימות חסומות | 101,103 |
| הסמכות | תפקידים (אם יש) | נהג,מפקד |
| שעות חסימה | שעות לא זמין | 10-14 / 22-6 / 08:30-16:00 |

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
| הסמכה נדרשת | תפקיד ספציפי | נהג |
| דירוג עצימות | 1-3 | 2 |
| תפקידים חסומים | מי לא יכול | קצין |
""")
    st.markdown("""
<div class="info-box">
<b>💡 v16 — שינויים עיקריים לעומת v15:</b><br>
<b>✅ חפיפה סימטרית (תיקון קריטי):</b> חייל יכול להיות בשתי משימות במקביל <em>רק אם שתיהן</em> מוגדרות
<code>overlap=True</code>. אם משימה אחת מוגדרת <code>False</code> — היא חוסמת את כל שאר המשימות,
גם אם הן עצמן מאשרות חפיפה.<br>
<b>✅ גריידי:</b> <code>is_free</code> בודק את אובייקט המשימה (לא רק דגל overlap) ומסתכל על כל המשימות
הקיימות של החייל בשעה זו.<br>
<b>✅ CP-SAT:</b> אילוץ חדש — לכל משימה non-overlap ושעה, נוסף אילוץ פרוד עם כל משימה אחרת (כולל
overlap), כך שהמודל לא יכול לשבץ חייל לשתיהן בו-זמנית.
</div>
""", unsafe_allow_html=True)

# ── ביצוע שיבוץ ─────────────────────────────────────────────────
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    with st.expander("⚙️ הגדרות מתקדמות"):
        use_cpsat  = st.toggle("שיפור CP-SAT אחרי גריידי (משפר הוגנות)", value=True)
        cpsat_time = st.slider("זמן CP-SAT (שניות)", 10, 120, 45, 5)
        st.markdown(
            '<div class="info-box" style="font-size:12px">'
            'הגריידי <b>תמיד</b> רץ ראשון ומחזיר תוצאה. '
            'CP-SAT רק משפר הוגנות ומנוחה — לא ישנה כיסוי.</div>',
            unsafe_allow_html=True)

    if sf and tf:
        if st.button('⚙️ צור שבצ"ק חכם (v16)', use_container_width=True, key="run_btn"):
            try:
                s_df = pd.read_excel(sf)
                t_df = pd.read_excel(tf)

                def find_col(df, keywords):
                    for kw in keywords:
                        matches = [c for c in df.columns if kw in str(c)]
                        if matches:
                            return matches[0]
                    return None

                id_col    = find_col(s_df, ['מספר', 'אישי'])
                name_col  = find_col(s_df, ['שם'])
                t_id_col  = find_col(t_df, ['מס"ד', 'מסד'])
                t_nm_col  = find_col(t_df, ['שם'])
                t_req_col = find_col(t_df, ['סד"כ', 'נדרש'])

                if not all([id_col, name_col, t_id_col, t_nm_col, t_req_col]):
                    st.markdown(
                        '<div class="error-box">❌ לא נמצאו עמודות חובה. '
                        'ודא שהקבצים לפי התבנית.</div>',
                        unsafe_allow_html=True)
                    st.stop()

                s_df = s_df.dropna(subset=[id_col, name_col])
                t_df = t_df.dropna(subset=[t_id_col, t_nm_col, t_req_col])

                def gcol(df, keys):
                    return find_col(df, keys)

                shift_col   = gcol(t_df, ['משך'])
                rest_col    = gcol(t_df, ['מנוחה'])
                overlap_col = gcol(t_df, ['חפיפה'])
                hours_col   = gcol(t_df, ['פעילות', 'שעות'])
                role_col    = gcol(t_df, ['הסמכה'])
                intense_col = gcol(t_df, ['עצימות', 'דירוג'])
                block_col   = gcol(t_df, ['חסומים'])

                def gval(row, col, default=''):
                    return row.get(col, default) if col else default

                soldiers = [
                    Soldier(r[id_col], r[name_col],
                            r.get('פטורים', ''), r.get('הסמכות', ''),
                            r.get('שעות חסימה', ''))
                    for _, r in s_df.iterrows()
                ]
                soldiers.append(Soldier(DUMMY_SID, "⚠️ חוסר", is_dummy=True))

                tasks = [
                    Task(r[t_id_col], r[t_nm_col], r[t_req_col],
                         gval(r, shift_col,   1),
                         gval(r, rest_col,    0),
                         gval(r, overlap_col, False),
                         gval(r, hours_col,   'all'),
                         gval(r, role_col,    ''),
                         gval(r, intense_col, 1),
                         gval(r, block_col,   ''))
                    for _, r in t_df.iterrows()
                ]

                real_soldiers = [s for s in soldiers if not s.is_dummy]
                st.info(f"📋 נטענו {len(real_soldiers)} חיילים ו-{len(tasks)} משימות")

                with st.spinner("⚡ שלב 1: שיבוץ גריידי..."):
                    schedule, dummy_slots, work_h, int_load = greedy_schedule(
                        soldiers, tasks)

                if use_cpsat:
                    with st.spinner(f"🔧 שלב 2: שיפור CP-SAT ({cpsat_time}ש׳)..."):
                        schedule = improve_with_cpsat(
                            soldiers, tasks, schedule, time_limit=cpsat_time)
                    st.markdown(
                        '<div class="info-box">🔧 CP-SAT הסתיים — שיפור הוגנות/מנוחה.</div>',
                        unsafe_allow_html=True)

                dummy_hours_count = sum(
                    1 for t in tasks for si in range(len(t.slots))
                    for h in range(24) if schedule[t.tid][si][h] == DUMMY_SID
                )

                if dummy_hours_count > 0:
                    st.markdown(
                        f'<div class="warn-box">⚠️ <b>{dummy_hours_count} שעות-סלוט ללא כיסוי</b> '
                        f'— שובצו לחייל רפאים. ראה שורת "⚠️ חוסר" בטבלה.</div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        '<div class="info-box">✅ <b>כיסוי מלא</b> — כל העמדות אויישו!</div>',
                        unsafe_allow_html=True)

                final_df = build_result_df(soldiers, tasks, schedule)

                if len(final_df) == 0:
                    st.markdown(
                        '<div class="error-box">❌ לא נוצרו שורות — בדוק את קבצי הקלט.</div>',
                        unsafe_allow_html=True)
                    st.stop()

                real_df   = final_df[~final_df["שם"].str.startswith("⚠️")]
                gap_h     = int(real_df["סך שעות"].max() - real_df["סך שעות"].min())
                avg_h     = real_df["סך שעות"].mean()
                avg_sleep = real_df["שעות שינה (22-08)"].mean()
                badge     = ("✅ מצוין" if gap_h <= 2
                             else ("⚠️ סביר" if gap_h <= 5 else "❗ גבוה"))

                st.markdown(f"""
                <div class="metric-row">
                  <div class="metric-card"><div class="mc-label">חיילים</div>
                    <div class="mc-value">{len(real_soldiers)}</div>
                    <div class="mc-sub">{len(tasks)} משימות</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שעות</div>
                    <div class="mc-value">{avg_h:.1f}</div>
                    <div class="mc-sub">לחייל</div></div>
                  <div class="metric-card"><div class="mc-label">פער הוגנות</div>
                    <div class="mc-value">{gap_h}</div>
                    <div class="mc-sub">{badge}</div></div>
                  <div class="metric-card"><div class="mc-label">ממוצע שינה</div>
                    <div class="mc-value">{avg_sleep:.1f}</div>
                    <div class="mc-sub">יעד: 7.0</div></div>
                  <div class="metric-card"><div class="mc-label">שעות חסרות</div>
                    <div class="mc-value">{"0 ✅" if dummy_hours_count == 0 else f"{dummy_hours_count} ⚠️"}</div>
                    <div class="mc-sub">{"כיסוי מלא" if dummy_hours_count == 0 else "חסר כוח אדם"}</div></div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📅 לוח השיבוץ הסופי")
                st.table(final_df)
                st.download_button("📥 הורד לוח שיבוץ (Excel)",
                                   data=to_excel_styled(final_df),
                                   file_name="Final_Shavtzak_v16.xlsx",
                                   use_container_width=True)

                st.markdown("---")
                st.subheader("📊 ניתוח עומסים")
                chart_df = real_df[["שם", "סך שעות", "מדד עצימות"]].copy()
                fig = px.bar(
                    chart_df, x="שם", y="סך שעות", color="מדד עצימות",
                    color_continuous_scale=["#a8d5a2", "#1a3d17"],
                    title="עומס שעות לחייל", text="סך שעות")
                fig.update_traces(textposition="outside", marker_line_width=0)
                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="Heebo", size=12),
                    xaxis=dict(tickangle=-30, title=""),
                    yaxis=dict(title='שעות סה"כ'),
                    margin=dict(t=50, b=80, l=30, r=20))
                st.plotly_chart(fig, use_container_width=True)

                with st.expander("💡 תובנות ופירוט"):
                    max_s = real_df[real_df["סך שעות"] == real_df["סך שעות"].max()]["שם"].tolist()
                    st.markdown(f"""
**שלב גריידי:** שיבץ {len(real_soldiers)} חיילים לפי עומס מינימלי.
**שיפור CP-SAT:** {"שיפר הוגנות / מנוחה / שינה בתוך 5% מהאופטימום." if use_cpsat else "לא הופעל."}
**חיילים עמוסים:** {", ".join(max_s)} — {real_df["סך שעות"].max()} שעות.
**פער הוגנות:** {gap_h} שעות — {"מצוין." if gap_h <= 2 else "מומלץ להוסיף חיילים." if gap_h > 4 else "סביר."}
**שינה:** ממוצע {avg_sleep:.1f} שעות.
{"**⚠️ שעות ללא כיסוי:** " + str(dummy_hours_count) + " — הוסף חיילים או קצר מנוחות." if dummy_hours_count > 0 else "**✅ כיסוי מלא** — כל העמדות מאויישות."}
                    """)

                    if dummy_slots:
                        st.markdown("**פירוט חוסרים:**")
                        for task_name, slot, sh, eh in dummy_slots:
                            st.markdown(f"  • {task_name} — סלוט {slot}, שעות {sh:02d}:00–{eh:02d}:59")

            except Exception:
                st.markdown(
                    '<div class="error-box">🚨 שגיאה טכנית:</div>',
                    unsafe_allow_html=True)
                st.code(traceback.format_exc())
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין תבניות? לחצו על הטאב <b>תבניות</b>.
        </div>
        """, unsafe_allow_html=True)--- דעתך?
