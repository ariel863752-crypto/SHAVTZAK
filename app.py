import io
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# 1. הגדרות עמוד
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

# ══════════════════════════════════════════════════════════════════
# 2. CSS — RTL מלא
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

html, body { direction: rtl; }

.stApp, [data-testid="stAppViewContainer"], .main, .block-container,
.stMarkdown, p, span, li, label, div,
[data-testid="stText"], [data-testid="stMarkdownContainer"],
[data-testid="stAlert"], [data-testid="stExpander"] summary,
[data-testid="stFileUploader"] label, [data-testid="stFileUploader"] div,
[data-testid="stSlider"] label, [data-testid="stSelectbox"] label {
    font-family: 'Heebo', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
}

.stApp { background: #f2f5f2; }
.block-container { padding: 2rem 2.5rem 3rem; max-width: 1400px; }

.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px; padding: 30px 35px; margin-bottom: 28px;
    box-shadow: 0 4px 20px rgba(45,90,39,0.25); text-align: right;
}
.app-header h1 {
    font-size: clamp(24px, 4vw, 42px); font-weight: 900; color: white;
    margin: 0 0 8px 0; letter-spacing: -0.5px;
}
.app-header p { font-size: 16px; color: rgba(255,255,255,0.88); margin: 0; }

[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important; justify-content: flex-start !important;
    gap: 6px; background: white; border-radius: 12px; padding: 5px;
    border: 1px solid #dde8dc; margin-bottom: 20px;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px; font-weight: 600; font-size: 14px;
    padding: 8px 20px; color: #5a7a57; direction: rtl;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: #2d5a27 !important; color: white !important;
}

.metric-row {
    display: flex; flex-direction: row-reverse;
    gap: 16px; margin: 22px 0; flex-wrap: wrap;
}
.metric-card {
    flex: 1; min-width: 160px; background: white; border-radius: 14px;
    padding: 22px; border: 1px solid #dde8dc;
    box-shadow: 0 2px 8px rgba(45,90,39,0.07); text-align: right; direction: rtl;
}
.mc-label {
    font-size: 11px; color: #7a9a77; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px;
}
.mc-value { font-size: 34px; font-weight: 900; color: #1a3d17; line-height: 1; }
.mc-sub   { font-size: 12px; color: #a0b89d; margin-top: 4px; }

div.stButton > button:first-child {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important; font-weight: 700 !important; font-size: 17px !important;
    border-radius: 10px !important; border: none !important;
    height: 3.4em; width: 100%;
    box-shadow: 0 4px 14px rgba(45,90,39,0.3); transition: all 0.18s !important;
}
div.stButton > button:first-child:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(45,90,39,0.4) !important;
}
[data-testid="stDownloadButton"] > button {
    background: #b84d00 !important; color: white !important;
    font-weight: 600 !important; border-radius: 10px !important; border: none !important;
}

[data-testid="stFileUploader"] {
    background: white; border-radius: 12px; padding: 14px 16px;
    border: 2px dashed #c0d8bc; direction: rtl; text-align: right;
}
[data-testid="stFileUploader"]:hover { border-color: #2d5a27; }

[data-testid="stTable"] table {
    width: 100%; border-collapse: collapse; font-size: 12.5px;
    background: white; direction: rtl;
}
[data-testid="stTable"] th {
    background: #2d5a27 !important; color: white !important;
    padding: 10px 13px !important; font-weight: 600 !important;
    font-size: 12px !important; text-align: right !important; direction: rtl !important;
}
[data-testid="stTable"] td {
    padding: 9px 13px !important; border-bottom: 1px solid #f0f0f0 !important;
    text-align: right !important; direction: rtl !important;
}
[data-testid="stTable"] tr:nth-child(even) td { background: #f8fcf7 !important; }
[data-testid="stTable"] tr:hover td { background: #f0f8ef !important; }

.guide-table {
    width: 100%; border-collapse: separate; border-spacing: 0;
    border-radius: 12px; overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07); margin: 16px 0;
    font-size: 14px; direction: rtl;
}
.guide-table th {
    background: #2d5a27; color: white; padding: 13px 16px;
    text-align: right; font-weight: 700; font-size: 13px;
}
.guide-table td {
    padding: 12px 16px; border-bottom: 1px solid #eef3ed;
    background: white; line-height: 1.7; vertical-align: top; text-align: right;
}
.guide-table tr:last-child td { border-bottom: none; }
.guide-table tr:hover td { background: #f5fbf4; }

.info-box {
    background: #edf5ec; border-right: 5px solid #2d5a27;
    padding: 14px 18px; margin: 14px 0; font-size: 14px;
    color: #1a3d17; line-height: 1.8; direction: rtl;
    text-align: right; border-radius: 0 10px 10px 0;
}
.warn-box {
    background: #fff8e6; border-right: 5px solid #e67e22;
    padding: 14px 18px; margin: 14px 0; font-size: 14px;
    color: #7a4500; line-height: 1.8; direction: rtl;
    text-align: right; border-radius: 0 10px 10px 0;
}
.error-box {
    background: #fdecea; border-right: 5px solid #c0392b;
    padding: 14px 18px; margin: 14px 0; font-size: 14px;
    color: #7a0010; line-height: 1.8; direction: rtl;
    text-align: right; border-radius: 0 10px 10px 0;
}

[data-testid="stExpander"] summary {
    direction: rtl; text-align: right;
    font-family: 'Heebo', sans-serif; font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 3. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val) -> list:
    if pd.isna(val) or str(val).strip().lower() in ('all', '', 'nan'):
        return list(range(24))
    res = set()
    for part in str(val).split(','):
        part = part.strip()
        if '-' in part:
            parts = part.split('-')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                s, e = int(parts[0]), int(parts[1])
                if s <= e:
                    for h in range(s, e + 1):
                        res.add(h % 24)
                else:
                    for h in range(s, 24):
                        res.add(h % 24)
                    for h in range(0, e + 1):
                        res.add(h % 24)
        elif part.isdigit():
            res.add(int(part) % 24)
    return list(res)


class Soldier:
    def __init__(self, s_id, name, restr="", roles="", unavail=""):
        self.soldier_id = str(s_id)
        self.name       = str(name).strip()
        self.restricted_tasks = (
            [int(float(t)) for t in str(restr).split(',')
             if str(t).strip().replace('.0', '').isdigit()]
            if pd.notna(restr) and str(restr).strip() not in ("", "nan") else []
        )
        self.roles = (
            [r.strip() for r in str(roles).split(',') if r.strip()]
            if pd.notna(roles) and str(roles).strip() not in ("", "nan") else []
        )
        self.unavail_hours = parse_time_ranges(unavail)


class Task:
    def __init__(self, t_id, name, req_p, shift_dur, rest_dur,
                 overlap, hours, req_roles, intensity, blocked_roles=""):
        self.task_id            = int(t_id)
        self.name               = str(name).strip()
        self.required_personnel = int(req_p)
        self.shift_duration     = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration      = int(rest_dur)  if pd.notna(rest_dur)  else 0
        self.allow_overlap      = str(overlap).strip().lower() == 'true'
        self.active_hours       = parse_time_ranges(hours)
        self.intensity          = int(intensity) if pd.notna(intensity) else 1
        self.blocked_roles      = (
            [r.strip() for r in str(blocked_roles).split(',') if r.strip()]
            if pd.notna(blocked_roles) and str(blocked_roles).strip() not in ("", "nan") else []
        )
        parsed_roles = (
            [r.strip() for r in str(req_roles).split(',')]
            if pd.notna(req_roles) and str(req_roles).strip() not in ("", "nan") else []
        )
        self.slots = parsed_roles.copy()
        while len(self.slots) < self.required_personnel:
            self.slots.append(None)


# ══════════════════════════════════════════════════════════════════
# 4. Excel מעוצב
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df: pd.DataFrame, sheet_name: str = 'שבצ"ק',
                    include_index: bool = True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        wb, ws = writer.book, writer.sheets[sheet_name]
        hfmt = wb.add_format({'bold': True, 'fg_color': '#2d5a27', 'font_color': 'white',
                               'border': 1, 'align': 'right', 'valign': 'vcenter'})
        efmt = wb.add_format({'fg_color': '#f0f8ef', 'align': 'right'})
        bfmt = wb.add_format({'align': 'right'})
        for col_num, col_val in enumerate(df.columns.values):
            col_idx = col_num + (1 if include_index else 0)
            ws.write(0, col_idx, col_val, hfmt)
            col_len = max(df[col_val].astype(str).map(len).max(), len(col_val)) + 5
            ws.set_column(col_idx, col_idx, min(col_len, 40))
        for row_num in range(1, len(df) + 1):
            ws.set_row(row_num, None, efmt if row_num % 2 == 0 else bfmt)
        if include_index:
            ws.set_column(0, 0, 5)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════
# 5. Greedy hint generator — מאיץ את מציאת FEASIBLE הראשון
# ══════════════════════════════════════════════════════════════════
def build_greedy_hints(soldiers, tasks, num_hours=24):
    hints = {}
    busy_until = {s.soldier_id: -1 for s in soldiers}
    last_task_end = {s.soldier_id: -1 for s in soldiers}

    for s in soldiers:
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                for h in range(num_hours):
                    hints[(s.soldier_id, t.task_id, slot_idx, h)] = 0

    soldier_work_hours = {s.soldier_id: 0 for s in soldiers}

    for t in tasks:
        for slot_idx, req_role in enumerate(t.slots):
            h_list = sorted(t.active_hours)
            shifts_start = []
            if h_list:
                run_start = h_list[0]
                for i in range(1, len(h_list)):
                    if h_list[i] != h_list[i-1] + 1:
                        shifts_start.append(run_start)
                        run_start = h_list[i]
                shifts_start.append(run_start)

            for h in h_list:
                best = None
                best_load = 999999
                for s in soldiers:
                    if h in s.unavail_hours:
                        continue
                    if t.task_id in s.restricted_tasks:
                        continue
                    if any(role in t.blocked_roles for role in s.roles):
                        continue
                    if req_role is not None and req_role not in s.roles:
                        continue
                    if busy_until[s.soldier_id] >= h and not t.allow_overlap:
                        continue
                    if soldier_work_hours[s.soldier_id] < best_load:
                        best_load = soldier_work_hours[s.soldier_id]
                        best = s
                if best is not None:
                    hints[(best.soldier_id, t.task_id, slot_idx, h)] = 1
                    soldier_work_hours[best.soldier_id] += 1
                    if not t.allow_overlap:
                        busy_until[best.soldier_id] = max(
                            busy_until[best.soldier_id],
                            h + t.shift_duration + t.rest_duration - 1
                        )
    return hints


# ══════════════════════════════════════════════════════════════════
# 6. מנוע CP-SAT v8 — משולב ומתוקן (ללא פרצות אלגוריתמיות!)
# ══════════════════════════════════════════════════════════════════
def solve_scheduling(soldiers: list, tasks: list, num_hours: int = 24,
                     time_limit: float = 120.0, soft_rest: bool = True):
    model = cp_model.CpModel()
    x, start = {}, {}
    SLEEP_WINDOW = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8]
    zero_var = model.NewIntVar(0, 0, 'zero_const')

    # ── יצירת משתנים ──
    for s in soldiers:
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.soldier_id, t.task_id, slot_idx, h] = model.NewBoolVar(
                        f"x_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}")
                    start[s.soldier_id, t.task_id, slot_idx, h] = model.NewBoolVar(
                        f"st_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}")

    # ── Greedy hints ──
    hints = build_greedy_hints(soldiers, tasks, num_hours)
    for key, val in hints.items():
        if key in x:
            model.AddHint(x[key], val)

    # ── אילוצים לכל חייל ──
    rest_violation_vars = []  # לקנסות רכים

    for s in soldiers:

        # א. חסימת שעות קשיחה לחייל
        for h in s.unavail_hours:
            if h < num_hours:
                model.Add(
                    sum(x[s.soldier_id, t.task_id, slot, h]
                        for t in tasks for slot in range(len(t.slots))) == 0
                )

        for t in tasks:

            # ב. פטורים + חסימת תפקיד
            if any(role in t.blocked_roles for role in s.roles) \
                    or (t.task_id in s.restricted_tasks):
                for slot_idx in range(len(t.slots)):
                    for h in range(num_hours):
                        model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == 0)
                continue

            for slot_idx, required_role in enumerate(t.slots):

                # ג. התאמת תפקיד ל-slot
                if required_role is not None and required_role not in s.roles:
                    for h in range(num_hours):
                        model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == 0)
                    continue

                for h in range(num_hours):

                    # ד. נעילת משמרת מעגלית (קשיח — לא ניתן לוותר)
                    relevant_starts = [
                        start[s.soldier_id, t.task_id, slot_idx, (h - i) % num_hours]
                        for i in range(t.shift_duration)
                    ]
                    model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == sum(relevant_starts))

                    # ה. מנוחה — תקין ומדויק
                    if t.rest_duration > 0 and not t.allow_overlap:
                        rest_window = t.shift_duration + t.rest_duration
                        # המנוחה נספרת *רק* לאחר סיום המשמרת (החל מהשעה ה-shift_duration)
                        for rest_offset in range(t.shift_duration, min(rest_window, num_hours)):
                            rest_h = (h + rest_offset) % num_hours
                            for other_t in tasks:
                                if not other_t.allow_overlap:
                                    for other_slot in range(len(other_t.slots)):
                                        if soft_rest:
                                            # רך: אילוץ "אטום" שמכריח לשלם קנס על כל חפיפה
                                            viol = model.NewBoolVar(
                                                f'rv_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}_{rest_h}_{other_t.task_id}_{other_slot}')
                                            model.Add(viol >= start[s.soldier_id, t.task_id, slot_idx, h] + 
                                                      x[s.soldier_id, other_t.task_id, other_slot, rest_h] - 1)
                                            rest_violation_vars.append(viol)
                                        else:
                                            # קשיח
                                            model.Add(
                                                x[s.soldier_id, other_t.task_id, other_slot, rest_h] == 0
                                            ).OnlyEnforceIf(start[s.soldier_id, t.task_id, slot_idx, h])

    # ו. כיסוי עמדות — קשיח תמיד (מטרת-על)
    for t in tasks:
        for slot_idx in range(len(t.slots)):
            for h in range(num_hours):
                assigned = sum(x[s.soldier_id, t.task_id, slot_idx, h] for s in soldiers)
                if h in t.active_hours:
                    model.Add(assigned == 1)
                else:
                    model.Add(assigned == 0)

    # ז. חד-ערכיות
    for s in soldiers:
        for h in range(num_hours):
            blocking = [
                x[s.soldier_id, t.task_id, slot_idx, h]
                for t in tasks if not t.allow_overlap
                for slot_idx in range(len(t.slots))
            ]
            model.Add(sum(blocking) <= 1)

    # ── פונקציית מטרה ──
    s_total_hours, s_intensity_scores, sleep_penalties = [], [], []

    for s in soldiers:
        total_h = sum(
            x[s.soldier_id, t.task_id, slot, h]
            for t in tasks for slot in range(len(t.slots)) for h in range(num_hours)
        )
        s_total_hours.append(total_h)

        intensity_score = sum(
            x[s.soldier_id, t.task_id, slot, h] * t.intensity
            for t in tasks for slot in range(len(t.slots)) for h in range(num_hours)
        )
        s_intensity_scores.append(intensity_score)

        night_work = sum(
            x[s.soldier_id, t.task_id, slot, h]
            for t in tasks if not t.allow_overlap
            for slot in range(len(t.slots)) for h in SLEEP_WINDOW
        )
        night_work_var = model.NewIntVar(0, len(SLEEP_WINDOW), f'nw_{s.soldier_id}')
        model.Add(night_work_var == night_work)
        shifted = model.NewIntVar(-len(SLEEP_WINDOW), len(SLEEP_WINDOW), f'sh_{s.soldier_id}')
        model.Add(shifted == night_work_var - 4)
        penalty = model.NewIntVar(0, len(SLEEP_WINDOW), f'sp_{s.soldier_id}')
        model.AddMaxEquality(penalty, [zero_var, shifted])
        sleep_penalties.append(penalty)

    max_load = model.NewIntVar(0, 1000, 'max_load')
    min_load = model.NewIntVar(0, 1000, 'min_load')
    model.AddMaxEquality(max_load, s_total_hours)
    model.AddMinEquality(min_load, s_total_hours)
    load_diff = model.NewIntVar(0, 1000, 'load_diff')
    model.Add(load_diff == max_load - min_load)

    max_int = model.NewIntVar(0, 1000, 'max_int')
    min_int = model.NewIntVar(0, 1000, 'min_int')
    model.AddMaxEquality(max_int, s_intensity_scores)
    model.AddMinEquality(min_int, s_intensity_scores)
    int_diff = model.NewIntVar(0, 1000, 'int_diff')
    model.Add(int_diff == max_int - min_int)

    # קנס מנוחה רכה ×500 (מרכזי), שינה ×200, הוגנות ×100/×50
    total_rest_violations = sum(rest_violation_vars) if rest_violation_vars else model.NewIntVar(0,0,'no_rv')
    model.Minimize(
        500 * total_rest_violations
        + 100 * load_diff
        + 50  * int_diff
        + 200 * sum(sleep_penalties)
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers  = 8
    solver.parameters.log_search_progress = False
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, 0

    # ── בניית DataFrame תוצאות ──
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
    total_rest_viol_count = sum(solver.Value(v) for v in rest_violation_vars) if rest_violation_vars else 0

    for s in soldiers:
        row = {"שם": s.name}
        night_count = sum(
            solver.Value(x[s.soldier_id, t.task_id, slot, h])
            for t in tasks if not t.allow_overlap
            for slot in range(len(t.slots)) for h in SLEEP_WINDOW
        )
        for h in range(num_hours):
            active = [
                t.name for t in tasks
                for slot_idx in range(len(t.slots))
                if solver.Value(x[s.soldier_id, t.task_id, slot_idx, h]) == 1
            ]
            row[hour_labels[h]] = " + ".join(active) if active else "—"
        row["סך שעות"] = sum(
            1 for h in range(num_hours)
            if any(solver.Value(x[s.soldier_id, t.task_id, slot, h]) == 1
                   for t in tasks for slot in range(len(t.slots)))
        )
        row["מדד עצימות"] = sum(
            solver.Value(x[s.soldier_id, t.task_id, slot, h]) * t.intensity
            for t in tasks for slot in range(len(t.slots)) for h in range(num_hours)
        )
        row["שעות שינה (22-08)"] = len(SLEEP_WINDOW) - night_count
        rows.append(row)

    df = pd.DataFrame(rows)
    df.index = range(1, len(df) + 1)
    return df, total_rest_viol_count


# ══════════════════════════════════════════════════════════════════
# 7. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v8)</h1>
  <p>אילוצי מנוחה רכים מדויקים · Greedy hints · 8 ליבות מקבילות · הגנות מובנות</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs([
    "🚀  ביצוע שיבוץ",
    "📖  מדריך מפורט",
    "📥  תבניות אקסל",
])

# ══════════════════════════════════════════════════════════════════
# טאב: תבניות
# ══════════════════════════════════════════════════════════════════
with tab_templates:
    st.markdown("### 📥 הורדת תבניות עבודה")
    st.markdown(
        '<div class="info-box">הורידו את התבניות, מלאו לפי המדריך, '
        'ואז חזרו לטאב <b>ביצוע שיבוץ</b>.</div>',
        unsafe_allow_html=True,
    )

    s_ex = pd.DataFrame({
        'מספר אישי': [1001, 1002, 1003],
        'שם מלא':    ['ישראל ישראלי', 'יוסי כהן', 'אבי לוי'],
        'פטורים':    ['', '101', ''],
        'הסמכות':    ['נהג, מפקד', 'קצין', 'נהג'],
        'שעות חסימה': ['', '10-14', ''],
    })
    t_ex = pd.DataFrame({
        'מס"ד משימה':             [101, 102, 103],
        'שם המשימה':             ['סיור רכוב', 'שמירת ש.ג.', 'כוננות קבועה'],
        'סד"כ נדרש למשימה':     [4, 2, 8],
        'משך משמרת':             [4, 4, 24],
        'שעות מנוחה בין משימות': [8, 8, 0],
        'אישור חפיפה בין משימות': [False, False, True],
        'שעות פעילות':           ['all', 'all', 'all'],
        'הסמכה נדרשת':           ['נהג, מפקד', '', ''],
        'דירוג עצימות המשימה':   [3, 2, 1],
        'תפקידים חסומים':        ['', 'קצין', ''],
    })

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**👥 תבנית חיילים (5 עמודות)**")
        st.dataframe(s_ex, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ הורד תבנית חיילים",
            data=to_excel_styled(s_ex, "Soldiers", False),
            file_name="Soldiers_v8.xlsx", use_container_width=True,
        )
    with c2:
        st.markdown("**📋 תבנית משימות (10 עמודות)**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ הורד תבנית משימות",
            data=to_excel_styled(t_ex, "Tasks", False),
            file_name="Tasks_v8.xlsx", use_container_width=True,
        )

# ══════════════════════════════════════════════════════════════════
# טאב: מדריך
# ══════════════════════════════════════════════════════════════════
with tab_guide:
    st.markdown("### 📖 מדריך למילוי פורמט השבצ\"ק")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:20%">עמודה (חיילים)</th><th style="width:40%">הסבר</th><th>הנחיות</th></tr></thead>
      <tbody>
        <tr><td><b>מספר אישי</b></td><td>מזהה ייחודי</td><td>מספר שלם, ללא כפילויות</td></tr>
        <tr><td><b>שם מלא</b></td><td>שם להצגה</td><td>שם פרטי + משפחה</td></tr>
        <tr><td><b>פטורים</b></td><td>מס"ד משימות חסומות</td><td>למשל: <code>101,103</code>. ריק = אין פטורים</td></tr>
        <tr><td><b>הסמכות</b></td><td>תפקידים וכישורים</td><td>למשל: <code>נהג, מפקד</code>. זהה לכתיב בקובץ משימות</td></tr>
        <tr><td><b>שעות חסימה</b></td><td>שעות לא זמין</td><td>טווח <code>10-14</code> או רשימה <code>7,8</code> או טווח חוצה חצות <code>22-6</code></td></tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:20%">עמודה (משימות)</th><th style="width:40%">הסבר</th><th>הנחיות</th></tr></thead>
      <tbody>
        <tr><td><b>מס"ד משימה</b></td><td>מזהה משימה</td><td>תואם לעמודת פטורים בחיילים</td></tr>
        <tr><td><b>שם המשימה</b></td><td>שם קצר</td><td>למשל: "סיור רכוב"</td></tr>
        <tr><td><b>סד"כ נדרש</b></td><td>כמות חיילים</td><td>מספר שלם, חובה</td></tr>
        <tr><td><b>משך משמרת</b></td><td>שעות רצופות</td><td>מספר שלם. תומך במשמרות מעגליות</td></tr>
        <tr><td><b>שעות מנוחה</b></td><td>צינון אחרי משמרת</td><td>אילוץ רך — המערכת שומרת עליו בקפדנות, ותחרוג רק למניעת עמדה ריקה.</td></tr>
        <tr><td><b>אישור חפיפה</b></td><td>מקביל למשימות אחרות</td><td><code>True</code> = כוננות/קשר. <code>False</code> = שמירה/סיור</td></tr>
        <tr><td><b>שעות פעילות</b></td><td>מתי עמדה קיימת</td><td><code>all</code> / <code>8-12</code> / <code>7,8,9</code></td></tr>
        <tr><td><b>הסמכה נדרשת</b></td><td>תקנים ייעודיים</td><td>למשל: <code>נהג, מפקד</code> — שני תקנים שמורים</td></tr>
        <tr><td><b>דירוג עצימות</b></td><td>קושי המשימה 1-3</td><td>מחולק שוויונית בין החיילים</td></tr>
        <tr><td><b>תפקידים חסומים</b></td><td>מי לא יכול לבצע</td><td>למשל: <code>קצין</code></td></tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    <b>💡 v8 — שינויים עיקריים:</b><br>
    • <b>אילוצי מנוחה חכמים:</b> המערכת רושמת חריגות אמת. קנס של ×500 מתמטי על כל פגיעה במנוחה מבטיח שחריגה תבוצע <b>אך ורק בלית ברירה</b>.<br>
    • <b>היגיון זמן מתוקן:</b> מנוחה נספרת בצורה מדויקת <u>אך ורק</u> לאחר תום משך המשמרת (ולא תוך כדי).<br>
    • <b>Greedy hints:</b> הסולבר מקבל רמזים חכמים, מה שמאיץ מציאת פתרונות לכוחות גדולים.<br>
    • <b>8 ליבות מקבילות:</b> ניצול מקסימלי של כוח החישוב עד לזמן שנגדיר.
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# טאב: ביצוע שיבוץ
# ══════════════════════════════════════════════════════════════════
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    with st.expander("⚙️ הגדרות מתקדמות"):
        time_limit = st.slider("זמן מקסימלי לפתרון (שניות)", 30, 300, 120, 15)
        soft_rest  = st.toggle("אילוצי מנוחה רכים (מומלץ לכוחות גדולים)", value=True)
        st.markdown("""
        <div class="info-box" style="font-size:12px">
        <b>מנוחה רכה (מופעל):</b> המערכת תעשה כל מאמץ לשמור על מנוחה, אך במצב קריטי תשבץ בכל זאת ותדווח למפקד.<br>
        <b>מנוחה קשיחה (כבוי):</b> אסור להפר מנוחה בשום פנים. אם אין מספיק חיילים, המערכת תקרוס ותחזיר שגיאה.
        </div>
        """, unsafe_allow_html=True)

    if sf and tf:
        try:
            s_df = pd.read_excel(sf)
            t_df = pd.read_excel(tf)
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ שגיאה בקריאת הקבצים: {e}</div>',
                        unsafe_allow_html=True)
            st.stop()

        miss_s = {"מספר אישי", "שם מלא"} - set(s_df.columns)
        miss_t = {'מס"ד משימה', 'שם המשימה', 'סד"כ נדרש למשימה'} - set(t_df.columns)
        if miss_s:
            st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ חיילים: {miss_s}</div>',
                        unsafe_allow_html=True)
        if miss_t:
            st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ משימות: {miss_t}</div>',
                        unsafe_allow_html=True)
        if miss_s or miss_t:
            st.stop()

        if st.button('⚙️ צור שבצ"ק חכם (v8)', use_container_width=True):
            with st.spinner("בונה מודל ומשקלל אילוצים..."):
                soldiers = [
                    Soldier(r['מספר אישי'], r['שם מלא'],
                            r.get('פטורים', ''), r.get('הסמכות', ''),
                            r.get('שעות חסימה', ''))
                    for _, r in s_df.iterrows()
                ]
                tasks = [
                    Task(r['מס"ד משימה'], r['שם המשימה'], r['סד"כ נדרש למשימה'],
                         r.get('משך משמרת'), r.get('שעות מנוחה בין משימות'),
                         r.get('אישור חפיפה בין משימות'), r.get('שעות פעילות'),
                         r.get('הסמכה נדרשת', ''), r.get('דירוג עצימות המשימה', 1),
                         r.get('תפקידים חסומים', ''))
                    for _, r in t_df.iterrows()
                ]

                needed_roles    = {role for t in tasks for role in t.slots if role is not None}
                available_roles = {role for s in soldiers for role in s.roles}
                missing_roles   = needed_roles - available_roles

            if missing_roles:
                st.markdown(
                    f'<div class="error-box">❌ חסרים חיילים עם התפקידים: '
                    f'<b>{", ".join(missing_roles)}</b>.</div>',
                    unsafe_allow_html=True,
                )
            else:
                with st.spinner(f"מריץ אופטימיזציה מבוזרת ({time_limit}ש׳ מקסימום, 8 ליבות)..."):
                    final_df, rest_viols = solve_scheduling(
                        soldiers, tasks,
                        time_limit=time_limit,
                        soft_rest=soft_rest
                    )

                if final_df is not None:
                    gap_h     = int(final_df["סך שעות"].max() - final_df["סך שעות"].min())
                    avg_h     = final_df["סך שעות"].mean()
                    avg_sleep = final_df["שעות שינה (22-08)"].mean()
                    badge     = "✅ מצוין" if gap_h <= 2 else ("⚠️ סביר" if gap_h <= 5 else "❗ גבוה")

                    # ── אזהרת הפרות מנוחה ──
                    if rest_viols > 0:
                        st.markdown(
                            f'<div class="warn-box">⚠️ <b>{rest_viols} הפרות מנוחה אמיתיות</b> נמצאו בפתרון — '
                            f'המערכת שיבצה למרות אילוץ המנוחה כדי לכסות את כל העמדות (קנס מתמטי כבד הופעל). '
                            f'מומלץ לבדוק את השיבוץ ולאזן ידנית אם נדרש.</div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown(
                            '<div class="info-box">✅ כל אילוצי המנוחה נשמרו באופן מלא! לא בוצעה שום חריגה.</div>',
                            unsafe_allow_html=True,
                        )

                    st.markdown(f"""
                    <div class="metric-row">
                      <div class="metric-card">
                        <div class="mc-label">חיילים בשיבוץ</div>
                        <div class="mc-value">{len(soldiers)}</div>
                        <div class="mc-sub">{len(tasks)} סוגי משימות</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">ממוצע שעות</div>
                        <div class="mc-value">{avg_h:.1f}</div>
                        <div class="mc-sub">לחייל</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">פער הוגנות</div>
                        <div class="mc-value">{gap_h}</div>
                        <div class="mc-sub">{badge}</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">ממוצע שינה</div>
                        <div class="mc-value">{avg_sleep:.1f}</div>
                        <div class="mc-sub">יעד: 7.0 שעות</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">הפרות מנוחה</div>
                        <div class="mc-value">{"0 ✅" if rest_viols == 0 else f"{rest_viols} ⚠️"}</div>
                        <div class="mc-sub">{"ללא הפרות" if rest_viols == 0 else "בדוק לוז"}</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown("---")
                    st.subheader("📅 לוח השיבוץ הסופי")
                    st.table(final_df)
                    st.download_button(
                        "📥 הורד לוח שיבוץ (Excel)",
                        data=to_excel_styled(final_df),
                        file_name="Final_Shavtzak_v8.xlsx",
                        use_container_width=True,
                    )

                    st.markdown("---")
                    st.subheader("📊 ניתוח עומסים")
                    fig = px.bar(
                        final_df, x="שם", y="סך שעות", color="מדד עצימות",
                        color_continuous_scale=["#a8d5a2", "#1a3d17"],
                        title="עומס שעות ועצימות לחייל",
                        text="סך שעות",
                    )
                    fig.update_traces(textposition="outside", marker_line_width=0)
                    fig.update_layout(
                        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                        font=dict(family="Heebo", size=12),
                        title_font=dict(size=15, color="#1a3d17"),
                        xaxis=dict(tickangle=-30, gridcolor="#f0f0f0", title=""),
                        yaxis=dict(gridcolor="#f0f0f0", title='שעות סה"כ'),
                        margin=dict(t=55, b=80, l=30, r=20),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    with st.expander("💡 תובנות אופטימיזציה לגרסה v8"):
                        max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                        st.markdown(f"""
**חיילים עמוסים:** {', '.join(max_s)} — {final_df['סך שעות'].max()} שעות.

**פער הוגנות:** {gap_h} שעות — {'חלוקה אחידה.' if gap_h <= 2 else 'מומלץ להוסיף חיילים לאיזון.' if gap_h > 4 else 'חלוקה סבירה.'}

**שינה:** ממוצע {avg_sleep:.1f} שעות. הופעל קנס ×200 מתמטי על כל שעת עבודת לילה מעבר ל-4 שעות (כדי להגן על חלון שינה של 7 שעות).

**מנוחה חכמה:** {'המערכת הצליחה לשבץ ללא שום חריגת מנוחה! עבודה יפה.' if rest_viols == 0 else f'{rest_viols} הפרות מנוחה — הסולבר נלחם כדי להימנע מכך אך נאלץ להקריב מנוחות אלו במקום להשאיר את העמדה ריקה. הקנס שעליו שילם המודל: {rest_viols * 500} נקודות.'}
                        """)
                else:
                    st.markdown("""
                    <div class="error-box">
                    ❌ לא נמצא פתרון גם עם אילוצים רכים.<br><br>
                    <b>סיבות אפשריות:</b><br>
                    • יש משימות עם הסמכות נדרשות שאין מספיק חיילים מוסמכים לאיישן.<br>
                    • שעות חסימה של חיילים חוסמות שעות קריטיות שאין מי שיאייש.<br>
                    • הסד"כ הנדרש גדול ממספר החיילים הזמינים.<br>
                    • נסה להגדיל את <b>זמן הפתרון</b> בהגדרות המתקדמות ולחץ שוב.
                    </div>
                    """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין לכם תבניות? עברו לטאב <b>תבניות אקסל</b> והורידו.
        </div>
        """, unsafe_allow_html=True)
