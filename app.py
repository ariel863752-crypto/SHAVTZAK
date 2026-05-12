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

/* ── כותרת ── */
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

/* ── טאבים ── */
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

/* ── כרטיסי מדד ── */
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

/* ── כפתורים ── */
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

/* ── העלאת קבצים ── */
[data-testid="stFileUploader"] {
    background: white; border-radius: 12px; padding: 14px 16px;
    border: 2px dashed #c0d8bc; direction: rtl; text-align: right;
}
[data-testid="stFileUploader"]:hover { border-color: #2d5a27; }
[data-testid="stFileUploadDropzone"] input { color: transparent !important; }

/* ── טבלאות ── */
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

/* ── טבלת מדריך ── */
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

/* ── תיבות ── */
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

/* ── expander ── */
[data-testid="stExpander"] summary {
    direction: rtl; text-align: right;
    font-family: 'Heebo', sans-serif; font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 3. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
def parse_time_ranges(val) -> list[int]:
    """
    מפרסר שעות פעילות / חסימה.
    תומך ב: 'all', מספר בודד '8', טווח '8-12', טווח חוצה חצות '22-2', רשימה '7,8,9'.
    """
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
        # slots: תפקידים מיוחדים + מילוי ב-None עד req_p
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
# 5. מנוע CP-SAT (v6.4 - High Speed Production)
#
# אילוצים (קשיחים):
#   א. חסימת שעות לחייל (unavail_hours)
#   ב. פטורים + תפקידים חסומים — חסימה ברמת חייל×משימה
#   ג. התאמת slot לתפקיד — slot מסוים ≡ תפקיד מסוים
#   ד. נעילת משמרת מעגלית — x[h] = sum(start[h-i] for i in range(shift))
#   ה. מנוחה חובה — OnlyEnforceIf(start[h]): חסום כל עמדה חוסמת בחלון מנוחה
#   ו. כיסוי עמדות — בדיוק 1 חייל לכל slot×שעה פעילה (מטרת-על)
#   ז. חד-ערכיות — חייל בעמדה חוסמת אחת בכל שעה
#
# פונקציית מטרה (רכה):
#   min 100×(max_load - min_load)   ← הוגנות שעות
#      + 50×(max_int - min_int)     ← הוגנות עצימות
#      + 200×Σ sleep_penalty        ← שמירת שינה
# ══════════════════════════════════════════════════════════════════
def solve_scheduling(soldiers: list, tasks: list, num_hours: int = 24):
    model = cp_model.CpModel()
    x, start = {}, {}
    # חלון שינה: 22:00–08:59 = 11 שעות
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

    # ── אילוצים לכל חייל ──
    for s in soldiers:

        # א. חסימת שעות קשיחה לחייל
        for h in s.unavail_hours:
            if h < num_hours:
                model.Add(
                    sum(x[s.soldier_id, t.task_id, slot, h]
                        for t in tasks for slot in range(len(t.slots))) == 0
                )

        for t in tasks:

            # ב. פטורים + חסימת תפקיד — חסום כל השעות לכל ה-slots
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

                    # ד. נעילת משמרת מעגלית
                    relevant_starts = [
                        start[s.soldier_id, t.task_id, slot_idx, (h - i) % num_hours]
                        for i in range(t.shift_duration)
                    ]
                    model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == sum(relevant_starts))

                    # ה. מנוחה חובה — OnlyEnforceIf(start[h])
                    # רק ממשימות חוסמות, רק לשעות מנוחה אחרי סיום המשמרת
                    if t.rest_duration > 0 and not t.allow_overlap:
                        rest_window = t.shift_duration + t.rest_duration
                        for rest_offset in range(1, min(rest_window, num_hours)):
                            rest_h = (h + rest_offset) % num_hours
                            for other_t in tasks:
                                if not other_t.allow_overlap:
                                    for other_slot in range(len(other_t.slots)):
                                        model.Add(
                                            x[s.soldier_id, other_t.task_id, other_slot, rest_h] == 0
                                        ).OnlyEnforceIf(start[s.soldier_id, t.task_id, slot_idx, h])

    # ו. כיסוי עמדות — מטרת-על
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

        # קנס שינה: עובד >4 שעות בחלון הלילה → קנס על כל שעה מעבר
        night_work = sum(
            x[s.soldier_id, t.task_id, slot, h]
            for t in tasks if not t.allow_overlap
            for slot in range(len(t.slots)) for h in SLEEP_WINDOW
        )
        night_work_var = model.NewIntVar(0, len(SLEEP_WINDOW), f'nw_{s.soldier_id}')
        model.Add(night_work_var == night_work)
        shifted = model.NewIntVar(-len(SLEEP_WINDOW), len(SLEEP_WINDOW), f'sh_{s.soldier_id}')
        model.Add(shifted == night_work_var - 4)   # יעד: 11-4=7 שעות שינה
        penalty = model.NewIntVar(0, len(SLEEP_WINDOW), f'sp_{s.soldier_id}')
        model.AddMaxEquality(penalty, [zero_var, shifted])
        sleep_penalties.append(penalty)

    # משתני הוגנות
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

    model.Minimize(100 * load_diff + 50 * int_diff + 200 * sum(sleep_penalties))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 45.0
    solver.parameters.num_search_workers  = 4
    solver.parameters.log_search_progress = False
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # ── בניית DataFrame תוצאות ──
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
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
    return df


# ══════════════════════════════════════════════════════════════════
# 6. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v6.4)</h1>
  <p>אופטימיזציה מקבילית עם ניהול תפקידים, עצימות, שעות חסימה ושמירת שינה.</p>
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
        'מס"ד משימה':            [101, 102, 103],
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
            file_name="Soldiers_v6.4.xlsx", use_container_width=True,
        )
    with c2:
        st.markdown("**📋 תבנית משימות (10 עמודות)**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ הורד תבנית משימות",
            data=to_excel_styled(t_ex, "Tasks", False),
            file_name="Tasks_v6.4.xlsx", use_container_width=True,
        )

# ══════════════════════════════════════════════════════════════════
# טאב: מדריך
# ══════════════════════════════════════════════════════════════════
with tab_guide:
    st.markdown("### 📖 מדריך למילוי פורמט השבצ\"ק")
    st.markdown("כדי שהאלגוריתם יוכל לשבץ את הכוחות בצורה מושלמת, חשוב להזין את הנתונים בתבניות האקסל בדיוק לפי הכללים הבאים:")

    st.markdown("#### 👥 קובץ חיילים — `Soldiers.xlsx`")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:20%">עמודה</th><th style="width:40%">הסבר</th><th>הנחיות ומקרי קצה</th></tr></thead>
      <tbody>
        <tr><td><b>מספר אישי</b></td><td>מזהה ייחודי לכל חייל במערכת.</td><td>חובה להזין מספר שלם, ללא כפילויות בין חיילים.</td></tr>
        <tr><td><b>שם מלא</b></td><td>שם החייל כפי שתרצו שיוצג.</td><td>יופיע בלוח השיבוץ הסופי ובגרפים. מומלץ שם פרטי ומשפחה.</td></tr>
        <tr><td><b>פטורים</b></td><td>משימות ספציפיות שהחייל מנוע מלבצע.</td><td>יש להזין את <b>מס"ד המשימה</b>. למשל: <code>101,103</code>. במידה ואין פטור, השאירו ריק.</td></tr>
        <tr><td><b>הסמכות</b></td><td>התפקידים והכישורים של החייל.</td><td>הפרידו בפסיקים. למשל: <code>נהג, מפקד, חובש</code>. חשוב שהכתיב יהיה זהה לחלוטין למה שייכתב בקובץ המשימות.</td></tr>
        <tr><td><b>שעות חסימה</b></td><td>שעות (0-23) בהן החייל אינו זמין כלל לשיבוץ (למשל: הפניה, יציאה).</td><td>ניתן להזין טווחים (למשל <code>10-14</code>) או שעות בודדות מופרדות בפסיק (<code>7,8,18</code>). המערכת תומכת בטווחים שחוצים חצות (למשל <code>22-6</code>).</td></tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 📋 קובץ משימות — `Tasks.xlsx`")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:20%">עמודה</th><th style="width:40%">הסבר</th><th>הנחיות ומקרי קצה</th></tr></thead>
      <tbody>
        <tr><td><b>מס"ד משימה</b></td><td>מספר סידורי וזיהוי המשימה.</td><td>חייב להיות תואם למה שנכתב בעמודת ה"פטורים" אצל החיילים.</td></tr>
        <tr><td><b>שם המשימה</b></td><td>השם שיוצג בלוח השיבוץ.</td><td>הקפידו על שמות קצרים וברורים (למשל: "סיור רכוב", "תורנות מטבח").</td></tr>
        <tr><td><b>סד"כ נדרש למשימה</b></td><td>הכמות הכוללת של החיילים שצריכים לאייש את העמדה.</td><td>מספר שלם. האלגוריתם לא יאפשר למשימה לפעול בחסר.</td></tr>
        <tr><td><b>משך משמרת</b></td><td>כמות השעות הרצופות של המשמרת.</td><td>מספר שלם (למשל: <code>4</code>). המערכת תומכת אוטומטית במשמרות שמתחילות בלילה ונגמרות בבוקר (מעגליות).</td></tr>
        <tr><td><b>שעות מנוחה בין משימות</b></td><td>זמן הצינון (מנוחה) שחייל חייב לקבל מיד בתום המשמרת.</td><td>בזמן זה החייל יהיה חסום לשיבוץ במשימות חוסמות אחרות. <b>שימו לב:</b> חלון החסימה הכולל הוא המשמרת + המנוחה.</td></tr>
        <tr><td><b>אישור חפיפה בין משימות</b></td><td>האם ניתן לבצע את המשימה במקביל למשימה אחרת?</td><td><code>True</code> (למשל: כוננות, קשר) או <code>False</code> (למשל: שמירה בעמדה). משימה ללא חפיפה חוסמת שיבוצים מקבילים.</td></tr>
        <tr><td><b>שעות פעילות</b></td><td>באיזה שעות ביום העמדה הזו בכלל קיימת?</td><td>הזינו <code>all</code> לפעילות 24/7, או טווחים (למשל <code>8-12</code>), או רשימה פרטנית (<code>7,8,9</code>).</td></tr>
        <tr><td><b>הסמכה נדרשת</b></td><td>הגדרת "תקנים" קשיחים מתוך הסד"כ הנדרש.</td><td>אם משימה דורשת 4 אנשים וכתבתם <code>נהג, מפקד</code> — התקן הראשון ישוריין אך ורק לנהג, התקן השני למפקד, ושני החיילים הנותרים ייבחרו מכלל הסד"כ הפנוי.</td></tr>
        <tr><td><b>דירוג עצימות המשימה</b></td><td>עד כמה המשימה שוחקת/קשה (דירוג 1 עד 3).</td><td>האלגוריתם שואף לפזר משימות קשות בצורה שוויונית בין כולם למניעת שחיקה.</td></tr>
        <tr><td><b>תפקידים חסומים</b></td><td>מי <b>אסור</b> לו בשום אופן לבצע את המשימה.</td><td>למשל משימת רס"ר: <code>קצין, נגד</code>. חייל שמחזיק בהסמכה כזו חסום הרמטית.</td></tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    <b>💡 הערות חשובות:</b><br>
    • <b>מטרת-על:</b> כיסוי כל העמדות — אילוץ קשיח לעולם לא מופר.<br>
    • <b>מעגליות:</b> משמרת שמתחילה ב-22:00 ונגמרת ב-02:00 עובדת תקין.<br>
    • <b>ריבוי ליבות:</b> הסולבר רץ על 4 ליבות במקביל — עד 45 שניות.<br>
    • <b>שינה:</b> עונש ×200 על כל שעת לילה מעל 4. יעד: 7 שעות שינה רצופות.
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

    if sf and tf:
        try:
            s_df = pd.read_excel(sf)
            t_df = pd.read_excel(tf)
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ שגיאה בקריאת הקבצים: {e}</div>',
                        unsafe_allow_html=True)
            st.stop()

        # ── ולידציה: עמודות חובה ──
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

        if st.button('⚙️ צור שבצ"ק חכם (v6.4)', use_container_width=True):
            with st.spinner("בונה מודל..."):
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

                # ולידציה: תפקידים חסרים
                needed_roles    = {role for t in tasks for role in t.slots if role is not None}
                available_roles = {role for s in soldiers for role in s.roles}
                missing_roles   = needed_roles - available_roles

            if missing_roles:
                st.markdown(
                    f'<div class="error-box">❌ חסרים חיילים עם התפקידים: '
                    f'<b>{", ".join(missing_roles)}</b>.<br>'
                    f'לא ניתן לאייש עמדות הדורשות תפקידים אלה.</div>',
                    unsafe_allow_html=True,
                )
            else:
                with st.spinner("מריץ אופטימיזציה מקבילית (עד 45 שניות)..."):
                    final_df = solve_scheduling(soldiers, tasks)

                if final_df is not None:
                    gap_h = int(final_df["סך שעות"].max() - final_df["סך שעות"].min())
                    avg_h = final_df["סך שעות"].mean()
                    avg_sleep = final_df["שעות שינה (22-08)"].mean()
                    badge = "✅ מצוין" if gap_h <= 2 else ("⚠️ סביר" if gap_h <= 5 else "❗ גבוה")

                    # ── כרטיסי מדד ──
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
                    </div>
                    """, unsafe_allow_html=True)

                    # ── לוח שיבוץ ──
                    st.markdown("---")
                    st.subheader("📅 לוח השיבוץ הסופי")
                    st.table(final_df)
                    st.download_button(
                        "📥 הורד לוח שיבוץ (Excel)",
                        data=to_excel_styled(final_df),
                        file_name="Final_Shavtzak_v6.4.xlsx",
                        use_container_width=True,
                    )

                    # ── גרף עומסים ──
                    st.markdown("---")
                    st.subheader("📊 ניתוח עומסים")
                    fig = px.bar(
                        final_df, x="שם", y="סך שעות", color="מדד עצימות",
                        color_continuous_scale=["#a8d5a2", "#1a3d17"],
                        title="עומס שעות ועצימות לחייל (ירוק כהה = משימות קשות יותר)",
                        text="סך שעות",
                    )
                    fig.update_traces(textposition="outside", marker_line_width=0)
                    fig.update_layout(
                        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                        font=dict(family="Heebo", size=12),
                        title_font=dict(size=15, color="#1a3d17"),
                        coloraxis_showscale=True,
                        xaxis=dict(tickangle=-30, gridcolor="#f0f0f0", title=""),
                        yaxis=dict(gridcolor="#f0f0f0", title='שעות סה"כ'),
                        margin=dict(t=55, b=80, l=30, r=20),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # ── ניתוח ──
                    with st.expander("💡 למה הלוּז נראה ככה? — תובנות אופטימיזציה"):
                        max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                        st.markdown(f"""
**אילוצי תפקידים:** כל slot עם תפקיד מיוחד (נהג, מפקד) שוריין אוטומטית — רק בעל הכישור המתאים שובץ.

**תפקידים חסומים:** חיילים עם תפקידים שהוגדרו כחסומים למשימה לא הוכנסו כלל לאותה עמדה.

**חיילים עמוסים:** {', '.join(max_s)} — {final_df['סך שעות'].max()} שעות. ייתכן שהם הכישורים הנדרשים ביותר.

**פער הוגנות:** {gap_h} שעות — {'חלוקה אחידה.' if gap_h <= 2 else 'מומלץ להוסיף חיילים עם תפקידים זהים לאיזון.' if gap_h > 4 else 'חלוקה סבירה.'}

**שינה:** ממוצע {avg_sleep:.1f} שעות. האלגוריתם מעניש ×200 כל שעת לילה מעל 4 כדי להגן על חלון השינה.
                        """)
                else:
                    st.markdown("""
                    <div class="error-box">
                    ❌ לא נמצא פתרון חוקי העומד בכל האילוצים הקשיחים.<br><br>
                    <b>המלצות:</b><br>
                    • בדוק <b>תפקידים חסומים</b> — ייתכן שחסמת יותר מדי ולא נשאר כוח אדם.<br>
                    • בדוק <b>שעות חסימה</b> לחיילים — ייתכן שיוצרות חוסר בשעות קריטיות.<br>
                    • הפחת <b>שעות מנוחה בין משימות</b>.<br>
                    • הוסף חיילים או הפחת דרישות כוח אדם.
                    </div>
                    """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין לכם תבניות? עברו לטאב <b>תבניות אקסל</b> והורידו.
        </div>
        """, unsafe_allow_html=True)
