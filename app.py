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
# 2. CSS — RTL מלא + עיצוב
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* ── RTL שורש ── */
html, body { direction: rtl; }

.stApp,
[data-testid="stAppViewContainer"],
.main, .block-container,
.stMarkdown, p, span, li,
label, div,
[data-testid="stText"],
[data-testid="stMarkdownContainer"],
[data-testid="stAlert"],
[data-testid="stExpander"] summary,
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] div,
[data-testid="stSlider"] label,
[data-testid="stSelectbox"] label {
    font-family: 'Heebo', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
}

/* ── רקע ── */
.stApp { background: #f2f5f2; }
.block-container { padding: 2rem 2.5rem 3rem; max-width: 1400px; }

/* ── כותרת ── */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 30px 35px;
    margin-bottom: 28px;
    box-shadow: 0 4px 20px rgba(45,90,39,0.25);
    text-align: right;
}
.app-header h1 {
    font-size: clamp(24px, 4vw, 42px);
    font-weight: 900;
    color: white;
    margin: 0 0 8px 0;
    letter-spacing: -0.5px;
}
.app-header p {
    font-size: 16px;
    color: rgba(255,255,255,0.88);
    margin: 0;
}

/* ── טאבים — מימין לשמאל ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important;
    justify-content: flex-start !important;
    gap: 6px;
    background: white;
    border-radius: 12px;
    padding: 5px;
    border: 1px solid #dde8dc;
    margin-bottom: 20px;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px;
    font-family: 'Heebo', sans-serif;
    font-weight: 600;
    font-size: 14px;
    padding: 8px 20px;
    color: #5a7a57;
    direction: rtl;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: #2d5a27 !important;
    color: white !important;
}

/* ── כרטיסי מדד ── */
.metric-row {
    display: flex;
    flex-direction: row-reverse;
    gap: 16px;
    margin: 22px 0;
    flex-wrap: wrap;
}
.metric-card {
    flex: 1;
    min-width: 160px;
    background: white;
    border-radius: 14px;
    padding: 22px;
    border: 1px solid #dde8dc;
    box-shadow: 0 2px 8px rgba(45,90,39,0.07);
    text-align: right;
    direction: rtl;
}
.mc-label {
    font-size: 11px;
    color: #7a9a77;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 6px;
}
.mc-value {
    font-size: 34px;
    font-weight: 900;
    color: #1a3d17;
    line-height: 1;
}
.mc-sub {
    font-size: 12px;
    color: #a0b89d;
    margin-top: 4px;
}

/* ── כפתורים ── */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important;
    font-family: 'Heebo', sans-serif !important;
    font-weight: 700 !important;
    font-size: 17px !important;
    border-radius: 10px !important;
    border: none !important;
    height: 3.4em;
    width: 100%;
    box-shadow: 0 4px 14px rgba(45,90,39,0.3);
    transition: all 0.18s !important;
}
div.stButton > button:first-child:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(45,90,39,0.4) !important;
}
[data-testid="stDownloadButton"] > button {
    background: #b84d00 !important;
    color: white !important;
    font-family: 'Heebo', sans-serif !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    border: none !important;
}

/* ── העלאת קבצים ── */
[data-testid="stFileUploader"] {
    background: white;
    border-radius: 12px;
    padding: 14px 16px;
    border: 2px dashed #c0d8bc;
    direction: rtl;
    text-align: right;
}
[data-testid="stFileUploader"]:hover { border-color: #2d5a27; }

/* ── טבלאות ── */
[data-testid="stTable"] table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12.5px;
    background: white;
    direction: rtl;
}
[data-testid="stTable"] th {
    background: #2d5a27 !important;
    color: white !important;
    padding: 10px 13px !important;
    font-family: 'Heebo', sans-serif !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    text-align: right !important;
    direction: rtl !important;
}
[data-testid="stTable"] td {
    padding: 9px 13px !important;
    border-bottom: 1px solid #f0f0f0 !important;
    text-align: right !important;
    direction: rtl !important;
    font-family: 'Heebo', sans-serif !important;
}
[data-testid="stTable"] tr:nth-child(even) td { background: #f8fcf7 !important; }
[data-testid="stTable"] tr:hover td { background: #f0f8ef !important; }

/* ── טבלת מדריך ── */
.guide-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    margin: 16px 0;
    font-size: 14px;
    font-family: 'Heebo', sans-serif;
    direction: rtl;
}
.guide-table th {
    background: #2d5a27;
    color: white;
    padding: 13px 16px;
    text-align: right;
    font-weight: 700;
    font-size: 13px;
}
.guide-table td {
    padding: 12px 16px;
    border-bottom: 1px solid #eef3ed;
    background: white;
    line-height: 1.7;
    vertical-align: top;
    text-align: right;
}
.guide-table tr:last-child td { border-bottom: none; }
.guide-table tr:hover td { background: #f5fbf4; }

/* ── תיבות מידע ── */
.info-box {
    background: #edf5ec;
    border-right: 5px solid #2d5a27;
    border-left: none;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #1a3d17;
    line-height: 1.8;
    direction: rtl;
    text-align: right;
}
.warn-box {
    background: #fff8e6;
    border-right: 5px solid #e67e22;
    border-left: none;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #7a4500;
    line-height: 1.8;
    direction: rtl;
    text-align: right;
}
.error-box {
    background: #fdecea;
    border-right: 5px solid #c0392b;
    border-left: none;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #7a0010;
    line-height: 1.8;
    direction: rtl;
    text-align: right;
}

/* ── expander ── */
[data-testid="stExpander"] summary {
    direction: rtl;
    text-align: right;
    font-family: 'Heebo', sans-serif;
    font-weight: 600;
    font-size: 15px;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 3. מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
class Soldier:
    def __init__(self, s_id, name, restr="", roles=""):
        self.soldier_id = str(s_id)
        self.name = str(name).strip()

        # פטורים — קודי משימות שהחייל אינו יכול לבצע
        if pd.notna(restr) and str(restr).strip() not in ("", "nan"):
            self.restricted_tasks = [
                int(float(t)) for t in str(restr).split(',')
                if str(t).strip().replace('.0', '').lstrip('-').isdigit()
            ]
        else:
            self.restricted_tasks = []

        # תפקידים / כישורים (נהג, חובש, מפקד…)
        if pd.notna(roles) and str(roles).strip() not in ("", "nan"):
            self.roles = [r.strip() for r in str(roles).split(',') if r.strip()]
        else:
            self.roles = []


class Task:
    NIGHT: set = {22, 23, 0, 1, 2, 3, 4, 5, 6, 7}

    def __init__(self, t_id, name, req_p, shift_dur, rest_dur=0,
                 overlap=False, hours="all", req_role=""):
        self.task_id            = int(t_id)
        self.name               = str(name).strip()
        self.required_personnel = int(req_p)
        self.shift_duration     = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration      = int(rest_dur)  if pd.notna(rest_dur)  else 0
        self.allow_overlap      = str(overlap).strip().lower() == 'true'

        # תפקיד חובה למשימה — None אם כל חייל מתאים
        raw_role = str(req_role).strip() if pd.notna(req_role) else ""
        self.required_role = raw_role if raw_role not in ("", "nan") else None

        # שעות פעילות
        if pd.isna(hours) or str(hours).strip().lower() in ('all', '', 'nan'):
            self.active_hours = list(range(24))
        else:
            self.active_hours = [
                int(x.strip()) for x in str(hours).split(',')
                if str(x).strip().isdigit()
            ]


# ══════════════════════════════════════════════════════════════════
# 4. Excel מעוצב
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df: pd.DataFrame, sheet_name: str = 'שבצ"ק', include_index: bool = True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        wb  = writer.book
        ws  = writer.sheets[sheet_name]
        hfmt = wb.add_format({
            'bold': True, 'fg_color': '#2d5a27',
            'font_color': 'white', 'border': 1,
            'align': 'right', 'valign': 'vcenter',
        })
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
# 5. מנוע CP-SAT
#
# תיקונים לוגיים לעומת הגרסה המקורית:
#   א. אילוץ תפקידים הועבר לפני לולאת h (מניעת כפילות אילוצים)
#   ב. טווח max_load תוקן ל-num_hours*len(tasks) לתמיכה בחפיפה
#   ג. אילוץ חציית חצות (start=0 אם h+shift>num_hours) נוסף לפני
#      חישוב relevant_starts כדי למנוע אינדקסים שליליים
# ══════════════════════════════════════════════════════════════════
def solve_scheduling(soldiers: list, tasks: list, num_hours: int = 24):
    model = cp_model.CpModel()
    x, start = {}, {}

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h]     = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"st_{s.soldier_id}_{t.task_id}_{h}")

    for s in soldiers:
        for t in tasks:
            # ── אילוץ א: תפקיד נדרש ──
            # אם למשימה דרוש תפקיד שאין לחייל — חסום לחלוטין
            # (חייב להיות *לפני* לולאת h כדי לא לשכפל)
            if t.required_role and t.required_role not in s.roles:
                for h in range(num_hours):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)
                continue          # אין טעם להוסיף אילוצים נוספים לחייל זה

            for h in range(num_hours):
                # ── אילוץ ב: חסימת חציית חצות ──
                if h + t.shift_duration > num_hours:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

                # ── אילוץ ג: נעילת משמרות (x = סכום starts שמכסים h) ──
                relevant_starts = [
                    start[s.soldier_id, t.task_id, i]
                    for i in range(max(0, h - t.shift_duration + 1), h + 1)
                ]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))

                # ── אילוץ ד: מנוחה חובה אחרי משמרת ──
                if t.rest_duration > 0:
                    total_busy = t.shift_duration + t.rest_duration
                    for next_h in range(h + 1, min(h + total_busy, num_hours)):
                        for other_t in tasks:
                            if not other_t.allow_overlap:
                                model.AddImplication(
                                    start[s.soldier_id, t.task_id, h],
                                    x[s.soldier_id, other_t.task_id, next_h].Not(),
                                )

    # ── אילוץ ה: כיסוי עמדות (מטרת-על) ──
    for h in range(num_hours):
        for t in tasks:
            assigned = [x[s.soldier_id, t.task_id, h] for s in soldiers]
            if h in t.active_hours:
                model.Add(sum(assigned) == t.required_personnel)
            else:
                model.Add(sum(assigned) == 0)

    # ── אילוץ ו: חד-ערכיות (חייל בעמדה אחת בכל שעה) ──
    for s in soldiers:
        for h in range(num_hours):
            blocking = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking) <= 1)

    # ── אילוץ ז: פטורים ──
    for s in soldiers:
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # ── פונקציית מטרה ──
    # תוקן: טווח max_load = num_hours * len(tasks) לתמיכה בחיילים עם חפיפה
    upper_bound = num_hours * len(tasks)
    s_loads = [
        sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours))
        for s in soldiers
    ]
    max_load = model.NewIntVar(0, upper_bound, 'max_load')
    min_load = model.NewIntVar(0, upper_bound, 'min_load')
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)
    load_range = model.NewIntVar(0, upper_bound, 'load_range')
    model.Add(load_range == max_load - min_load)

    night_work = sum(
        x[s.soldier_id, t.task_id, h]
        for s in soldiers
        for t in tasks
        if not t.allow_overlap
        for h in Task.NIGHT
        if h < num_hours
    )

    # ממזערים פער עומסים (הוגנות) + עבודת לילה
    model.Minimize(50 * load_range + night_work)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    solver.parameters.num_search_workers  = 4
    solver.parameters.log_search_progress = False
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # ── בניית תוצאות ──
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
    for s in soldiers:
        row = {"שם": s.name}
        night_count = 0
        for h in range(num_hours):
            active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if h in Task.NIGHT and any(
                solver.Value(x[s.soldier_id, t.task_id, h]) == 1
                for t in tasks if not t.allow_overlap
            ):
                night_count += 1
        row["סך שעות"]   = sum(
            1 for h in range(num_hours)
            if any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks)
        )
        row["שעות לילה"] = night_count
        rows.append(row)

    df = pd.DataFrame(rows)
    df.index = range(1, len(df) + 1)
    return df


# ══════════════════════════════════════════════════════════════════
# 6. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות ותורנויות — כולל ניהול תפקידים וכישורים</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs([
    "🚀  ביצוע שיבוץ",
    "📖  מדריך למילוי שבצ\"ק",
    "📥  תבניות אקסל",
])

# ────────────────── תבניות ──────────────────
with tab_templates:
    st.markdown("### 📥 הורדת תבניות עבודה")
    st.markdown(
        '<div class="info-box">הורידו את התבניות, מלאו לפי המדריך, ואז חזרו לטאב <b>ביצוע שיבוץ</b>.</div>',
        unsafe_allow_html=True,
    )

    s_ex = pd.DataFrame({
        'מספר_אישי': [1001, 1002],
        'שם':        ['ישראל ישראלי', 'יוסי כהן'],
        'פטורים':    ['', '1'],
        'תפקידים':   ['נהג, מפקד', 'קלע'],
    })
    t_ex = pd.DataFrame({
        'קוד_משימה':        [1,             2,          10],
        'שם':               ['שמירת שער',   'סיור רכוב','כוננות'],
        'כוח_אדם_נדרש':    [1,             1,           8],
        'משך_משמרת':        [4,             4,           24],
        'שעות_מנוחה_אחרי': [8,             4,           0],
        'אישור_חפיפה':      [False,         False,       True],
        'שעות_פעילות':      ['all',         'all',       'all'],
        'תפקיד_נדרש':       ['',            'נהג',       ''],
    })

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**👥 תבנית חיילים**")
        st.dataframe(s_ex, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ הורד תבנית חיילים",
            data=to_excel_styled(s_ex, "Soldiers", False),
            file_name="Soldiers_with_Roles.xlsx",
            use_container_width=True,
        )
    with c2:
        st.markdown("**📋 תבנית משימות**")
        st.dataframe(t_ex, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ הורד תבנית משימות",
            data=to_excel_styled(t_ex, "Tasks", False),   # תוקן: שם גיליון ≠ שם קובץ
            file_name="Tasks_with_Roles.xlsx",
            use_container_width=True,
        )

# ────────────────── מדריך ──────────────────
with tab_guide:
    st.markdown("### 📖 מדריך מלא — ניהול תפקידים וכישורים")

    st.markdown("#### 👥 קובץ חיילים — `Soldiers.xlsx`")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:22%">עמודה</th><th style="width:45%">הסבר</th><th>הנחיות</th></tr></thead>
      <tbody>
        <tr>
          <td><b>מספר_אישי</b></td>
          <td>מזהה ייחודי לכל חייל.</td>
          <td>מספר שלם, חובה, ללא כפילויות.</td>
        </tr>
        <tr>
          <td><b>שם</b></td>
          <td>שם החייל כפי שיופיע בלוח.</td>
          <td>שם פרטי + משפחה לבהירות מרבית.</td>
        </tr>
        <tr>
          <td><b>פטורים</b></td>
          <td>קודי משימות שהחייל חסום אליהן ללא קשר לתפקידו.</td>
          <td>קוד_משימה, מופרד בפסיק: <code>1,3</code>. ריק = אין מגבלה.</td>
        </tr>
        <tr>
          <td><b>תפקידים</b></td>
          <td>רשימת הכשרות / כישורים של החייל.</td>
          <td>מופרד בפסיק: <code>נהג, חובש</code>. חייב להתאים למה שרשום ב-<b>תפקיד_נדרש</b> במשימות.</td>
        </tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 📋 קובץ משימות — `Tasks.xlsx`")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:22%">עמודה</th><th style="width:45%">הסבר</th><th>הנחיות</th></tr></thead>
      <tbody>
        <tr>
          <td><b>קוד_משימה</b></td>
          <td>מספר זיהוי ייחודי למשימה.</td>
          <td>חייב להתאים לפטורים בקובץ החיילים.</td>
        </tr>
        <tr>
          <td><b>שם</b></td>
          <td>שם המשימה כפי שיופיע בלוח.</td>
          <td>קצר וברור.</td>
        </tr>
        <tr>
          <td><b>כוח_אדם_נדרש</b></td>
          <td>כמה חיילים בעמדה בכל שעה פעילה.</td>
          <td>האלגוריתם מבטיח בדיוק מספר זה.</td>
        </tr>
        <tr>
          <td><b>משך_משמרת</b></td>
          <td>שעות רצופות שחייל "נעול" במשימה.</td>
          <td>חייל שהוקצה לשעה H נשאר עד H+N. משמרות שחוצות חצות נחסמות אוטומטית.</td>
        </tr>
        <tr>
          <td><b>שעות_מנוחה_אחרי</b></td>
          <td>שעות מנוחה חובה לאחר סיום המשמרת.</td>
          <td>4ש' שמירה + 8ש' מנוחה = חסום 12 שעות מרגע ההתחלה.</td>
        </tr>
        <tr>
          <td><b>אישור_חפיפה</b></td>
          <td>האם ניתן לשבץ חייל למשימה נוספת במקביל?</td>
          <td><code>True</code> — כוננות. <code>False</code> — שמירה / מטבח (חוסמת לחלוטין).</td>
        </tr>
        <tr>
          <td><b>שעות_פעילות</b></td>
          <td>השעות (0–23) שבהן המשימה מתקיימת.</td>
          <td><code>all</code> — 24/7. ספציפי: <code>7,8,9,12,13,14</code>.</td>
        </tr>
        <tr>
          <td><b>תפקיד_נדרש</b></td>
          <td>הכשרה שחייב להיות לחייל כדי לבצע משימה זו.</td>
          <td>לדוגמה: <code>נהג</code> לסיור רכוב. ריק = כל חייל מתאים. חייב להתאים בדיוק לשם התפקיד בקובץ החיילים.</td>
        </tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    <b>💡 הערות חשובות:</b><br>
    • <b>מטרת-על:</b> כיסוי כל העמדות — אילוץ קשיח, לעולם לא מופר.<br>
    • <b>תפקידים:</b> חייל ללא התפקיד הנדרש לא ישובץ למשימה, גם אם הוא זמין.<br>
    • <b>חציית חצות:</b> משמרת שמתחילה בשעה X ומשכה N שעות — אם X+N&gt;24 היא תיחסם. יש לתכנן בהתאם.
    </div>
    """, unsafe_allow_html=True)

# ────────────────── ביצוע שיבוץ ──────────────────
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
            st.markdown(
                f'<div class="error-box">❌ שגיאה בקריאת הקבצים: {e}</div>',
                unsafe_allow_html=True,
            )
            st.stop()

        # בדיקת עמודות חובה
        miss_s = {"מספר_אישי", "שם"} - set(s_df.columns)
        miss_t = {"קוד_משימה", "שם", "כוח_אדם_נדרש"} - set(t_df.columns)
        if miss_s:
            st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ חיילים: {miss_s}</div>', unsafe_allow_html=True)
        if miss_t:
            st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ משימות: {miss_t}</div>', unsafe_allow_html=True)
        if miss_s or miss_t:
            st.stop()

        if st.button("⚙️  צור שבצ\"ק ונתח תובנות", use_container_width=True):
            with st.spinner("מחשב שיבוץ הכולל התאמת תפקידים…"):
                soldiers = [
                    Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים', ''), r.get('תפקידים', ''))
                    for _, r in s_df.iterrows()
                ]
                tasks = [
                    Task(
                        r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'],
                        r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'),
                        r.get('אישור_חפיפה'), r.get('שעות_פעילות'),
                        r.get('תפקיד_נדרש', ''),
                    )
                    for _, r in t_df.iterrows()
                ]

                # בדיקת תפקידים חסרים בסד"כ לפני הרצת הסולבר
                needed_roles   = {t.required_role for t in tasks if t.required_role}
                available_roles = {role for s in soldiers for role in s.roles}
                missing_roles  = needed_roles - available_roles

            if missing_roles:
                st.markdown(
                    f'<div class="error-box">❌ חסרים חיילים עם התפקידים: '
                    f'<b>{", ".join(missing_roles)}</b>.<br>'
                    f'המערכת אינה יכולה לאייש את המשימות הדורשות תפקידים אלה.</div>',
                    unsafe_allow_html=True,
                )
            else:
                with st.spinner("מריץ אופטימיזציה…"):
                    final_df = solve_scheduling(soldiers, tasks)

                if final_df is not None:
                    gap = int(final_df["סך שעות"].max() - final_df["סך שעות"].min())
                    avg = final_df["סך שעות"].mean()

                    # כרטיסי מדד
                    st.markdown(f"""
                    <div class="metric-row">
                      <div class="metric-card">
                        <div class="mc-label">חיילים בשיבוץ</div>
                        <div class="mc-value">{len(soldiers)}</div>
                        <div class="mc-sub">{len(tasks)} משימות</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">ממוצע שעות</div>
                        <div class="mc-value">{avg:.1f}</div>
                        <div class="mc-sub">לחייל</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">פער הוגנות</div>
                        <div class="mc-value">{gap}</div>
                        <div class="mc-sub">{"✅ מצוין" if gap<=2 else "⚠️ סביר" if gap<=5 else "❗ גבוה"}</div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">שעות לילה</div>
                        <div class="mc-value">{int(final_df["שעות לילה"].sum())}</div>
                        <div class="mc-sub">סה"כ בכל הכוח</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown("---")
                    st.subheader("📅 לוח השיבוץ הסופי")
                    st.table(final_df)
                    st.download_button(
                        "📥 הורד לוח שיבוץ (Excel)",
                        data=to_excel_styled(final_df),
                        file_name="Final_Shavtzak.xlsx",
                        use_container_width=True,
                    )

                    st.markdown("---")
                    st.subheader("📊 ניתוח עומסים")
                    fig = px.bar(
                        final_df, x="שם", y="סך שעות",
                        color="סך שעות",
                        color_continuous_scale=["#a8d5a2", "#1a3d17"],
                        title="חלוקת עומס לחייל",
                        text="סך שעות",
                    )
                    fig.update_traces(textposition="outside", marker_line_width=0)
                    fig.update_layout(
                        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                        font=dict(family="Heebo", size=12),
                        coloraxis_showscale=False,
                        xaxis=dict(tickangle=-30, gridcolor="#f0f0f0", title=""),
                        yaxis=dict(gridcolor="#f0f0f0", title="שעות"),
                        margin=dict(t=50, b=80, l=30, r=20),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    with st.expander("💡 למה הלוּז נראה ככה?"):
                        max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                        st.markdown(f"""
**התאמת תפקידים:** המערכת וידאה שרק חיילים עם ההכשרה המתאימה שובצו למשימות הנדרשות (למשל נהגים לסיור רכוב).

**חיילים עמוסים:** {', '.join(max_s)} — ייתכן כי יש להם תפקיד נדיר שדורש אותם לעמדות ספציפיות.

**פערי עומס ({gap} שעות):** {'חלוקה אחידה — אין צורך בפעולה.' if gap<=2 else 'כדאי לשקול הוספת חיילים עם תפקידים זהים לאיזון טוב יותר.'}

**עבודת לילה:** האלגוריתם מצמצם אוטומטית שיבוצי לילה תוך שמירה על כיסוי מלא.
                        """)
                else:
                    st.markdown("""
                    <div class="error-box">
                    ❌ לא נמצא פתרון חוקי.<br><br>
                    <b>המלצות:</b><br>
                    • הגדל את מספר החיילים<br>
                    • הפחת את כוח_אדם_נדרש במשימות<br>
                    • קצר שעות מנוחה חובה<br>
                    • בדוק שאין פטורים ותפקידים שחוסמים יותר מדי
                    </div>
                    """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין לכם תבניות? עברו לטאב <b>תבניות אקסל</b> והורידו.
        </div>
        """, unsafe_allow_html=True)
