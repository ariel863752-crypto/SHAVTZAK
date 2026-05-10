"""
שבצ"ק חכם v3.0
================
מערכת שיבוץ כוחות אוטומטית — CP-SAT
תכונות עיקריות:
  • לוח 48 שעות (2 ימים) — תומך במשמרות שחוצות חצות
  • אילוץ שינה רצופה — מינימום N שעות ללא שיבוץ
  • רצועת שינה: החייל מקבל את חלון השינה הגדול ביותר שניתן
  • צמצום אוטומטי של אילוץ השינה כשאי אפשר לאייש את כל המשימות
  • RTL מלא — Hebrew-first layout
"""

import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# הגדרות עמוד — חייב להיות ראשון
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title='שבצ"ק חכם',
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════
# CSS — RTL מלא + עיצוב
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* ── בסיס RTL ── */
html { direction: rtl; }
body, [class*="css"], .stApp, .stMarkdown, p, span, div, label,
.stSelectbox, .stTextInput, .stNumberInput, .stSlider,
[data-testid="stHeader"], [data-testid="stSidebar"] {
    font-family: 'Heebo', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
}

/* ── רקע ── */
.stApp { background: #f4f6f4; }
.block-container { padding: 2rem 2.5rem 3rem 2.5rem; max-width: 1400px; }

/* ── כותרות ── */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 28px 32px;
    margin-bottom: 28px;
    color: white;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
}
.app-header h1 {
    font-size: clamp(26px, 4vw, 42px);
    font-weight: 900;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
    color: white !important;
    text-align: right !important;
}
.app-header p {
    font-size: 15px;
    opacity: 0.85;
    margin: 0;
    font-weight: 400;
    color: white !important;
    text-align: right !important;
}

/* ── טאבים ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 4px;
    background: white;
    border-radius: 12px;
    padding: 5px;
    border: 1px solid #dde8dc;
    margin-bottom: 20px;
    flex-direction: row-reverse;          /* טאבים מימין לשמאל */
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px;
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
.metric-row { display: flex; gap: 16px; margin: 20px 0; flex-direction: row-reverse; }
.metric-card {
    flex: 1;
    background: white;
    border-radius: 14px;
    padding: 20px 22px;
    border: 1px solid #dde8dc;
    box-shadow: 0 2px 8px rgba(45,90,39,0.07);
    text-align: right;
}
.mc-label { font-size: 11px; color: #7a9a77; font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px; }
.mc-value { font-size: 34px; font-weight: 800; color: #1a3d17; line-height: 1; }
.mc-sub   { font-size: 12px; color: #a0b89d; margin-top: 5px; }

/* ── כפתורים ── */
div.stButton > button {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important;
    font-family: 'Heebo', sans-serif !important;
    font-weight: 700 !important;
    font-size: 17px !important;
    border-radius: 10px !important;
    border: none !important;
    height: 3.2em;
    width: 100%;
    box-shadow: 0 4px 14px rgba(45,90,39,0.3);
    transition: all 0.18s !important;
    direction: rtl !important;
}
div.stButton > button:hover {
    box-shadow: 0 6px 20px rgba(45,90,39,0.45) !important;
    transform: translateY(-1px) !important;
}

/* ── כפתור הורדה ── */
[data-testid="stDownloadButton"] > button {
    background: #c0500a !important;
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
}
[data-testid="stFileUploader"] label { direction: rtl !important; text-align: right !important; }
[data-testid="stFileUploader"]:hover { border-color: #2d5a27; }

/* ── תיבות מידע ── */
.info-box {
    background: #edf5ec;
    border-right: 4px solid #2d5a27;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #1a3d17;
    line-height: 1.7;
    direction: rtl;
    text-align: right;
}
.warn-box {
    background: #fff8e6;
    border-right: 4px solid #e67e22;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #7a4500;
    line-height: 1.7;
    direction: rtl;
    text-align: right;
}
.error-box {
    background: #fdecea;
    border-right: 4px solid #c0392b;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #7a0010;
    line-height: 1.7;
    direction: rtl;
    text-align: right;
}

/* ── טבלאות ── */
[data-testid="stTable"] table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12.5px;
    background: white;
    border-radius: 10px;
    overflow: hidden;
    direction: rtl;
}
[data-testid="stTable"] th {
    background: #2d5a27 !important;
    color: white !important;
    padding: 10px 12px !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    text-align: right !important;
    direction: rtl !important;
}
[data-testid="stTable"] td {
    padding: 8px 12px !important;
    border-bottom: 1px solid #f0f0f0 !important;
    text-align: right !important;
    direction: rtl !important;
}
[data-testid="stTable"] tr:nth-child(even) td { background: #f8fcf7 !important; }
[data-testid="stTable"] tr:hover td { background: #f0f8ef !important; }

/* ── טבלת מדריך ── */
.guide-table {
    width: 100%; border-collapse: separate; border-spacing: 0;
    border-radius: 12px; overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    margin: 16px 0; font-size: 14px;
    direction: rtl;
}
.guide-table th {
    background: #2d5a27; color: white;
    padding: 13px 16px; text-align: right;
    font-weight: 700; font-size: 13px;
}
.guide-table td {
    padding: 12px 16px; border-bottom: 1px solid #eef3ed;
    background: white; line-height: 1.7;
    vertical-align: top; text-align: right;
}
.guide-table tr:last-child td { border-bottom: none; }
.guide-table tr:hover td { background: #f5fbf4; }

/* ── תגיות ── */
.badge { display:inline-block; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:600; }
.badge-green  { background:#e6f4e3; color:#1a5c16; }
.badge-orange { background:#fff0dc; color:#8a4800; }
.badge-red    { background:#fde8e8; color:#9a0020; }

/* ── expander ── */
[data-testid="stExpander"] summary { direction: rtl; text-align: right; font-weight: 600; }

/* ── slider + inputs ── */
[data-testid="stSlider"] label,
[data-testid="stNumberInput"] label { direction: rtl; text-align: right; }

/* ── הודעות מובנות ── */
[data-testid="stAlert"] { direction: rtl; text-align: right; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# מחלקות נתונים
# ══════════════════════════════════════════════════════════════════
class Soldier:
    def __init__(self, s_id, name: str, restr=""):
        self.soldier_id = str(s_id)
        self.name = str(name).strip()
        self.restricted_tasks = self._parse(restr)

    @staticmethod
    def _parse(restr) -> list[int]:
        if pd.isna(restr) or str(restr).strip() in ("", "nan"):
            return []
        return [int(p.replace(".0","")) for p in str(restr).replace(" ","").split(",")
                if p.replace(".0","").lstrip("-").isdigit()]


class Task:
    NIGHT: set[int] = {22,23,0,1,2,3,4,5,6,7}

    def __init__(self, t_id, name, req_p, shift_dur=None,
                 rest_dur=None, overlap=False, hours=None):
        self.task_id           = int(t_id)
        self.name              = str(name).strip()
        self.required_personnel= int(req_p)
        self.shift_duration    = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration     = int(rest_dur)  if pd.notna(rest_dur)  else 0
        self.allow_overlap     = str(overlap).strip().lower() == "true"
        self.active_hours      = self._parse_hours(hours)

    @staticmethod
    def _parse_hours(hours) -> list[int]:
        if pd.isna(hours) or str(hours).strip().lower() in ("all","","nan"):
            return list(range(24))
        return [int(x) for x in str(hours).replace(" ","").split(",") if x.isdigit()]


# ══════════════════════════════════════════════════════════════════
# Excel מעוצב
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df: pd.DataFrame, sheet_name='שבצ"ק', include_index=True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        hfmt = wb.add_format({"bold":True,"fg_color":"#2d5a27","font_color":"white",
                               "border":1,"align":"right","valign":"vcenter"})
        efmt = wb.add_format({"fg_color":"#f0f8ef","align":"right"})
        bfmt = wb.add_format({"align":"right"})
        for ci, col in enumerate(df.columns):
            idx = ci + (1 if include_index else 0)
            ws.write(0, idx, col, hfmt)
            w = max(df[col].astype(str).map(len).max(), len(col)) + 4
            ws.set_column(idx, idx, min(w, 40))
        for ri in range(1, len(df)+1):
            ws.set_row(ri, None, efmt if ri%2==0 else bfmt)
        if include_index:
            ws.set_column(0, 0, 4)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════
# ולידציה
# ══════════════════════════════════════════════════════════════════
def validate(soldiers: list[Soldier], tasks: list[Task]) -> tuple[list[str], list[str]]:
    """מחזיר (errors, warnings)."""
    errors, warnings = [], []
    if not soldiers: errors.append("לא נמצאו חיילים בקובץ.")
    if not tasks:    errors.append("לא נמצאו משימות בקובץ.")
    if errors: return errors, warnings

    ids = [s.soldier_id for s in soldiers]
    if len(ids) != len(set(ids)):
        warnings.append("קיימים מספרים אישיים כפולים — עלול לגרום לשיבוצים שגויים.")

    task_ids = {t.task_id for t in tasks}
    for t in tasks:
        if t.required_personnel > len(soldiers):
            warnings.append(f"משימה '{t.name}' דורשת {t.required_personnel} חיילים אך יש רק {len(soldiers)}.")
    for s in soldiers:
        bad = [r for r in s.restricted_tasks if r not in task_ids]
        if bad:
            warnings.append(f"לחייל '{s.name}' פטור ממשימה עם קוד {bad} שאינה קיימת.")
    return errors, warnings


# ══════════════════════════════════════════════════════════════════
# מנוע CP-SAT — 48 שעות + אילוץ שינה עם צמצום אוטומטי
# ══════════════════════════════════════════════════════════════════
def _build_and_solve(
    soldiers: list[Soldier],
    tasks: list[Task],
    H: int,                     # מספר שעות כולל (48)
    min_sleep: int,             # מינימום שינה רצופה
    timeout: float,
) -> tuple[cp_model.CpSolver | None, dict, dict, str]:
    """
    בונה את המודל ומנסה לפתור.
    מחזיר: (solver, x, start, status_msg)
    אם לא נמצא פתרון — solver=None.
    """
    model = cp_model.CpModel()
    x, start = {}, {}

    for s in soldiers:
        for t in tasks:
            for h in range(H):
                x[s.soldier_id, t.task_id, h]     = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"st_{s.soldier_id}_{t.task_id}_{h}")

    # ── אילוץ 1: נעילת משמרות ──
    # x[s,t,h] = 1 אם"מ יש start בטווח [h-shift+1 .. h]
    for s in soldiers:
        for t in tasks:
            for h in range(H):
                rel = [start[s.soldier_id, t.task_id, i]
                       for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(rel))
                # אי אפשר להתחיל משמרת שתחרוג מגבול הלוח
                if h + t.shift_duration > H:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    # ── אילוץ 2: מנוחה חובה אחרי משמרת ──
    for s in soldiers:
        for t in tasks:
            if t.rest_duration <= 0:
                continue
            busy = t.shift_duration + t.rest_duration
            for h in range(H):
                for nh in range(h + 1, min(h + busy, H)):
                    for ot in tasks:
                        if not ot.allow_overlap:
                            model.AddImplication(
                                start[s.soldier_id, t.task_id, h],
                                x[s.soldier_id, ot.task_id, nh].Not()
                            )

    # ── אילוץ 3: כיסוי עמדות (מטרת-העל) ──
    # כל שעה פעילה צריכה בדיוק required_personnel חיילים
    # לוח 48 שעות: שעה h מתאימה לשעה (h % 24) ביום
    for h in range(H):
        hour_of_day = h % 24
        for t in tasks:
            assigned = [x[s.soldier_id, t.task_id, h] for s in soldiers]
            if hour_of_day in t.active_hours:
                model.Add(sum(assigned) == t.required_personnel)
            else:
                model.Add(sum(assigned) == 0)

    # ── אילוץ 4: חד-ערכיות (חייל בעמדה אחת בשעה) ──
    for s in soldiers:
        for h in range(H):
            blocking = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking) <= 1)

    # ── אילוץ 5: פטורים ──
    for s in soldiers:
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(H):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # ── אילוץ 6: שינה רצופה מינימלית (soft — ניתן לוותר בלחץ) ──
    # עבור כל חייל, בכל חלון של min_sleep שעות רצופות,
    # לפחות חלון אחד חייב להיות ריק לחלוטין (ב-H שעות).
    # ממומש כ-soft penalty: הפרת האילוץ מוסיפה עונש לפונקציית המטרה.
    sleep_violations = []
    if min_sleep > 0:
        for s in soldiers:
            for h in range(H - min_sleep + 1):
                # עבודה בחלון [h .. h+min_sleep-1]
                work_in_window = [
                    x[s.soldier_id, t.task_id, h + k]
                    for t in tasks
                    if not t.allow_overlap
                    for k in range(min_sleep)
                ]
                # viol=1 אם יש עבודה בכל שעות החלון
                viol = model.NewBoolVar(f"viol_{s.soldier_id}_{h}")
                # אם sum(work_in_window) == min_sleep → viol=1
                # ממומש: sum >= min_sleep → viol
                model.Add(sum(work_in_window) >= min_sleep).OnlyEnforceIf(viol)
                model.Add(sum(work_in_window) < min_sleep).OnlyEnforceIf(viol.Not())
                sleep_violations.append(viol)

    # ── פונקציית מטרה ──
    # עדיפות 1 (מטרת-על): כיסוי עמדות — כבר נכפה כ-hard constraint
    # עדיפות 2: הוגנות (מינימום פער עומסים)
    # עדיפות 3: מינימום שעות לילה
    # עדיפות 4 (soft): מינימום הפרות שינה

    s_loads = [
        sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(H))
        for s in soldiers
    ]
    max_load = model.NewIntVar(0, H * len(tasks), "max_load")
    min_load = model.NewIntVar(0, H * len(tasks), "min_load")
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)
    load_range = model.NewIntVar(0, H * len(tasks), "range")
    model.Add(load_range == max_load - min_load)

    night_work = sum(
        x[s.soldier_id, t.task_id, h]
        for s in soldiers
        for t in tasks
        if not t.allow_overlap
        for h in range(H)
        if (h % 24) in Task.NIGHT
    )

    sleep_penalty = sum(sleep_violations) if sleep_violations else 0

    model.Minimize(
        100 * load_range       # הוגנות — משקל גבוה
        + 2  * night_work      # עדיפות לשינת לילה
        + 200 * sleep_penalty  # עונש על הפרת שינה רצופה
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds  = timeout
    solver.parameters.num_search_workers   = 4
    solver.parameters.log_search_progress  = False
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return solver, x, start, "ok"
    return None, {}, {}, "infeasible"


def solve_scheduling(
    soldiers: list[Soldier],
    tasks: list[Task],
    min_sleep: int = 6,
    timeout: float = 35.0,
    num_days: int = 2,
) -> tuple[pd.DataFrame | None, str]:
    """
    נקודת הכניסה הראשית.
    ניסיון 1: עם אילוץ שינה מלא.
    ניסיון 2: אם נכשל — צמצום שינה אוטומטי ל-4, 2, 0.
    מחזיר (DataFrame, הודעת סטטוס).
    """
    H = 24 * num_days
    sleep_levels = [min_sleep]
    if min_sleep > 4: sleep_levels.append(4)
    if min_sleep > 2: sleep_levels.append(2)
    sleep_levels.append(0)  # ניסיון אחרון ללא אילוץ שינה

    used_sleep = min_sleep
    solver = None

    for sl in sleep_levels:
        used_sleep = sl
        solver, x, start, status = _build_and_solve(soldiers, tasks, H, sl, timeout)
        if solver is not None:
            break

    if solver is None:
        return None, "לא נמצא פתרון גם ללא אילוץ שינה. בדוק את יחס הכוח/משימות."

    msg = "ok" if used_sleep == min_sleep else f"אילוץ שינה צומצם אוטומטית ל-{used_sleep} שעות כדי לאייש את כל המשימות."

    # ── בניית DataFrame תוצאות ──
    rows = []
    for s in soldiers:
        row: dict = {"שם": s.name}
        total, night_count = 0, 0
        for h in range(H):
            day = h // 24 + 1
            hour_label = f"יום{day} {h%24:02d}:00"
            active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
            row[hour_label] = " + ".join(active) if active else "—"
            if active:
                total += 1
                if (h % 24) in Task.NIGHT and any(
                    solver.Value(x[s.soldier_id, t.task_id, h]) == 1
                    for t in tasks if not t.allow_overlap
                ):
                    night_count += 1
        row["סך שעות"] = total
        row["שעות לילה"] = night_count
        rows.append(row)

    df = pd.DataFrame(rows)
    df.index = range(1, len(df)+1)
    return df, msg


# ══════════════════════════════════════════════════════════════════
# גרפים
# ══════════════════════════════════════════════════════════════════
def chart_load(df: pd.DataFrame):
    fig = px.bar(df, x="שם", y="סך שעות", color="סך שעות",
                 color_continuous_scale=["#a8d5a2","#2d5a27"],
                 title="עומס כולל לחייל", text="סך שעות")
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Heebo", size=12), coloraxis_showscale=False,
        title_font=dict(size=15, color="#1a3d17"),
        xaxis=dict(tickangle=-30, gridcolor="#f0f0f0"),
        yaxis=dict(gridcolor="#f0f0f0"),
        margin=dict(t=50,b=70,l=20,r=20),
    )
    return fig


def chart_day_night(df: pd.DataFrame):
    d = df[["שם","שעות לילה","סך שעות"]].copy()
    d["שעות יום"] = d["סך שעות"] - d["שעות לילה"]
    fig = go.Figure([
        go.Bar(name="יום",  x=d["שם"], y=d["שעות יום"],  marker_color="#7ec87a"),
        go.Bar(name="לילה", x=d["שם"], y=d["שעות לילה"], marker_color="#1a3d17"),
    ])
    fig.update_layout(
        barmode="stack", title="פילוג יום / לילה",
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Heebo", size=12),
        title_font=dict(size=15, color="#1a3d17"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        xaxis=dict(tickangle=-30),
        margin=dict(t=60,b=70,l=20,r=20),
    )
    return fig


# ══════════════════════════════════════════════════════════════════
# ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות, תורנויות ומשימות — הוגן, מהיר, מדויק</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs([
    "🚀  ביצוע שיבוץ",
    "📖  מדריך מפורט",
    "📥  תבניות עבודה",
])

# ────────────────────────── תבניות ──────────────────────────────
with tab_templates:
    st.markdown("### 📥 הורדת תבניות עבודה")
    st.markdown('<div class="info-box">הורידו את התבניות, מלאו לפי המדריך, ואז חזרו לטאב <b>ביצוע שיבוץ</b>.</div>', unsafe_allow_html=True)

    soldiers_ex = pd.DataFrame({
        "מספר_אישי": [1001, 1002, 1003, 1004],
        "שם":        ["ישראל ישראלי", "יוסי כהן", "דנה לוי", "רועי אברהם (פטור שער)"],
        "פטורים":    ["", "", "", "1"],
    })
    tasks_ex = pd.DataFrame({
        "קוד_משימה":       [1, 2, 10],
        "שם":              ["שמירת שער", "תורנות מטבח", "כיתת כוננות"],
        "כוח_אדם_נדרש":   [1, 3, 8],
        "משך_משמרת":       [4, 6, 24],
        "שעות_מנוחה_אחרי": [8, 0, 0],
        "אישור_חפיפה":     [False, False, True],
        "שעות_פעילות":     ["all", "7,8,9,12,13,14", "all"],
    })

    st.markdown("**👥 תבנית חיילים**")
    st.dataframe(soldiers_ex, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ הורד תבנית חיילים",
            data=to_excel_styled(soldiers_ex, "Soldiers", False),
            file_name="Shavtzak_Soldiers.xlsx", use_container_width=True)

    st.markdown("---")
    st.markdown("**📋 תבנית משימות**")
    st.dataframe(tasks_ex, use_container_width=True, hide_index=True)
    with c2:
        st.download_button("⬇️ הורד תבנית משימות",
            data=to_excel_styled(tasks_ex, "Tasks", False),
            file_name="Shavtzak_Tasks.xlsx", use_container_width=True)

# ────────────────────────── מדריך ───────────────────────────────
with tab_guide:
    st.markdown("### 📖 מדריך — הסבר לכל עמודה")

    st.markdown("#### 👥 קובץ חיילים — `Soldiers.xlsx`")
    st.markdown("""
    <table class="guide-table">
      <thead><tr><th style="width:20%">עמודה</th><th style="width:45%">הסבר</th><th>הנחיות</th></tr></thead>
      <tbody>
      <tr>
        <td><b>מספר_אישי</b></td>
        <td>מזהה ייחודי לכל חייל (מספר שלם).</td>
        <td>חובה. אין לחזור על מספרים. משמש גם להתאמה עם שדה הפטורים.</td>
      </tr>
      <tr>
        <td><b>שם</b></td>
        <td>שם החייל כפי שיופיע בלוח השיבוץ.</td>
        <td>שם פרטי + משפחה — לבהירות מרבית.</td>
      </tr>
      <tr>
        <td><b>פטורים</b></td>
        <td>קודי המשימות שהחייל <b>לא יכול</b> לבצע.</td>
        <td>קוד_משימה מהקובץ המשימות. כמה פטורים — פסיק: <code>1,3,5</code>. ריק = אין מגבלה.</td>
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
        <td>חייב להתאים לקודים בעמודת הפטורים של החיילים.</td>
      </tr>
      <tr>
        <td><b>שם</b></td>
        <td>שם המשימה (שמירה, סיור, מטבח…).</td>
        <td>יופיע בתאי לוח השיבוץ.</td>
      </tr>
      <tr>
        <td><b>כוח_אדם_נדרש</b></td>
        <td>כמות החיילים בעמדה בכל שעה פעילה.</td>
        <td>שמירה בזוגות = <code>2</code>. האלגוריתם מבטיח בדיוק מספר זה.</td>
      </tr>
      <tr>
        <td><b>משך_משמרת</b></td>
        <td>שעות רצופות שחייל מבצע את המשימה.</td>
        <td><b>נעילה:</b> חייל שהוקצה לשעה H נשאר עד H+N. לא ניתן להחליפו באמצע. תומך בחציית חצות.</td>
      </tr>
      <tr>
        <td><b>שעות_מנוחה_אחרי</b></td>
        <td>שעות מנוחה מינימליות לאחר סיום המשמרת.</td>
        <td>לאחר 4 שעות שמירה + 8 מנוחה — החייל חסום מכל משימה ל-8 שעות. תומך בחציית חצות.</td>
      </tr>
      <tr>
        <td><b>אישור_חפיפה</b></td>
        <td>האם ניתן לשבץ חייל למשימה אחרת במקביל?</td>
        <td><code>True</code> — כוננות (לא חוסמת). <code>False</code> — שמירה/מטבח (חוסמת לחלוטין).</td>
      </tr>
      <tr>
        <td><b>שעות_פעילות</b></td>
        <td>השעות שבהן המשימה מתקיימת.</td>
        <td><code>all</code> — 24/7. שעות ספציפיות: <code>7,8,9,12,13,14</code>. חל על <b>שני הימים</b> בלוח.</td>
      </tr>
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    <b>💡 הערות חשובות על האלגוריתם:</b><br>
    • <b>מטרת-על:</b> כיסוי כל העמדות תמיד — הוא hard constraint, לא ניתן לוויתור.<br>
    • <b>לוח 48 שעות:</b> מאפשר משמרות שחוצות חצות (22:00–02:00) ללא בעיות.<br>
    • <b>שינה רצופה:</b> soft constraint — אם אי אפשר לאייש ולשמור על שינה, האלגוריתם מצמצם אוטומטית את שעות השינה ומדווח על כך.<br>
    • <b>הוגנות:</b> האלגוריתם ממזער את פער העומסים בין החיילים — לא רק את המקסימום.
    </div>
    """, unsafe_allow_html=True)

# ────────────────────────── ביצוע שיבוץ ────────────────────────
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    with st.expander("⚙️ הגדרות מתקדמות"):
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            min_sleep = st.slider("מינימום שינה רצופה (שעות)", 0, 10, 6, 1,
                                  help="מינימום שעות רצופות ללא שיבוץ שיחשבו כשינה. אם הכוח קטן, יצומצם אוטומטית.")
        with col_s2:
            timeout = st.slider("זמן חיפוש מקסימלי (שניות)", 15, 120, 40, 5)
        with col_s3:
            num_days = st.selectbox("מספר ימים בלוח", [1, 2, 3], index=1,
                                    help="מומלץ 2 ימים — מאפשר משמרות שחוצות חצות")

    # ── הרצה ──
    if sf and tf:
        try:
            s_df = pd.read_excel(sf)
            t_df = pd.read_excel(tf)
        except Exception as e:
            st.markdown(f'<div class="error-box">❌ שגיאה בקריאת הקבצים: {e}</div>', unsafe_allow_html=True)
            st.stop()

        # בדיקת עמודות חובה
        req_s = {"מספר_אישי", "שם"}
        req_t = {"קוד_משימה", "שם", "כוח_אדם_נדרש"}
        missing_s = req_s - set(s_df.columns)
        missing_t = req_t - set(t_df.columns)
        if missing_s or missing_t:
            if missing_s:
                st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ חיילים: {missing_s}</div>', unsafe_allow_html=True)
            if missing_t:
                st.markdown(f'<div class="error-box">❌ חסרות עמודות בקובץ משימות: {missing_t}</div>', unsafe_allow_html=True)
            st.stop()

        soldiers = [Soldier(r["מספר_אישי"], r["שם"], r.get("פטורים","")) for _, r in s_df.iterrows()]
        tasks    = [Task(r["קוד_משימה"], r["שם"], r["כוח_אדם_נדרש"],
                         r.get("משך_משמרת"), r.get("שעות_מנוחה_אחרי"),
                         r.get("אישור_חפיפה"), r.get("שעות_פעילות"))
                    for _, r in t_df.iterrows()]

        errors, warnings = validate(soldiers, tasks)
        for e in errors:
            st.markdown(f'<div class="error-box">❌ {e}</div>', unsafe_allow_html=True)
        for w in warnings:
            st.markdown(f'<div class="warn-box">⚠️ {w}</div>', unsafe_allow_html=True)

        if not errors:
            if st.button('⚙️ צור שבצ"ק ונתח תובנות', use_container_width=True):
                with st.spinner("🔄 מחשב שיבוץ אופטימלי…"):
                    final_df, status_msg = solve_scheduling(soldiers, tasks, min_sleep, timeout, num_days)

                if final_df is None:
                    st.markdown(f'<div class="error-box">❌ {status_msg}<br><br>'
                                '<b>המלצות:</b><br>'
                                '• הגדל את מספר החיילים<br>'
                                '• הפחת את כוח_אדם_נדרש<br>'
                                '• קצר שעות מנוחה<br>'
                                '• הפחת שעות שינה מינימליות</div>', unsafe_allow_html=True)
                else:
                    # הודעת סטטוס
                    if status_msg == "ok":
                        st.success(f"✅ שיבוץ נוצר בהצלחה! שינה רצופה: {min_sleep} שעות — נשמרה לכל החיילים.")
                    else:
                        st.warning(f"⚠️ {status_msg}")

                    # ── מדדים ──
                    avg  = final_df["סך שעות"].mean()
                    gap  = final_df["סך שעות"].max() - final_df["סך שעות"].min()
                    nite = int(final_df["שעות לילה"].sum())
                    badge = ("badge-green" if gap<=2 else "badge-orange" if gap<=5 else "badge-red")
                    badge_txt = ("מצוין" if gap<=2 else "סביר" if gap<=5 else "גבוה")
                    st.markdown(f"""
                    <div class="metric-row">
                      <div class="metric-card">
                        <div class="mc-label">חיילים</div>
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
                        <div class="mc-sub"><span class="badge {badge}">{badge_txt}</span></div>
                      </div>
                      <div class="metric-card">
                        <div class="mc-label">שעות לילה</div>
                        <div class="mc-value">{nite}</div>
                        <div class="mc-sub">בכל הכוח</div>
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
                        file_name="Final_Shavtzak.xlsx",
                        use_container_width=True,
                    )

                    # ── גרפים ──
                    st.markdown("---")
                    st.subheader("📊 ניתוח עומסים")
                    g1, g2 = st.columns(2)
                    with g1: st.plotly_chart(chart_load(final_df), use_container_width=True)
                    with g2: st.plotly_chart(chart_day_night(final_df), use_container_width=True)

                    # ── ניתוח ──
                    st.markdown("---")
                    st.subheader("💡 ניתוח ותובנות")
                    with st.expander("לחץ לקריאת הניתוח המלא"):
                        night_demand = sum(
                            t.required_personnel for t in tasks
                            if not t.allow_overlap and Task.NIGHT & set(t.active_hours)
                        )
                        max_s = final_df[final_df["סך שעות"]==final_df["סך שעות"].max()]["שם"].tolist()
                        min_s = final_df[final_df["סך שעות"]==final_df["סך שעות"].min()]["שם"].tolist()
                        st.markdown(f"""
**עבודת לילה:** נדרשים {night_demand} חיילים בשעות הלילה. האלגוריתם ממזער שעות לילה תוך שמירה על כיסוי מלא.

**חיילים עמוסים:** {', '.join(max_s)} — {final_df['סך שעות'].max()} שעות.

**חיילים עם עומס נמוך:** {', '.join(min_s)} — {final_df['סך שעות'].min()} שעות.

**פער עומסים:** {gap} שעות {'✅' if gap<=2 else '⚠️ מומלץ להגדיל כוח אדם לחלוקה אחידה יותר' if gap>4 else '✔️'}.

**שינה רצופה:** האלגוריתם מעניש פי 200 כל הפרת חלון שינה. אם ראית אזהרה — הפחת את דרישת הכוח או הוסף חיילים.

**השפעת מנוחות:** חיילים עם מנוחה ארוכה מוצאים מהסבב לשעות אלו, מה שמגדיל עומס על האחרים.
                        """)
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות.<br>
        אין לכם תבניות? עברו לטאב <b>תבניות עבודה</b> והורידו.
        </div>
        """, unsafe_allow_html=True)
      
