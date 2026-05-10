"""
שבצ"ק חכם — מערכת שיבוץ כוחות אוטומטית
=========================================
גרסה: 2.0
תלויות: streamlit, pandas, ortools, plotly, xlsxwriter
"""

import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# ──────────────────────────────────────────────────────────────────
# 0. הגדרות עמוד
# ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="שבצ\"ק חכם",
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────────
# 1. CSS גלובלי — RTL + עיצוב מקצועי
# ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── ייבוא גופן ── */
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800&display=swap');

/* ── ריסט בסיסי ── */
*, *::before, *::after { box-sizing: border-box; }

html, body, [class*="css"] {
    font-family: 'Heebo', sans-serif;
    direction: rtl;
    text-align: right;
}

.stApp {
    background: linear-gradient(135deg, #f0f4f0 0%, #e8f0e8 50%, #f5f5f0 100%);
    min-height: 100vh;
}

/* ── כותרת ראשית ── */
.main-title {
    font-size: clamp(28px, 5vw, 46px);
    font-weight: 800;
    color: #1a3d17;
    letter-spacing: -0.5px;
    line-height: 1.15;
    margin-bottom: 6px;
}
.main-subtitle {
    font-size: 16px;
    color: #5a7a57;
    font-weight: 400;
    margin-bottom: 28px;
}

/* ── כרטיסי מדד ── */
.metric-card {
    background: white;
    border-radius: 14px;
    padding: 20px 22px;
    box-shadow: 0 2px 12px rgba(45,90,39,0.08);
    border: 1px solid rgba(45,90,39,0.1);
    transition: transform 0.2s, box-shadow 0.2s;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(45,90,39,0.13);
}
.metric-label { font-size: 12px; color: #7a9a77; font-weight: 500; text-transform: uppercase; letter-spacing: 0.6px; }
.metric-value { font-size: 32px; font-weight: 700; color: #1a3d17; line-height: 1.1; margin-top: 4px; }
.metric-note  { font-size: 12px; color: #a0b89d; margin-top: 4px; }

/* ── אזור העלאת קבצים ── */
[data-testid="stFileUploader"] {
    background: white;
    border-radius: 12px;
    padding: 16px;
    border: 2px dashed #c8ddc5;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #2d5a27; }

/* ── כפתורים ── */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #2d5a27 0%, #3d7a35 100%) !important;
    color: white !important;
    font-weight: 700 !important;
    font-size: 17px !important;
    border-radius: 10px !important;
    border: none !important;
    width: 100%;
    height: 3.4em;
    letter-spacing: 0.2px;
    box-shadow: 0 4px 15px rgba(45,90,39,0.3);
    transition: all 0.2s !important;
}
div.stButton > button:first-child:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 20px rgba(45,90,39,0.4) !important;
}

/* ── כפתור הורדה ── */
[data-testid="stDownloadButton"] > button {
    background: #e67e22 !important;
    color: white !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    border: none !important;
    box-shadow: 0 3px 10px rgba(230,126,34,0.3);
}

/* ── טאבים ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 6px;
    background: rgba(255,255,255,0.6);
    border-radius: 12px;
    padding: 6px;
    border: 1px solid rgba(45,90,39,0.1);
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px;
    font-weight: 600;
    font-size: 14px;
    padding: 8px 18px;
    color: #5a7a57;
    transition: all 0.2s;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: white !important;
    color: #2d5a27 !important;
    box-shadow: 0 2px 8px rgba(45,90,39,0.12);
}

/* ── טבלאות ── */
[data-testid="stTable"] table {
    border-collapse: collapse;
    width: 100%;
    font-size: 13px;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
}
[data-testid="stTable"] th {
    background: #2d5a27 !important;
    color: white !important;
    padding: 10px 12px !important;
    font-weight: 600;
    font-size: 12px;
    letter-spacing: 0.3px;
}
[data-testid="stTable"] td {
    padding: 9px 12px !important;
    border-bottom: 1px solid #f0f0f0;
}
[data-testid="stTable"] tr:nth-child(even) td { background: #fafcfa; }
[data-testid="stTable"] tr:hover td { background: #f0f8ef !important; }

/* ── טבלת מדריך ── */
.guide-table {
    width: 100%; border-collapse: separate; border-spacing: 0;
    border-radius: 12px; overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    margin-top: 16px; font-size: 14px;
}
.guide-table th {
    background: #2d5a27; color: white;
    padding: 13px 16px; text-align: right;
    font-weight: 600; font-size: 13px;
}
.guide-table td {
    padding: 12px 16px; border-bottom: 1px solid #eef3ed;
    background: white; line-height: 1.65;
    vertical-align: top;
}
.guide-table tr:last-child td { border-bottom: none; }
.guide-table tr:hover td { background: #f5fbf4; }

/* ── תגיות סטטוס ── */
.badge {
    display: inline-block; padding: 3px 10px;
    border-radius: 20px; font-size: 12px; font-weight: 600;
}
.badge-green  { background: #e6f4e3; color: #2d5a27; }
.badge-orange { background: #fef3e3; color: #b85c00; }
.badge-red    { background: #fde8e8; color: #b00020; }

/* ── הודעות ── */
.info-box {
    background: linear-gradient(135deg, #f0f8ef, #e8f4e6);
    border-right: 4px solid #2d5a27;
    border-radius: 0 10px 10px 0;
    padding: 14px 18px;
    margin: 14px 0;
    font-size: 14px;
    color: #1a3d17;
}

/* ── ספינר ── */
[data-testid="stSpinner"] { color: #2d5a27 !important; }

/* ── divider ── */
hr { border-color: rgba(45,90,39,0.1) !important; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────
# 2. מחלקות נתונים
# ──────────────────────────────────────────────────────────────────
class Soldier:
    """מייצג חייל עם מגבלות שיבוץ."""

    def __init__(self, s_id: int | str, name: str, restr: str = ""):
        self.soldier_id = str(s_id)
        self.name = str(name)
        self.restricted_tasks = self._parse_restrictions(restr)

    @staticmethod
    def _parse_restrictions(restr) -> list[int]:
        if pd.isna(restr) or str(restr).strip() in ("", "nan"):
            return []
        parts = str(restr).replace(" ", "").split(",")
        result = []
        for p in parts:
            clean = p.replace(".0", "")
            if clean.lstrip("-").isdigit():
                result.append(int(clean))
        return result


class Task:
    """מייצג משימה/עמדה עם פרמטרי תפעול."""

    NIGHT_HOURS: set[int] = {22, 23, 0, 1, 2, 3, 4, 5, 6, 7}

    def __init__(
        self,
        t_id: int,
        name: str,
        req_p: int,
        shift_dur=None,
        rest_dur=None,
        overlap=False,
        hours=None,
    ):
        self.task_id = int(t_id)
        self.name = str(name)
        self.required_personnel = int(req_p)
        self.shift_duration = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration = int(rest_dur) if pd.notna(rest_dur) else 0
        self.allow_overlap = str(overlap).strip().lower() == "true"
        self.active_hours = self._parse_hours(hours)

    @staticmethod
    def _parse_hours(hours) -> list[int]:
        if pd.isna(hours) or str(hours).strip().lower() in ("all", "", "nan"):
            return list(range(24))
        parts = str(hours).replace(" ", "").split(",")
        return [int(x) for x in parts if x.isdigit()]


# ──────────────────────────────────────────────────────────────────
# 3. אקסל מעוצב
# ──────────────────────────────────────────────────────────────────
def to_excel_styled(df: pd.DataFrame, sheet_name: str = 'שבצ"ק', include_index: bool = True) -> bytes:
    """יוצר קובץ Excel עם עיצוב, כותרות צבעוניות וגודל עמודות אוטומטי."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        header_fmt = wb.add_format({
            "bold": True, "fg_color": "#2d5a27", "font_color": "white",
            "border": 1, "align": "right", "valign": "vcenter",
        })
        alt_fmt = wb.add_format({"fg_color": "#f0f8ef", "align": "right"})
        base_fmt = wb.add_format({"align": "right"})

        for col_num, col_val in enumerate(df.columns.values):
            col_idx = col_num + (1 if include_index else 0)
            ws.write(0, col_idx, col_val, header_fmt)
            max_len = max(df[col_val].astype(str).map(len).max(), len(col_val)) + 4
            ws.set_column(col_idx, col_idx, min(max_len, 40))

        for row_num in range(1, len(df) + 1):
            fmt = alt_fmt if row_num % 2 == 0 else base_fmt
            ws.set_row(row_num, None, fmt)

        if include_index:
            ws.set_column(0, 0, 4)

    return output.getvalue()


# ──────────────────────────────────────────────────────────────────
# 4. ולידציה
# ──────────────────────────────────────────────────────────────────
def validate_inputs(soldiers: list[Soldier], tasks: list[Task]) -> list[str]:
    """מחזיר רשימת שגיאות/אזהרות לפני הרצת הסולבר."""
    warnings = []

    if not soldiers:
        warnings.append("❌ לא נמצאו חיילים בקובץ.")
        return warnings

    if not tasks:
        warnings.append("❌ לא נמצאו משימות בקובץ.")
        return warnings

    # בדיקת כפילות מזהים
    ids = [s.soldier_id for s in soldiers]
    if len(ids) != len(set(ids)):
        warnings.append("⚠️ קיימים מספרים אישיים כפולים — עלול לגרום לשיבוצים שגויים.")

    # בדיקת היתכנות גסה: האם יש מספיק חיילים?
    for t in tasks:
        if t.required_personnel > len(soldiers):
            warnings.append(
                f"⚠️ משימה '{t.name}' דורשת {t.required_personnel} חיילים אך יש רק {len(soldiers)}."
            )

    # בדיקת פטורים תקינים
    task_ids = {t.task_id for t in tasks}
    for s in soldiers:
        invalid = [r for r in s.restricted_tasks if r not in task_ids]
        if invalid:
            warnings.append(f"⚠️ לחייל '{s.name}' יש פטור ממשימה עם קוד {invalid} שלא קיימת בקובץ המשימות.")

    return warnings


# ──────────────────────────────────────────────────────────────────
# 5. מנוע CP-SAT
# ──────────────────────────────────────────────────────────────────
def solve_scheduling(
    soldiers: list[Soldier],
    tasks: list[Task],
    num_hours: int = 24,
    solver_timeout: float = 30.0,
) -> pd.DataFrame | None:
    """
    פותר בעיית שיבוץ כוחות בגישת CP-SAT.

    מאפיינים:
    - נעילת משמרות (shift lock): חייל שהוקצה לשעה H נשאר עד H+shift_duration.
    - כיבוד מנוחה: לאחר משמרת + מנוחה, החייל חסום מכל משימה לא-חופפת.
    - פטורים: חייל עם פטור ממשימה לא ישובץ אליה לעולם.
    - הוגנות: מינימיזציה של פער בין עומסי החיילים + צמצום עבודת לילה.
    """
    model = cp_model.CpModel()
    night_hours = Task.NIGHT_HOURS

    # ── משתני החלטה ──
    x: dict = {}      # x[s, t, h] = 1 אם חייל s בעמדה t בשעה h
    start: dict = {}  # start[s, t, h] = 1 אם חייל s מתחיל משמרת t בשעה h

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(
                    f"x_{s.soldier_id}_{t.task_id}_{h}"
                )
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(
                    f"st_{s.soldier_id}_{t.task_id}_{h}"
                )

    # ── אילוץ 1: קישור start → x (נעילת משמרות) ──
    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                # x[h] == sum of starts that cover h
                relevant = [
                    start[s.soldier_id, t.task_id, i]
                    for i in range(max(0, h - t.shift_duration + 1), h + 1)
                ]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant))

                # אסור להתחיל משמרת אם היא חוצה את גבול 24 השעות
                if h + t.shift_duration > num_hours:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    # ── אילוץ 2: מנוחה בסיום משמרת ──
    for s in soldiers:
        for t in tasks:
            if t.rest_duration <= 0:
                continue
            total_busy = t.shift_duration + t.rest_duration
            for h in range(num_hours):
                for next_h in range(h + 1, min(h + total_busy, num_hours)):
                    for other_t in tasks:
                        if not other_t.allow_overlap:
                            model.AddImplication(
                                start[s.soldier_id, t.task_id, h],
                                x[s.soldier_id, other_t.task_id, next_h].Not(),
                            )

    # ── אילוץ 3: כיסוי עמדות בכל שעה ──
    for h in range(num_hours):
        for t in tasks:
            assigned = [x[s.soldier_id, t.task_id, h] for s in soldiers]
            if h in t.active_hours:
                model.Add(sum(assigned) == t.required_personnel)
            else:
                model.Add(sum(assigned) == 0)

    # ── אילוץ 4: חייל בעמדה אחת בכל שעה (ללא חפיפה) ──
    for s in soldiers:
        for h in range(num_hours):
            blocking = [
                x[s.soldier_id, t.task_id, h]
                for t in tasks
                if not t.allow_overlap
            ]
            model.Add(sum(blocking) <= 1)

    # ── אילוץ 5: כיבוד פטורים ──
    for s in soldiers:
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # ── פונקציית מטרה: הוגנות + צמצום לילה ──
    s_loads = [
        sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours))
        for s in soldiers
    ]
    max_load = model.NewIntVar(0, num_hours * len(tasks), "max_load")
    min_load = model.NewIntVar(0, num_hours * len(tasks), "min_load")
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)

    load_range = model.NewIntVar(0, num_hours * len(tasks), "load_range")
    model.Add(load_range == max_load - min_load)

    night_work = sum(
        x[s.soldier_id, t.task_id, h]
        for s in soldiers
        for t in tasks
        if not t.allow_overlap
        for h in night_hours
        if h < num_hours
    )

    model.Minimize(50 * load_range + night_work)

    # ── פתרון ──
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = solver_timeout
    solver.parameters.num_search_workers = 4   # ניצול מקבילי
    solver.parameters.log_search_progress = False

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # ── בניית DataFrame תוצאות ──
    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []

    for s in soldiers:
        row: dict = {"שם": s.name}
        total_hours = 0
        night_count = 0

        for h in range(num_hours):
            active = [
                t.name
                for t in tasks
                if solver.Value(x[s.soldier_id, t.task_id, h]) == 1
            ]
            row[hour_labels[h]] = " + ".join(active) if active else "—"
            if active:
                total_hours += 1
                if h in night_hours and any(
                    solver.Value(x[s.soldier_id, t.task_id, h]) == 1
                    for t in tasks
                    if not t.allow_overlap
                ):
                    night_count += 1

        row["סך שעות"] = total_hours
        row["שעות לילה"] = night_count
        rows.append(row)

    result_df = pd.DataFrame(rows)
    result_df.index = range(1, len(result_df) + 1)
    return result_df


# ──────────────────────────────────────────────────────────────────
# 6. רכיבי UI עזר
# ──────────────────────────────────────────────────────────────────
def metric_card(label: str, value: str, note: str = "") -> str:
    """מחזיר HTML של כרטיס מדד."""
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {"<div class='metric-note'>" + note + "</div>" if note else ""}
    </div>
    """


def render_summary_metrics(df: pd.DataFrame, soldiers: list[Soldier], tasks: list[Task]) -> None:
    """מציג שורת סיכום מדדים עיקריים."""
    c1, c2, c3, c4 = st.columns(4)
    avg_hrs = df["סך שעות"].mean()
    max_hrs = df["סך שעות"].max()
    min_hrs = df["סך שעות"].min()
    fairness = max_hrs - min_hrs

    with c1:
        st.markdown(metric_card("חיילים בשיבוץ", str(len(soldiers)), f"{len(tasks)} משימות"), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card("ממוצע שעות", f"{avg_hrs:.1f}", "לחייל"), unsafe_allow_html=True)
    with c3:
        color = "badge-green" if fairness <= 2 else ("badge-orange" if fairness <= 4 else "badge-red")
        st.markdown(
            metric_card("פער הוגנות", f"{fairness}", f'<span class="badge {color}">{"מצוין" if fairness<=2 else "סביר" if fairness<=4 else "גבוה"}</span>'),
            unsafe_allow_html=True,
        )
    with c4:
        total_night = df["שעות לילה"].sum()
        st.markdown(metric_card("סה\"כ שעות לילה", str(total_night), "בכל הכוח"), unsafe_allow_html=True)


def render_load_chart(df: pd.DataFrame) -> None:
    """גרף עמודות עומס חיילים."""
    fig = px.bar(
        df, x="שם", y="סך שעות",
        color="סך שעות",
        title="חלוקת עומס — שעות שיבוץ לחייל",
        color_continuous_scale=["#a8d5a2", "#2d5a27"],
        text="סך שעות",
    )
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Heebo", size=12, color="#333"),
        title_font=dict(size=16, color="#1a3d17"),
        coloraxis_showscale=False,
        xaxis=dict(tickangle=-30, gridcolor="#f0f0f0"),
        yaxis=dict(gridcolor="#f0f0f0"),
        margin=dict(t=50, b=60, l=20, r=20),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_night_chart(df: pd.DataFrame) -> None:
    """גרף שעות לילה ויחס לילה/יום."""
    df_plot = df[["שם", "שעות לילה", "סך שעות"]].copy()
    df_plot["שעות יום"] = df_plot["סך שעות"] - df_plot["שעות לילה"]
    fig = go.Figure()
    fig.add_trace(go.Bar(name="יום", x=df_plot["שם"], y=df_plot["שעות יום"], marker_color="#a8d5a2"))
    fig.add_trace(go.Bar(name="לילה", x=df_plot["שם"], y=df_plot["שעות לילה"], marker_color="#1a3d17"))
    fig.update_layout(
        barmode="stack",
        title="פילוג שעות יום/לילה לחייל",
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Heebo", size=12),
        title_font=dict(size=16, color="#1a3d17"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=60, l=20, r=20),
        xaxis=dict(tickangle=-30),
    )
    st.plotly_chart(fig, use_container_width=True)


# ──────────────────────────────────────────────────────────────────
# 7. ממשק ראשי
# ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-title">🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</div>
<div class="main-subtitle">אופטימיזציה אוטומטית של לוחות שמירה, תורנויות ומשימות — הוגן, מהיר, מדויק</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs([
    "🚀  ביצוע שיבוץ",
    "📖  מדריך מפורט",
    "📥  תבניות עבודה",
])

# ──────────────────── טאב תבניות ────────────────────
with tab_templates:
    st.markdown("### 📥 הורדת תבניות Excel")
    st.markdown(
        '<div class="info-box">הורידו את התבניות, מלאו אותן לפי ההנחיות במדריך, ואז חזרו לטאב <b>ביצוע שיבוץ</b>.</div>',
        unsafe_allow_html=True,
    )
    s_example = pd.DataFrame({
        "מספר_אישי": [1001, 1002, 1003],
        "שם": ["ישראל ישראלי", "יוסי כהן", "דנה לוי (פטור שער)"],
        "פטורים": ["", "", "1"],
    })
    t_example = pd.DataFrame({
        "קוד_משימה": [1, 2, 10],
        "שם": ["שמירת שער", "תורנות מטבח", "כיתת כוננות"],
        "כוח_אדם_נדרש": [1, 3, 8],
        "משך_משמרת": [4, 6, 24],
        "שעות_מנוחה_אחרי": [8, 0, 0],
        "אישור_חפיפה": [False, False, True],
        "שעות_פעילות": ["all", "7,8,9,12,13,14", "all"],
    })

    st.markdown("**👥 תבנית חיילים** — חובה: מספר_אישי ייחודי, שם, פטורים (קודי משימות מופרדים בפסיק)")
    st.dataframe(s_example, use_container_width=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "⬇️ הורד תבנית חיילים",
            data=to_excel_styled(s_example, "Soldiers", False),
            file_name="Shavtzak_Soldiers.xlsx",
            use_container_width=True,
        )
    st.markdown("---")
    st.markdown("**📋 תבנית משימות** — הגדרת כל עמדה ומאפייניה")
    st.dataframe(t_example, use_container_width=True)
    with c2:
        st.download_button(
            "⬇️ הורד תבנית משימות",
            data=to_excel_styled(t_example, "Tasks", False),
            file_name="Shavtzak_Tasks.xlsx",
            use_container_width=True,
        )

# ──────────────────── טאב מדריך ────────────────────
with tab_guide:
    st.markdown("### 📖 מדריך מלא — הסבר לכל עמודה")

    st.markdown("#### 👥 קובץ חיילים — `Soldiers.xlsx`")
    st.markdown("""
    <table class="guide-table">
        <thead><tr><th style="width:20%">עמודה</th><th style="width:50%">הסבר</th><th>הנחיות</th></tr></thead>
        <tbody>
        <tr>
            <td><b>מספר_אישי</b></td>
            <td>מזהה ייחודי לכל חייל (מספר שלם).</td>
            <td>חובה. אין לחזור על מספרים. המערכת מסתמכת עליו להבחנה בין חיילים בעלי שם זהה.</td>
        </tr>
        <tr>
            <td><b>שם</b></td>
            <td>שם החייל כפי שיופיע בלוח השיבוץ.</td>
            <td>מומלץ שם פרטי + משפחה לבהירות מרבית.</td>
        </tr>
        <tr>
            <td><b>פטורים</b></td>
            <td>קודי המשימות שהחייל <b>לא יכול</b> לבצע.</td>
            <td>רשמו את <b>קוד_משימה</b> מהקובץ המשימות (לדוגמה: <code>1</code> לשמירת שער). כמה פטורים — הפרידו בפסיק (<code>1,3,5</code>). השאירו ריק אם אין הגבלה.</td>
        </tr>
        </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("<br>#### 📋 קובץ משימות — `Tasks.xlsx`", unsafe_allow_html=True)
    st.markdown("""
    <table class="guide-table">
        <thead><tr><th style="width:22%">עמודה</th><th style="width:48%">הסבר</th><th>הנחיות</th></tr></thead>
        <tbody>
        <tr>
            <td><b>קוד_משימה</b></td>
            <td>מספר זיהוי ייחודי למשימה.</td>
            <td>חייב להתאים למספרים שנרשמו בעמודת "פטורים" בקובץ החיילים.</td>
        </tr>
        <tr>
            <td><b>שם</b></td>
            <td>שם המשימה (שמירה, סיור, מטבח…).</td>
            <td>שם זה יופיע בתאי לוח השיבוץ.</td>
        </tr>
        <tr>
            <td><b>כוח_אדם_נדרש</b></td>
            <td>מספר החיילים שיש לאייש בכל שעה פעילה.</td>
            <td>שמירה בזוגות = <code>2</code>. האלגוריתם מבטיח בדיוק מספר זה בכל שעה.</td>
        </tr>
        <tr>
            <td><b>משך_משמרת</b></td>
            <td>מספר שעות רצופות שחייל מבצע את המשימה.</td>
            <td><b>נעילת משמרת:</b> חייל שהוקצה לשעה H ישאר שם עד H+N. לא ניתן להחליפו באמצע.</td>
        </tr>
        <tr>
            <td><b>שעות_מנוחה_אחרי</b></td>
            <td>שעות מנוחה מינימליות לאחר סיום המשמרת.</td>
            <td>לאחר משמרת של 4 שעות עם 8 שעות מנוחה, החייל חסום מכל משימה ל-8 השעות הבאות.</td>
        </tr>
        <tr>
            <td><b>אישור_חפיפה</b></td>
            <td>האם ניתן להיות במשימה הזו במקביל למשימה אחרת?</td>
            <td><code>True</code> — כוננות (לא חוסמת). <code>False</code> — שמירה/מטבח (חוסמת לחלוטין).</td>
        </tr>
        <tr>
            <td><b>שעות_פעילות</b></td>
            <td>השעות שבהן המשימה מתקיימת.</td>
            <td><code>all</code> — 24/7. לשעות ספציפיות: <code>7,8,9,12,13,14</code> (בלי רווחים).</td>
        </tr>
        </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    💡 <b>שינת לילה:</b> המערכת מעדיפה אוטומטית לאפשר לכל חייל מנוחת לילה (22:00–08:00) ככל שהסד"כ מאפשר.
    פונקציית המטרה מכבידה פי 50 על פער העומסים ומעדיפה חלוקה אחידה, בנוסף לצמצום עבודת לילה.
    </div>
    """, unsafe_allow_html=True)

# ──────────────────── טאב ביצוע שיבוץ ────────────────────
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים (Soldiers.xlsx)", type="xlsx", key="sf")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות (Tasks.xlsx)", type="xlsx", key="tf")

    # הגדרות מתקדמות
    with st.expander("⚙️ הגדרות מתקדמות"):
        num_hours = st.slider("מספר שעות בלוח", min_value=8, max_value=48, value=24, step=1)
        solver_timeout = st.slider("זמן חיפוש מקסימלי (שניות)", min_value=10, max_value=120, value=30, step=5)

    if sf and tf:
        s_df = pd.read_excel(sf)
        t_df = pd.read_excel(tf)

        # פריסת נתונים
        required_s_cols = {"מספר_אישי", "שם"}
        required_t_cols = {"קוד_משימה", "שם", "כוח_אדם_נדרש"}
        missing_s = required_s_cols - set(s_df.columns)
        missing_t = required_t_cols - set(t_df.columns)

        if missing_s or missing_t:
            if missing_s:
                st.error(f"חסרות עמודות בקובץ חיילים: {missing_s}")
            if missing_t:
                st.error(f"חסרות עמודות בקובץ משימות: {missing_t}")
        else:
            soldiers = [
                Soldier(r["מספר_אישי"], r["שם"], r.get("פטורים", ""))
                for _, r in s_df.iterrows()
            ]
            tasks = [
                Task(
                    r["קוד_משימה"], r["שם"], r["כוח_אדם_נדרש"],
                    r.get("משך_משמרת"), r.get("שעות_מנוחה_אחרי"),
                    r.get("אישור_חפיפה"), r.get("שעות_פעילות"),
                )
                for _, r in t_df.iterrows()
            ]

            # ולידציה
            warnings = validate_inputs(soldiers, tasks)
            for w in warnings:
                st.warning(w)

            hard_errors = [w for w in warnings if w.startswith("❌")]
            if not hard_errors:
                if st.button("⚙️ צור שבצ\"ק ונתח תובנות", use_container_width=True):
                    with st.spinner("🔄 מחשב שיבוץ אופטימלי… (עשוי לקחת עד 30 שניות)"):
                        final_df = solve_scheduling(soldiers, tasks, num_hours, solver_timeout)

                    if final_df is not None:
                        st.success("✅ שיבוץ נוצר בהצלחה!")

                        # מדדים
                        render_summary_metrics(final_df, soldiers, tasks)
                        st.markdown("---")

                        # לוח שיבוץ
                        st.subheader("📅 לוח השיבוץ הסופי")
                        st.table(final_df)
                        st.download_button(
                            "📥 הורד לוח שיבוץ (Excel)",
                            data=to_excel_styled(final_df),
                            file_name="Final_Shavtzak.xlsx",
                            use_container_width=True,
                        )

                        # גרפים
                        st.divider()
                        st.subheader("📊 ניתוח עומסים")
                        g1, g2 = st.columns(2)
                        with g1:
                            render_load_chart(final_df)
                        with g2:
                            render_night_chart(final_df)

                        # הסבר
                        st.divider()
                        st.subheader("💡 למה הלו\"ז נראה ככה?")
                        with st.expander("לחץ לקריאת הניתוח"):
                            all_hours_tasks = [
                                t for t in tasks
                                if set(t.active_hours) & Task.NIGHT_HOURS and not t.allow_overlap
                            ]
                            night_demand = sum(t.required_personnel for t in all_hours_tasks)
                            max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                            min_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].min()]["שם"].tolist()
                            load_gap = final_df["סך שעות"].max() - final_df["סך שעות"].min()

                            st.markdown(f"""
**1. עבודת לילה:** נדרשים {night_demand} חיילים בשעות הלילה.
האלגוריתם מחלק את הלילה בהוגנות ככל האפשר, אך כשאין מספיק אנשים — חלקם נאלצים לוותר על שינה.

**2. חיילים עמוסים:** {', '.join(max_s)} — {final_df['סך שעות'].max()} שעות.
**3. חיילים עם עומס קל:** {', '.join(min_s)} — {final_df['סך שעות'].min()} שעות.
**4. פער עומסים:** {load_gap} שעות. {'✅ מצוין' if load_gap <= 2 else '⚠️ ניתן לשפר על-ידי הגדלת כוח האדם' if load_gap > 4 else '⚠️ סביר'}.

**5. השפעת מנוחות:** כל חייל שהוגדרה לו מנוחה ארוכה "יוצא מהסבב" לשעות אלו, מה שמגדיל את עומס האחרים.
                            """)
                    else:
                        st.error("""
**❌ לא נמצא פתרון חוקי.**

סיבות אפשריות:
- מספר החיילים קטן מדי ביחס לדרישות הכוח האדם
- הגבלות מנוחה נוקשות מדי
- פטורים רבים מדי מצמצמים את האפשרויות

**המלצות:** הגדל את מספר החיילים, קצר את שעות המנוחה, או הפחת את כוח האדם הנדרש.
                        """)
    else:
        st.markdown("""
        <div class="info-box">
        👆 <b>כדי להתחיל:</b> העלו קובץ חיילים וקובץ משימות. אין לכם תבניות? הורידו אותן מטאב <b>תבניות עבודה</b>.
        </div>
        """, unsafe_allow_html=True)
