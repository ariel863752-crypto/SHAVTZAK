import io
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
import traceback

st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

# ══════════════════════════════════════════════════════════════════
# 1. CSS
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');
html, body { direction: rtl; }
.stApp, [data-testid="stAppViewContainer"], .main, .block-container, .stMarkdown, p, span, li, label, div, [data-testid="stText"], [data-testid="stMarkdownContainer"], [data-testid="stAlert"], [data-testid="stExpander"] summary, [data-testid="stFileUploader"] label, [data-testid="stFileUploader"] div, [data-testid="stSlider"] label, [data-testid="stSelectbox"] label { font-family: 'Heebo', sans-serif !important; direction: rtl !important; text-align: right !important; }
.stApp { background: #f2f5f2; }
.block-container { padding: 2rem 2.5rem 3rem; max-width: 1400px; }
.app-header { background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%); border-radius: 16px; padding: 30px 35px; margin-bottom: 28px; box-shadow: 0 4px 20px rgba(45,90,39,0.25); text-align: right; }
.app-header h1 { font-size: clamp(24px, 4vw, 42px); font-weight: 900; color: white; margin: 0 0 8px 0; letter-spacing: -0.5px; }
.app-header p { font-size: 16px; color: rgba(255,255,255,0.88); margin: 0; }
[data-testid="stTabs"] [data-baseweb="tab-list"] { flex-direction: row-reverse !important; justify-content: flex-start !important; gap: 6px; background: white; border-radius: 12px; padding: 5px; border: 1px solid #dde8dc; margin-bottom: 20px; }
[data-testid="stTabs"] [data-baseweb="tab"] { border-radius: 8px; font-weight: 600; font-size: 14px; padding: 8px 20px; color: #5a7a57; direction: rtl; }
[data-testid="stTabs"] [aria-selected="true"] { background: #2d5a27 !important; color: white !important; }
.metric-row { display: flex; flex-direction: row-reverse; gap: 16px; margin: 22px 0; flex-wrap: wrap; }
.metric-card { flex: 1; min-width: 160px; background: white; border-radius: 14px; padding: 22px; border: 1px solid #dde8dc; box-shadow: 0 2px 8px rgba(45,90,39,0.07); text-align: right; direction: rtl; }
.mc-label { font-size: 11px; color: #7a9a77; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px; }
.mc-value { font-size: 34px; font-weight: 900; color: #1a3d17; line-height: 1; }
.mc-sub   { font-size: 12px; color: #a0b89d; margin-top: 4px; }
div.stButton > button:first-child { background: linear-gradient(135deg, #2d5a27, #3d7a35) !important; color: white !important; font-weight: 700 !important; font-size: 17px !important; border-radius: 10px !important; border: none !important; height: 3.4em; width: 100%; box-shadow: 0 4px 14px rgba(45,90,39,0.3); transition: all 0.18s !important; }
[data-testid="stDownloadButton"] > button { background: #b84d00 !important; color: white !important; font-weight: 600 !important; border-radius: 10px !important; border: none !important; }
[data-testid="stFileUploader"] { background: white; border-radius: 12px; padding: 14px 16px; border: 2px dashed #c0d8bc; direction: rtl; text-align: right; }
[data-testid="stTable"] table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: white; direction: rtl; }
[data-testid="stTable"] th { background: #2d5a27 !important; color: white !important; padding: 10px 13px !important; font-weight: 600 !important; font-size: 12px !important; text-align: right !important; direction: rtl !important; }
[data-testid="stTable"] td { padding: 9px 13px !important; border-bottom: 1px solid #f0f0f0 !important; text-align: right !important; direction: rtl !important; }
.info-box { background: #edf5ec; border-right: 5px solid #2d5a27; padding: 14px 18px; margin: 14px 0; font-size: 14px; color: #1a3d17; line-height: 1.8; direction: rtl; text-align: right; border-radius: 0 10px 10px 0; }
.warn-box { background: #fff8e6; border-right: 5px solid #e67e22; padding: 14px 18px; margin: 14px 0; font-size: 14px; color: #7a4500; line-height: 1.8; direction: rtl; text-align: right; border-radius: 0 10px 10px 0; }
.error-box { background: #fdecea; border-right: 5px solid #c0392b; padding: 14px 18px; margin: 14px 0; font-size: 14px; color: #7a0010; line-height: 1.8; direction: rtl; text-align: right; border-radius: 0 10px 10px 0; }
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
        part = part.strip()
        if '-' in part:
            parts = part.split('-')
            if len(parts) == 2:
                s_str, e_str = parts[0].split(':')[0].strip(), parts[1].split(':')[0].strip()
                if s_str.isdigit() and e_str.isdigit():
                    s, e = int(s_str), int(e_str)
                    if s <= e:
                        for h in range(s, e + 1): res.add(h % 24)
                    else:
                        for h in range(s, 24): res.add(h % 24)
                        for h in range(0, e + 1): res.add(h % 24)
        elif part.replace('.0', '').isdigit():
            res.add(int(float(part)) % 24)
        elif ':' in part:
            h_str = part.split(':')[0].strip()
            if h_str.isdigit(): res.add(int(h_str) % 24)
    return sorted(list(res))

class Soldier:
    def __init__(self, s_id, name, restr="", roles="", unavail=""):
        self.soldier_id = str(s_id).replace('.0', '').strip()
        self.name       = str(name).strip()
        self.restricted_tasks = ([int(float(t)) for t in str(restr).split(',') if str(t).strip().replace('.0', '').isdigit()] if pd.notna(restr) and str(restr).strip() not in ("", "nan") else [])
        self.roles = ([r.strip() for r in str(roles).split(',') if r.strip()] if pd.notna(roles) and str(roles).strip() not in ("", "nan") else [])
        self.unavail_hours = parse_time_ranges(unavail)

class Task:
    def __init__(self, t_id, name, req_p, shift_dur, rest_dur, overlap, hours, req_roles, intensity, blocked_roles=""):
        self.task_id            = int(float(t_id))
        self.name               = str(name).strip()
        self.required_personnel = int(req_p)
        self.shift_duration     = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration      = int(rest_dur)  if pd.notna(rest_dur)  else 0
        self.allow_overlap      = str(overlap).strip().lower() == 'true'
        self.active_hours       = parse_time_ranges(hours)
        self.intensity          = int(intensity) if pd.notna(intensity) else 1
        self.blocked_roles      = ([r.strip() for r in str(blocked_roles).split(',') if r.strip()] if pd.notna(blocked_roles) and str(blocked_roles).strip() not in ("", "nan") else [])
        parsed_roles = ([r.strip() for r in str(req_roles).split(',')] if pd.notna(req_roles) and str(req_roles).strip() not in ("", "nan") else [])
        self.slots = parsed_roles.copy()
        while len(self.slots) < self.required_personnel:
            self.slots.append(None)

def to_excel_styled(df: pd.DataFrame, sheet_name: str = 'שבצ"ק', include_index: bool = True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        wb, ws = writer.book, writer.sheets[sheet_name]
        hfmt = wb.add_format({'bold': True, 'fg_color': '#2d5a27', 'font_color': 'white', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
        efmt = wb.add_format({'fg_color': '#f0f8ef', 'align': 'right'})
        bfmt = wb.add_format({'align': 'right'})
        for col_num, col_val in enumerate(df.columns.values):
            col_idx = col_num + (1 if include_index else 0)
            ws.write(0, col_idx, col_val, hfmt)
            col_len = max(df[col_val].astype(str).map(len).max(), len(col_val)) + 5
            ws.set_column(col_idx, col_idx, min(col_len, 40))
        for row_num in range(1, len(df) + 1):
            ws.set_row(row_num, None, efmt if row_num % 2 == 0 else bfmt)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════
# 3. מנוע CP-SAT v9 — אנושי וגמיש
# ══════════════════════════════════════════════════════════════════
def solve_scheduling(soldiers: list, tasks: list, num_hours: int = 24, time_limit: float = 120.0, soft_rest: bool = True):
    model = cp_model.CpModel()
    x = {}
    SLEEP_WINDOW = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8]
    zero_var = model.NewIntVar(0, 0, 'zero_const')

    # יצירת "חייל חסר" (Dummy) שיכול לקחת כל תפקיד ולעבוד כל שעה
    all_roles = list({r for t in tasks for r in t.slots if r})
    dummy = Soldier(s_id="DUMMY_999", name="⚠️ חייל חסר", roles=",".join(all_roles))
    all_soldiers = soldiers + [dummy]

    # יצירת משתנים
    for s in all_soldiers:
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                for h in range(num_hours):
                    x[s.soldier_id, t.task_id, slot_idx, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}")

    # אילוצים
    for s in all_soldiers:
        if s.soldier_id != dummy.soldier_id:
            for h in s.unavail_hours:
                if h < num_hours:
                    model.Add(sum(x[s.soldier_id, t.task_id, slot, h] for t in tasks for slot in range(len(t.slots))) == 0)

        for t in tasks:
            if s.soldier_id != dummy.soldier_id:
                if any(role in t.blocked_roles for role in s.roles) or (t.task_id in s.restricted_tasks):
                    for slot_idx in range(len(t.slots)):
                        for h in range(num_hours):
                            model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == 0)
                    continue

            for slot_idx, required_role in enumerate(t.slots):
                if s.soldier_id != dummy.soldier_id and required_role is not None and required_role not in s.roles:
                    for h in range(num_hours):
                        model.Add(x[s.soldier_id, t.task_id, slot_idx, h] == 0)
                    continue

    # כיסוי עמדות 100%
    for t in tasks:
        for slot_idx in range(len(t.slots)):
            for h in range(num_hours):
                assigned = sum(x[s.soldier_id, t.task_id, slot_idx, h] for s in all_soldiers)
                if h in t.active_hours:
                    model.Add(assigned == 1)
                else:
                    model.Add(assigned == 0)

    starts_vars = []
    rest_violation_vars = []

    # לוגיקה אנושית (משמרות עד X שעות) לחיילים אמיתיים
    for s in soldiers: 
        for h in range(num_hours):
            blocking = [x[s.soldier_id, t.task_id, slot_idx, h] for t in tasks if not t.allow_overlap for slot_idx in range(len(t.slots))]
            model.Add(sum(blocking) <= 1)
            
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                # משמרת מקסימלית (ברצף)
                if t.shift_duration > 0:
                    for h in range(num_hours):
                        window = [x[s.soldier_id, t.task_id, slot_idx, (h+i)%num_hours] for i in range(t.shift_duration + 1)]
                        model.Add(sum(window) <= t.shift_duration)
                
                for h in range(num_hours):
                    # משתנה התחלת משמרת
                    start_var = model.NewBoolVar(f"st_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}")
                    model.Add(start_var >= x[s.soldier_id, t.task_id, slot_idx, h] - x[s.soldier_id, t.task_id, slot_idx, (h-1)%num_hours])
                    starts_vars.append(start_var)
                    
                    # משתנה סיום משמרת -> תחילת מנוחה
                    if t.rest_duration > 0 and not t.allow_overlap:
                        end_var = model.NewBoolVar(f"en_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}")
                        model.Add(end_var >= x[s.soldier_id, t.task_id, slot_idx, (h-1)%num_hours] - x[s.soldier_id, t.task_id, slot_idx, h])
                        
                        for offset in range(t.rest_duration):
                            rest_h = (h + offset) % num_hours
                            for other_t in tasks:
                                if not other_t.allow_overlap:
                                    for other_slot in range(len(other_t.slots)):
                                        if soft_rest:
                                            viol = model.NewBoolVar(f"rv_{s.soldier_id}_{t.task_id}_{slot_idx}_{h}_{rest_h}_{other_t.task_id}_{other_slot}")
                                            model.Add(viol >= end_var + x[s.soldier_id, other_t.task_id, other_slot, rest_h] - 1)
                                            rest_violation_vars.append(viol)
                                        else:
                                            model.Add(x[s.soldier_id, other_t.task_id, other_slot, rest_h] == 0).OnlyEnforceIf(end_var)

    # פונקציית מטרה
    s_total_hours, s_intensity_scores, sleep_penalties = [], [], []

    for s in soldiers:
        total_h = sum(x[s.soldier_id, t.task_id, slot, h] for t in tasks for slot in range(len(t.slots)) for h in range(num_hours))
        s_total_hours.append(total_h)

        intensity_score = sum(x[s.soldier_id, t.task_id, slot, h] * t.intensity for t in tasks for slot in range(len(t.slots)) for h in range(num_hours))
        s_intensity_scores.append(intensity_score)

        night_work = sum(x[s.soldier_id, t.task_id, slot, h] for t in tasks if not t.allow_overlap for slot in range(len(t.slots)) for h in SLEEP_WINDOW)
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

    dummy_work = sum(x[dummy.soldier_id, t.task_id, slot, h] for t in tasks for slot in range(len(t.slots)) for h in range(num_hours))
    total_rest_violations = sum(rest_violation_vars) if rest_violation_vars else model.NewIntVar(0,0,'no_rv')
    
    # אתחול לגישה מהירה מאוד למציאת פתרון: ניתן ל"חייל חסר" לאייש הכל ואז נתחיל לייעל
    for s in all_soldiers:
        for t in tasks:
            for slot_idx in range(len(t.slots)):
                for h in range(num_hours):
                    if s.soldier_id == dummy.soldier_id:
                        model.AddHint(x[s.soldier_id, t.task_id, slot_idx, h], 1 if h in t.active_hours else 0)
                    else:
                        model.AddHint(x[s.soldier_id, t.task_id, slot_idx, h], 0)

    model.Minimize(
        100000 * dummy_work          # עדיפות עליונה: לא להשתמש בחייל חסר
        + 500 * total_rest_violations # עדיפות שנייה: לא להפר מנוחות
        + 10 * sum(starts_vars)       # עדיפות שלישית: לא לפצל משמרות (מאלץ רצפים)
        + 100 * load_diff
        + 50  * int_diff
        + 200 * sum(sleep_penalties)
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers  = 4
    solver.parameters.log_search_progress = False
    
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, 0, 0

    hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
    rows = []
    
    dummy_usage_count = int(solver.Value(dummy_work))
    total_rest_viol_count = int(solver.Value(total_rest_violations))

    for s in all_soldiers:
        if s.name == "nan" or not s.name: continue
        
        row_hours = 0
        active_list = []
        for h in range(num_hours):
            active = [t.name for t in tasks for slot_idx in range(len(t.slots)) if solver.Value(x[s.soldier_id, t.task_id, slot_idx, h]) == 1]
            active_list.append(" + ".join(active) if active else "—")
            if active: row_hours += 1
            
        # להסתיר את החייל החסר אם לא היה בו צורך
        if s.soldier_id == dummy.soldier_id and row_hours == 0:
            continue
            
        row = {"שם": s.name}
        for h in range(num_hours):
            row[hour_labels[h]] = active_list[h]
            
        row["סך שעות"] = row_hours
        if s.soldier_id != dummy.soldier_id:
            night_count = sum(solver.Value(x[s.soldier_id, t.task_id, slot, h]) for t in tasks if not t.allow_overlap for slot in range(len(t.slots)) for h in SLEEP_WINDOW)
            row["מדד עצימות"] = sum(solver.Value(x[s.soldier_id, t.task_id, slot, h]) * t.intensity for t in tasks for slot in range(len(t.slots)) for h in range(num_hours))
            row["שעות שינה (22-08)"] = len(SLEEP_WINDOW) - night_count
        else:
            row["מדד עצימות"] = 0
            row["שעות שינה (22-08)"] = 11

        rows.append(row)

    df = pd.DataFrame(rows)
    df.index = range(1, len(df) + 1)
    
    # הבאת "החייל החסר" לשורה הראשונה להבלטה
    if dummy_usage_count > 0:
        dummy_row = df[df["שם"] == dummy.name]
        other_rows = df[df["שם"] != dummy.name]
        df = pd.concat([dummy_row, other_rows]).reset_index(drop=True)
        df.index = range(1, len(df) + 1)

    return df, total_rest_viol_count, dummy_usage_count


# ══════════════════════════════════════════════════════════════════
# 4. ממשק ראשי
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה (v9)</h1>
  <p>אופטימיזציה אנושית · משמרות גמישות · רשת ביטחון מפני קריסות סד"כ</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀  ביצוע שיבוץ", "📖  מדריך מפורט", "📥  תבניות אקסל"])

with tab_templates:
    st.markdown("### 📥 הורדת תבניות עבודה")
    s_ex = pd.DataFrame({'מספר אישי': [1001, 1002], 'שם מלא': ['ישראל ישראלי', 'יוסי כהן'], 'פטורים': ['', '101'], 'הסמכות': ['נהג', 'קצין'], 'שעות חסימה': ['', '10-14']})
    t_ex = pd.DataFrame({'מס"ד משימה': [101, 102], 'שם המשימה': ['סיור', 'שמירה'], 'סד"כ נדרש למשימה': [1, 2], 'משך משמרת': [4, 4], 'שעות מנוחה בין משימות': [8, 4], 'אישור חפיפה בין משימות': [False, False], 'שעות פעילות': ['all', '05:30-07:30'], 'הסמכה נדרשת': ['', ''], 'דירוג עצימות משימה (1-3)': [3, 1], 'תפקידים חסומים': ['', '']})
    c1, c2 = st.columns(2)
    with c1: st.download_button("⬇️ הורד תבנית חיילים", data=to_excel_styled(s_ex, "Soldiers", False), file_name="Soldiers_v9.xlsx")
    with c2: st.download_button("⬇️ הורד תבנית משימות", data=to_excel_styled(t_ex, "Tasks", False), file_name="Tasks_v9.xlsx")

with tab_guide:
    st.markdown("""
    ### 📖 מדריך v9 - אופטימיזציה אנושית
    **מה חדש?**
    1. **משמרות גמישות:** "משך המשמרת" הוא כעת המקסימום. המערכת יכולה לקצר משמרת (למשל שעתיים במקום 4) אם העמדה נסגרת, אבל לא תפצל סתם (כמו אדם אנושי).
    2. **חייל רפאים:** אם שעות החסימה של החיילים לא מאפשרות לסגור עמדה (למשל הס/מ"פ חסום בחצות), המערכת **לא תקרוס יותר**. היא תשבץ במקומו "⚠️ חייל חסר", והמפקד יידע שהוא צריך לפתור את זה נקודתית בשטח.
    """)

with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1: sf = st.file_uploader("📂 קובץ חיילים (xlsx)", type="xlsx", key="sf")
    with col_u2: tf = st.file_uploader("📂 קובץ משימות (xlsx)", type="xlsx", key="tf")

    with st.expander("⚙️ הגדרות מתקדמות"):
        time_limit = st.slider("זמן מקסימלי לפתרון (שניות)", 30, 300, 120, 15)
        soft_rest  = st.toggle("אילוצי מנוחה רכים (מומלץ למניעת כשלונות)", value=True)

    if sf and tf:
        if st.button('⚙️ צור שבצ"ק חכם (v9)', key="run_btn"):
            try:
                with st.spinner("מנקה נתוני אקסל ובונה מטריצות..."):
                    s_df = pd.read_excel(sf)
                    t_df = pd.read_excel(tf)

                    s_df_clean = s_df.dropna(subset=[c for c in s_df.columns if 'מספר' in c or 'אישי' in c][:1] + [c for c in s_df.columns if 'שם' in c][:1])
                    t_df_clean = t_df.dropna(subset=[c for c in t_df.columns if 'מס"ד' in c or 'משימה' in c][:1])

                    id_col = [c for c in s_df_clean.columns if 'מספר' in c or 'אישי' in c][0]
                    name_col = [c for c in s_df_clean.columns if 'שם' in c][0]
                    t_id_col = [c for c in t_df_clean.columns if 'מס"ד' in c or 'משימה' in c][0]
                    t_name_col = [c for c in t_df_clean.columns if 'שם' in c or 'משימה' in c][1]
                    t_req_col = [c for c in t_df_clean.columns if 'סד"כ' in c or 'נדרש' in c][0]

                    soldiers = [Soldier(r[id_col], r[name_col], r.get('פטורים', ''), r.get('הסמכות', ''), r.get('שעות חסימה', '')) for _, r in s_df_clean.iterrows()]
                    tasks = [Task(r[t_id_col], r[t_name_col], r[t_req_col], r.get('משך משמרת'), r.get('שעות מנוחה בין משימות'), r.get('אישור חפיפה בין משימות'), r.get('שעות פעילות'), r.get('הסמכה נדרשת', ''), r.get('דירוג עצימות משימה (1-3)', r.get('דירוג עצימות המשימה', 1)), r.get('תפקידים חסומים', '')) for _, r in t_df_clean.iterrows()]

                with st.spinner("מריץ אופטימיזציה בענן..."):
                    final_df, rest_viols, dummy_usage = solve_scheduling(soldiers, tasks, time_limit=time_limit, soft_rest=soft_rest)

                if final_df is not None:
                    gap_h     = int(final_df["סך שעות"].max() - final_df["סך שעות"].min())
                    avg_h     = final_df["סך שעות"].mean()
                    badge     = "✅ מצוין" if gap_h <= 2 else ("⚠️ סביר" if gap_h <= 5 else "❗ גבוה")

                    if dummy_usage > 0:
                        st.markdown(f'<div class="error-box">🚨 <b>סד"כ חסר!</b> המערכת לא מצאה חיילים זמינים כדי לכסות {dummy_usage} שעות (עקב שעות חסימה או חוסר בהסמכות). השעות סומנו בטבלה כ"⚠️ חייל חסר" כדי שתשלימו ידנית.</div>', unsafe_allow_html=True)
                    elif rest_viols > 0:
                        st.markdown(f'<div class="warn-box">⚠️ <b>שימו לב:</b> בוצעו {rest_viols} חריגות ממנוחת חובה כדי לא להשאיר עמדות ריקות.</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="info-box">✅ שיבוץ מושלם! כל העמדות מלאות, ללא חריגות מנוחה.</div>', unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="metric-row">
                      <div class="metric-card"><div class="mc-label">חיילים פעילים</div><div class="mc-value">{len(final_df) - (1 if dummy_usage > 0 else 0)}</div></div>
                      <div class="metric-card"><div class="mc-label">ממוצע שעות</div><div class="mc-value">{avg_h:.1f}</div></div>
                      <div class="metric-card"><div class="mc-label">הפרות מנוחה</div><div class="mc-value">{rest_viols}</div></div>
                      <div class="metric-card"><div class="mc-label">שעות חסרות איוש</div><div class="mc-value">{dummy_usage}</div></div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.subheader("📅 לוח השיבוץ הסופי")
                    st.table(final_df)
                    st.download_button("📥 הורד לוח שיבוץ (Excel)", data=to_excel_styled(final_df), file_name="Final_Shavtzak_v9.xlsx")
                else:
                    st.markdown('<div class="error-box">❌ לא נמצא פתרון. נסו להגדיל את זמן הפתרון בהגדרות.</div>', unsafe_allow_html=True)
            except Exception as inner_error:
                st.error("🚨 תקלה קרתה במהלך עיבוד הנתונים!")
                st.code(traceback.format_exc())
