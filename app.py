import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. עיצוב ממשק RTL וסטיילינג (שבצ"ק)
# ==========================================
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #fcfcfc; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    .main-title { color: #2d5a27; font-size: 45px; font-weight: bold; border-bottom: 4px solid #e67e22; padding-bottom: 10px; margin-bottom: 25px; }
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; border: 1px solid #ddd; }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 16px; }
    div.stButton > button:first-child {
        background-color: #e67e22 !important;
        color: white !important;
        font-weight: bold !important;
        width: 100%;
        height: 3.5em;
        border-radius: 10px !important;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. מחלקות הנתונים (Objects)
# ==========================================
class Soldier:
    def __init__(self, s_id, name, restr=""):
        self.soldier_id = str(s_id)
        self.name = name
        self.restricted_tasks = [int(float(t)) for t in str(restr).split(',') if str(t).strip().replace('.0','').isdigit()] if pd.notna(restr) else []

class Task:
    def __init__(self, t_id, name, req_p, shift_dur, rest_dur=0, overlap=False, hours="all"):
        self.task_id = int(t_id)
        self.name = name
        self.required_personnel = int(req_p)
        self.shift_duration = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration = int(rest_dur) if pd.notna(rest_dur) else 0
        self.allow_overlap = str(overlap).lower() == 'true'
        if pd.isna(hours) or str(hours).lower() in ['all', '']:
            self.active_hours = list(range(24))
        else:
            self.active_hours = [int(x.strip()) for x in str(hours).split(',') if str(x).strip().isdigit()]

# ==========================================
# 3. המנוע המתמטי - שינה, מנוחה, רציפות והוגנות
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=24):
    model = cp_model.CpModel()
    x = {}      
    start = {}  

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"start_{s.soldier_id}_{t.task_id}_{h}")

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                # אילוץ רציפות המשמרת
                relevant_starts = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))
                
                # אילוץ מנוחה לאחר משימה מעייפת
                if t.rest_duration > 0:
                    total_busy = t.shift_duration + t.rest_duration
                    for next_h in range(h + 1, min(h + total_busy, num_hours)):
                        for other_t in tasks:
                            if not other_t.allow_overlap:
                                model.AddImplication(start[s.soldier_id, t.task_id, h], x[s.soldier_id, other_t.task_id, next_h].Not())

                if h + t.shift_duration > num_hours:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers: model.Add(x[s.soldier_id, t.task_id, h] == 0)

    for s in soldiers:
        for h in range(num_hours):
            blocking = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking) <= 1)
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אופטימיזציה: הוגנות + שינה (22:00 עד 08:00)
    night_hours = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7]
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_loads: model.Add(max_load >= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in night_hours)
    model.Minimize(30 * max_load + night_work)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        hour_labels = [f"{h:02d}:00" for h in range(num_hours)]
        res_list = []
        for s in soldiers:
            row = {"שם": s.name}
            n_work = 0
            for h in range(num_hours):
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[hour_labels[h]] = " + ".join(active) if active else "-"
                if h in night_hours and any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks if not t.allow_overlap):
                    n_work += 1
            row["סך שעות"] = sum(1 for h in range(num_hours) if any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks))
            row["שעות לילה"] = n_work
            res_list.append(row)
        df = pd.DataFrame(res_list)
        df.index = range(1, len(df) + 1)
        return df
    return None

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=True, sheet_name='שבצ"ק')
    return output.getvalue()

# ==========================================
# 4. ממשק המשתמש (Tabs & Insights)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת שיבוץ וניתוח נתונים</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות אקסל"])

with tab_templates:
    st.subheader("הורדת תבניות אקסל")
    s_tmp = pd.DataFrame([{'מספר_אישי': 1001, 'שם': 'חייל א', 'פטורים': ''}])
    t_tmp = pd.DataFrame([{'קוד_משימה': 1, 'שם': 'שמירה', 'כוח_אדם_נדרש': 1, 'משך_משמרת': 4, 'שעות_מנוחה_אחרי': 8, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'}])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel(s_tmp), file_name="Template_Soldiers.xlsx")
    with c2: st.download_button("📥 הור
