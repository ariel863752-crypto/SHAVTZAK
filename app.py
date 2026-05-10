import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import base64

# ==========================================
# 1. הגדרות ועיצוב האתר
# ==========================================
st.set_page_config(page_title="מערכת שיבוץ חיילים", page_icon="🪖", layout="wide")

# עיצוב עברית בסיסי
st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] div[role="button"] p { font-size: 1.2rem; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 2. מחלקות הנתונים ופונקציות עזר (אותו הגיון מהקוד הקודם)
# ==========================================
class Soldier:
    def __init__(self, soldier_id, name, qualifications=None, restricted_tasks=None, max_consecutive_hours=6):
        self.soldier_id = soldier_id
        self.name = name
        self.qualifications = qualifications if qualifications is not None else []
        self.restricted_tasks = restricted_tasks if restricted_tasks is not None else []
        self.max_consecutive_hours = max_consecutive_hours


class Task:
    def __init__(self, task_id, name, required_personnel, duration, required_qualifications=None, allow_overlap=False,
                 difficulty=1, active_hours=None):
        self.task_id = task_id
        self.name = name
        self.required_personnel = required_personnel
        self.duration = duration
        self.required_qualifications = required_qualifications if required_qualifications is not None else {}
        self.allow_overlap = allow_overlap
        self.difficulty = difficulty
        self.active_hours = active_hours if active_hours is not None else list(range(25))


def parse_list(val): return [x.strip() for x in str(val).split(',')] if pd.notna(val) else []


def parse_dict(val):
    if pd.isna(val): return {}
    return {item.split(':')[0].strip(): int(item.split(':')[1].strip()) for item in str(val).split(',') if ':' in item}


# ==========================================
# 3. מנוע האופטימיזציה (מותאם לאתר)
# ==========================================
def run_solver(soldiers, tasks, num_hours=25):
    model = cp_model.CpModel()
    x = {}
    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f'x_{s.soldier_id}_{t.task_id}_{h}')

    # אילוצים (אותם אילוצים מהקוד שעבד לנו)
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

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        results = []
        for s in soldiers:
            row = {"שם": s.name}
            for h in range(num_hours):
                label = f"{h:02d}:00" if h < 24 else "00:00 (מחר)"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[label] = " + ".join(active) if active else "-"
            results.append(row)
        return pd.DataFrame(results)
    return None


# ==========================================
# 4. ממשק האתר (Streamlit)
# ==========================================
st.title("🪖 מערכת שיבוץ חיילים חכמה")
st.write("העלה את הקבצים וקבל לו\"ז אופטימלי בשניות")

with st.sidebar:
    st.header("הגדרות")
    num_hours = st.slider("טווח שעות לתכנון", 1, 25, 25)

col1, col2 = st.columns(2)
with col1:
    soldiers_file = st.file_uploader("בחר קובץ חיילים", type="xlsx")
with col2:
    tasks_file = st.file_uploader("בחר קובץ משימות", type="xlsx")

if soldiers_file and tasks_file:
    try:
        s_df = pd.read_excel(soldiers_file)
        t_df = pd.read_excel(tasks_file)

        # המרה לאובייקטים
        soldiers = [Soldier(int(r['מספר_אישי']), r['שם'], parse_list(r.get('הכשרות')), parse_list(r.get('פטורים'))) for
                    _, r in s_df.iterrows()]
        tasks = [Task(int(r['קוד_משימה']), r['שם'], int(r['כוח_אדם_נדרש']), int(r['משך_זמן']),
                      parse_dict(r.get('הכשרות_נדרשות'))) for _, r in t_df.iterrows()]

        if st.button("🚀 חשב שיבוץ אופטימלי"):
            with st.spinner("האלגוריתם מחשב את הפתרון הטוב ביותר..."):
                final_df = run_solver(soldiers, tasks, num_hours)

                if final_df is not None:
                    st.success("✅ השיבוץ הושלם!")
                    st.dataframe(final_df, use_container_width=True)

                    # אפשרות הורדה
                    csv = final_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 הורד לו\"ז כקובץ CSV", csv, "schedule.csv", "text/csv")
                else:
                    st.error("❌ לא נמצא פתרון חוקי. בדוק אם יש מספיק חיילים או התנגשות בשעות.")
    except Exception as e:
        st.error(f"שגיאה בקריאת הקבצים: {e}")