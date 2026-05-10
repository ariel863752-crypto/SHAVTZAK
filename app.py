import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# ==========================================
# 1. הגדרות ועיצוב האתר
# ==========================================
st.set_page_config(page_title="מערכת שיבוץ חיילים", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    .stAlert { text-align: right; }
    div[data-testid="stExpander"] { direction: rtl; }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 2. מחלקות הנתונים (OOP)
# ==========================================
class Soldier:
    def __init__(self, soldier_id, name, qualifications=None, restricted_tasks=None, max_consecutive_hours=6,
                 total_planning_hours=25):
        self.soldier_id = soldier_id
        self.name = name
        self.qualifications = [str(q).strip() for q in str(qualifications).split(',')] if pd.notna(
            qualifications) else []
        self.restricted_tasks = [int(t.strip()) for t in str(restricted_tasks).split(',') if
                                 str(t).strip().isdigit()] if pd.notna(restricted_tasks) else []
        self.max_consecutive_hours = int(max_consecutive_hours) if pd.notna(max_consecutive_hours) else 6
        self.availability = {t: 1 for t in range(total_planning_hours)}


class Task:
    def __init__(self, task_id, name, required_personnel, duration, required_qualifications=None, allow_overlap=False,
                 difficulty=1, active_hours=None, total_planning_hours=25):
        self.task_id = task_id
        self.name = name
        self.required_personnel = int(required_personnel)
        self.duration = int(duration)

        self.required_qualifications = {}
        if pd.notna(required_qualifications):
            for item in str(required_qualifications).split(','):
                if ':' in item:
                    k, v = item.split(':')
                    self.required_qualifications[k.strip()] = int(v.strip())

        self.allow_overlap = bool(allow_overlap) if pd.notna(allow_overlap) else False
        self.difficulty = int(difficulty) if pd.notna(difficulty) else 1

        if pd.isna(active_hours) or str(active_hours).strip().lower() == 'all':
            self.active_hours = list(range(total_planning_hours))
        else:
            raw_hours = [int(x.strip()) for x in str(active_hours).split(',')]
            self.active_hours = [h if h < total_planning_hours else 0 for h in raw_hours]


# ==========================================
# 3. בקרת שגיאות קפדנית (Deep Validation)
# ==========================================
def validate_excel_files(s_df, t_df):
    errors = []
    required_s_cols = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    for col in required_s_cols:
        if col not in s_df.columns:
            errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ חיילים.")

    required_t_cols = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'הכשרות_נדרשות', 'אישור_חפיפה', 'רמת_קושי',
                       'שעות_פעילות']
    for col in required_t_cols:
        if col not in t_df.columns:
            errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ משימות.")

    if errors: return errors

    if s_df['מספר_אישי'].isnull().any():
        errors.append("⚠️ שגיאת נתונים: ישנם חיילים ללא 'מספר_אישי'.")

    for col in ['משך_זמן', 'כוח_אדם_נדרש']:
        if not pd.api.types.is_numeric_dtype(t_df[col]):
            errors.append(f"⚠️ שגיאת נתונים: העמודה '{col}' חייבת להכיל מספרים בלבד.")

    return errors


# ==========================================
# 4. פונקציות עזר ליצירת תבניות
# ==========================================
def create_template_files():
    s_example = pd.DataFrame([
        {'מספר_אישי': 1234567, 'שם': 'ישראל ישראלי', 'הכשרות': 'נהג,חובש', 'פטורים': '101', 'מקסימום_שעות_ברצף': 6}
    ])
    t_example = pd.DataFrame([
        {'קוד_משימה': 101, 'שם': 'שמירת שער', 'כוח_אדם_נדרש': 2, 'משך_זמן': 4, 'הכשרות_נדרשות': 'נהג:1',
         'אישור_חפיפה': False, 'רמת_קושי': 2, 'שעות_פעילות': 'all'}
    ])
    return s_example, t_example


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Template')
    return output.getvalue()


# ==========================================
# 5. אלגוריתם האופטימיזציה
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=25):
    model = cp_model.CpModel()
    x = {}
    start = {}

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"start_{s.soldier_id}_{t.task_id}_{h}")

    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers:
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

    for s in soldiers:
        for h in range(num_hours):
            model.Add(sum(x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap) <= 1)

        for t in tasks:
            for h in range(num_hours):
                starts_in_window = [start[s.soldier_id, t.task_id, h - d] for d in range(t.duration) if h - d >= 0]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(starts_in_window))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for s in soldiers:
            row = {"שם": s.name}
            for h in range(num_hours):
                lbl = f"{h:02d}:00" if h < 24 else "00:00 (מחר)"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[lbl] = " + ".join(active) if active else "-"
            res.append(row)
        return pd.DataFrame(res)
    return None


# ==========================================
# 6. ממשק האתר
# ==========================================
st.title("🪖 מערכת שיבוץ כוח אדם חכמה")

tab_main, tab_instructions, tab_templates = st.tabs(["🚀 שיבוץ", "📖 הוראות", "📥 תבניות"])

with tab_templates:
    s_temp, t_temp = create_template_files()
    st.download_button("📥 תבנית חיילים", data=to_excel(s_temp), file_name="Soldiers.xlsx")
    st.download_button("📥 תבנית משימות", data=to_excel(t_temp), file_name="Tasks.xlsx")

with tab_instructions:
    st.write("ודא שכל עמודות החובה קיימות ושהערכים תקינים לפי הפורמט.")

with tab_main:
    col1, col2 = st.columns(2)
    with col1:
        s_file = st.file_uploader("קובץ חיילים", type="xlsx")
    with col2:
        t_file = st.file_uploader("קובץ משימות", type="xlsx")

    if s_file and t_file:
        s_df = pd.read_excel(s_file)
        t_df = pd.read_excel(t_file)
        v_errors = validate_excel_files(s_df, t_df)

        if v_errors:
            for e in v_errors: st.error(e)
        else:
            if st.button("⚙️ בצע שיבוץ", type="primary", use_container_width=True):
                with st.spinner("מחשב..."):
                    s_list = [
                        Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף'))
                        for _, r in s_df.iterrows()]
                    t_list = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('הכשרות_נדרשות'),
                                   r.get('אישור_חפיפה'), r.get('רמת_קושי'), r.get('שעות_פעילות')) for _, r in
                              t_df.iterrows()]

                    final_df = solve_scheduling(s_list, t_list)
                    if final_df is not None:
                        st.dataframe(final_df)
                    else:
                        st.error("לא נמצא פתרון.")