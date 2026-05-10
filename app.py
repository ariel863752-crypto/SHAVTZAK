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
    def __init__(self, soldier_id, name, qualifications=None, restricted_tasks=None, max_consecutive_hours=6):
        self.soldier_id = soldier_id
        self.name = name
        self.qualifications = [str(q).strip() for q in str(qualifications).split(',')] if pd.notna(qualifications) else []
        self.restricted_tasks = [int(t.strip()) for t in str(restricted_tasks).split(',') if str(t).strip().isdigit()] if pd.notna(restricted_tasks) else []
        self.max_consecutive_hours = int(max_consecutive_hours) if pd.notna(max_consecutive_hours) else 6

class Task:
    def __init__(self, task_id, name, required_personnel, duration, required_qualifications=None, allow_overlap=False, active_hours=None):
        self.task_id = task_id
        self.name = name
        self.required_personnel = int(required_personnel)
        self.duration = int(duration)
        self.allow_overlap = bool(allow_overlap) if pd.notna(allow_overlap) else False
        
        if pd.isna(active_hours) or str(active_hours).strip().lower() == 'all':
            self.active_hours = list(range(25))
        else:
            self.active_hours = [int(x.strip()) for x in str(active_hours).split(',') if str(x).strip().isdigit()]

# ==========================================
# 3. פונקציית עזר ליצירת קבצי אקסל אמיתיים
# ==========================================
def to_excel_real(df):
    output = io.BytesIO()
    # שימוש ב-xlsxwriter כדי להבטיח קובץ אקסל תקין
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='נתונים')
        # התאמת רוחב עמודות אוטומטית
        worksheet = writer.sheets['נתונים']
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    return output.getvalue()

# ==========================================
# 4. בקרת שגיאות
# ==========================================
def validate_excel_files(s_df, t_df):
    errors = []
    required_s = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    required_t = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'הכשרות_נדרשות', 'אישור_חפיפה', 'רמת_קושי', 'שעות_פעילות']
    
    for col in required_s:
        if col not in s_df.columns: errors.append(f"❌ חסרה עמודה בקובץ חיילים: {col}")
    for col in required_t:
        if col not in t_df.columns: errors.append(f"❌ חסרה עמודה בקובץ משימות: {col}")
    return errors

# ==========================================
# 5. אלגוריתם
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=25):
    model = cp_model.CpModel()
    x = {}
    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")

    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers: model.Add(x[s.soldier_id, t.task_id, h] == 0)

    for s in soldiers:
        for h in range(num_hours):
            model.Add(sum(x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap) <= 1)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 15.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for s in soldiers:
            row = {"שם": s.name}
            for h in range(num_hours):
                lbl = f"{h:02d}:00"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[lbl] = active[0] if active else "-"
            res.append(row)
        return pd.DataFrame(res)
    return None

# ==========================================
# 6. ממשק משתמש
# ==========================================
st.title("🪖 מערכת שיבוץ כוח אדם מקצועית")

tab1, tab2, tab3 = st.tabs(["🚀 עבודה", "📖 הוראות", "📥 תבניות להורדה"])

with tab3:
    st.subheader("הורד תבניות אקסל נקיות")
    s_ex = pd.DataFrame([{'מספר_אישי': 1234567, 'שם': 'ישראל ישראלי', 'הכשרות': 'נהג', 'פטורים': '', 'מקסימום_שעות_ברצף': 6}])
    t_ex = pd.DataFrame([{'קוד_משימה': 101, 'שם': 'אבטחה', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'הכשרות_נדרשות': '', 'אישור_חפיפה': False, 'רמת_קושי': 1, 'שעות_פעילות': 'all'}])
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button("📥 הורד תבנית חיילים (Excel)", data=to_excel_real(s_ex), file_name="Soldiers_Template.xlsx")
    with col_d2:
        st.download_button("📥 הורד תבנית משימות (Excel)", data=to_excel_real(t_ex), file_name="Tasks_Template.xlsx")

with tab2:
    st.info("מלא את האקסלים בדיוק לפי העמודות בתבנית.")

with tab1:
    up1, up2 = st.columns(2)
    with up1: s_file = st.file_uploader("העלה חיילים", type="xlsx")
    with up2: t_file = st.file_uploader("העלה משימות", type="xlsx")

    if s_file and t_file:
        s_df = pd.read_excel(s_file)
        t_df = pd.read_excel(t_file)
        errors = validate_excel_files(s_df, t_df)
        if errors:
            for e in errors: st.error(e)
        else:
            if st.button("⚙️ בצע שיבוץ", type="primary"):
                with st.spinner("מחשב..."):
                    s_list = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                    t_list = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('הכשרות_נדרשות'), r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                    final = solve_scheduling(s_list, t_list)
                    if final is not None:
                        st.balloons()
                        st.dataframe(final)
                    else:
                        st.error("לא נמצא פתרון.")
