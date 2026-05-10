import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# ==========================================
# 1. עיצוב ו-RTL
# ==========================================
st.set_page_config(page_title="מערכת שיבוץ חיילים", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    .stAlert { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] { direction: rtl; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th { background-color: #f0f2f6; text-align: right; padding: 12px; border: 1px solid #ddd; }
    td { padding: 10px; border: 1px solid #ddd; text-align: right; }
    .guide-table { width: 100%; background-color: #f8f9fa; border: 1px solid #dee2e6; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. מחלקות נתונים
# ==========================================
class Soldier:
    def __init__(self, soldier_id, name, qualifications=None, restricted_tasks=None, max_consecutive_hours=6):
        self.soldier_id = str(soldier_id)
        self.name = name
        self.qualifications = [str(q).strip() for q in str(qualifications).split(',')] if pd.notna(qualifications) and str(qualifications).strip() != "" else []
        self.restricted_tasks = [int(float(t)) for t in str(restricted_tasks).split(',') if str(t).strip().replace('.0','').isdigit()] if pd.notna(restricted_tasks) else []
        self.max_consecutive_hours = int(max_consecutive_hours) if pd.notna(max_consecutive_hours) and str(max_consecutive_hours).isdigit() else 6

class Task:
    def __init__(self, task_id, name, required_personnel, duration, allow_overlap=False, active_hours=None):
        self.task_id = int(task_id)
        self.name = name
        self.required_personnel = int(required_personnel)
        self.duration = int(duration)
        if isinstance(allow_overlap, str):
            self.allow_overlap = allow_overlap.strip().lower() == 'true'
        else:
            self.allow_overlap = bool(allow_overlap) if pd.notna(allow_overlap) else False
        
        if pd.isna(active_hours) or str(active_hours).strip().lower() in ['all', '']:
            self.active_hours = list(range(25))
        else:
            self.active_hours = [int(x.strip()) for x in str(active_hours).split(',') if str(x).strip().isdigit()]

# ==========================================
# 3. פונקציות עזר (Excel)
# ==========================================
def to_excel_output(df, sheet_name='Schedule'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            column_len = max(df[value].astype(str).map(len).max(), len(value)) + 2
            worksheet.set_column(col_num, col_num, column_len)
    return output.getvalue()

def validate_input(s_df, t_df):
    errors = []
    req_s = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    req_t = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'אישור_חפיפה', 'שעות_פעילות']
    for col in req_s:
        if col not in s_df.columns: errors.append(f"❌ חסרה עמודה בקובץ חיילים: {col}")
    for col in req_t:
        if col not in t_df.columns: errors.append(f"❌ חסרה עמודה בקובץ משימות: {col}")
    if not errors and s_df['מספר_אישי'].duplicated().any():
        errors.append("❌ קיימים מספרי אישי כפולים בקובץ החיילים.")
    return errors

# ==========================================
# 4. האלגוריתם עם חלוקת עומס (MinMax)
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=25):
    model = cp_model.CpModel()
    x = {}
    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")

    # אילוץ כיסוי משימות
    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers: model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אילוץ חייל אחד למשימה אחת (למעט חפיפות מאושרות)
    for s in soldiers:
        for h in range(num_hours):
            blocking_tasks = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking_tasks) <= 1)
        
        # אילוץ פטורים
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # חלוקת עומס שוויונית - מזעור העומס המקסימלי
    soldier_total_hours = []
    for s in soldiers:
        total = sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours))
        soldier_total_hours.append(total)
    
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in soldier_total_hours:
        model.Add(max_load >= load)
    
    model.Minimize(max_load)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for s in soldiers:
            row = {"שם": s.name}
            total_work = 0
            for h in range(num_hours):
                lbl = f"{h:02d}:00"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[lbl] = " + ".join(active) if active else "-"
                if active: total_work += 1
            row["סך שעות"] = total_work
            res.append(row)
        return pd.DataFrame(res)
    return None

# ==========================================
# 5. ממשק משתמש (UI)
# ==========================================
st.title("🪖 מערכת שיבוץ כוח אדם מקצועית")

tab1, tab2, tab3 = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למשתמש", "📥 תבניות אקסל"])

with tab3:
    st.header("הורדת תבניות עבודה")
    s_ex = pd.DataFrame([{'מספר_אישי': 123, 'שם': 'חייל דוגמה', 'הכשרות': 'נהג', 'פטורים': '102', 'מקסימום_שעות_ברצף': 6}])
    t_ex = pd.DataFrame([
        {'קוד_משימה': 101, 'שם': 'שער', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'},
        {'קוד_משימה': 102, 'שם': 'כוננות', 'כוח_אדם_נדרש': 8, 'משך_זמן': 25, 'אישור_חפיפה': True, 'שעות_פעילות': 'all'}
    ])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 תבנית חיילים", data=to_excel_output(s_ex), file_name="Soldiers.xlsx")
    with c2: st.download_button("📥 תבנית משימות", data=to_excel_output(t_ex), file_name="Tasks.xlsx")

with tab2:
    st.header("📖 מדריך מפורט")
    st.markdown("""
    ### כללים למילוי הקבצים:
    1. **אישור_חפיפה:** אם המשימה יכולה לקרות במקביל לאחרות (כמו כיתת כוננות), רשמו **True**.
    2. **פטורים:** רשמו את מספר קוד המשימה. אם יש כמה, הפרידו בפסיק (לדוגמה: `101, 102`).
    3. **שעות_פעילות:** רשמו `all` לפעילות מסביב לשעון, או רשימת שעות (לדוגמה: `8,9,10`).
    4. **מקסימום שעות:** המערכת מחלקת את העומס שווה בשווה באופן אוטומטי.
    """)

with tab1:
    st.subheader("העלאה והרצה")
    u1, u2 = st.columns(2)
    with u1: sf = st.file_uploader("📂 קובץ חיילים", type="xlsx")
    with u2: tf = st.file_uploader("📂 קובץ משימות", type="xlsx")

    if sf and tf:
        s_df = pd.read_excel(sf)
        t_df = pd.read_excel(tf)
        errs = validate_input(s_df, t_df)
        if errs:
            for e in errs: st.error(e)
        else:
            st.success("✅ הקבצים תקינים")
            if st.button("⚙️ צור שיבוץ אופטימלי", type="primary", use_container_width=True):
                with st.spinner("מחשב חלוקת עומס הוגנת..."):
                    s_l = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                    t_l = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                    final = solve_scheduling(s_l, t_l)
                    if final is not None:
                        st.balloons()
                        st.dataframe(final, use_container_width=True)
                        st.download_button("📥 הורד לוח סופי (Excel)", data=to_excel_output(final), file_name="Final_Schedule.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    else:
                        st.error("לא נמצא פתרון חוקי. נסה להוסיף חיילים או להפחית משימות.")
