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

    # 1. בדיקת עמודות חובה חיילים
    required_s_cols = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    for col in required_s_cols:
        if col not in s_df.columns:
            errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ חיילים.")

    # 2. בדיקת עמודות חובה משימות
    required_t_cols = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'הכשרות_נדרשות', 'אישור_חפיפה', 'רמת_קושי',
                       'שעות_פעילות']
    for col in required_t_cols:
        if col not in t_df.columns:
            errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ משימות.")

    if errors: return errors  # אם חסרות עמודות, אל תמשיך לבדוק תוכן

    # 3. בדיקת תוכן חיילים
    if s_df['מספר_אישי'].isnull().any():
        errors.append("⚠️ שגיאת נתונים: ישנם חיילים ללא 'מספר_אישי'.")
    if s_df['שם'].isnull().any():
        errors.append("⚠️ שגיאת נתונים: ישנם חיילים ללא 'שם'.")

    # 4. בדיקת תוכן משימות
    if not pd.api.types.is_numeric_dtype(t_df['משך_זמן']):
        errors.append("⚠️ שגיאת נתונים: העמודה 'משך_זמן' חייבת להכיל מספרים בלבד.")
    if not pd.api.types.is_numeric_dtype(t_df['כוח_אדם_נדרש']):
        errors.append("⚠️ שגיאת נתונים: העמודה 'כוח_אדם_נדרש' חייבת להכיל מספרים בלבד.")

    return errors


# ==========================================
# 4. פונקציות עזר ויצירת פורמטים להורדה
# ==========================================
def create_template_files():
    soldiers_df = pd.DataFrame(columns=['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף'])
    soldiers_example = pd.DataFrame([
        {'מספר_אישי': 1234567, 'שם': 'ישראל ישראלי', 'הכשרות': 'נהג,חובש', 'פטורים': '101,102', 'מקסימום_שעות_ברצף': 6},
        {'מספר_אישי': 7654321, 'שם': 'דוד כהן', 'הכשרות': '', 'פטורים': '', 'מקסימום_שעות_ברצף': 4}
    ])

    tasks_df = pd.DataFrame(
        columns=['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'הכשרות_נדרשות', 'אישור_חפיפה', 'רמת_קושי',
                 'שעות_פעילות'])
    tasks_example = pd.DataFrame([
        {'קוד_משימה': 101, 'שם': 'שמירת שער', 'כוח_אדם_נדרש': 2, 'משך_זמן': 4, 'הכשרות_נדרשות': '',
         'אישור_חפיפה': False, 'רמת_קושי': 2, 'שעות_פעילות': 'all'},
        {'קוד_משימה': 102, 'שם': 'סיור רכוב', 'כוח_אדם_נדרש': 1, 'משך_זמן': 2, 'הכשרות_נדרשות': 'נהג:1',
         'אישור_חפיפה': False, 'רמת_קושי': 3, 'שעות_פעילות': '0,1,2,3,4'}
    ])
    return soldiers_example, tasks_example


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='תבנית')
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
                for req_role, req_amount in t.required_qualifications.items():
                    model.Add(sum(
                        x[s.soldier_id, t.task_id, h] for s in soldiers if req_role in s.qualifications) >= req_amount)
            else:
                for s in soldiers:
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    for s in soldiers:
        for h in range(num_hours):
            blocking_tasks = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking_tasks) <= 1)

        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

            for h in range(num_hours):
                starts_in_window = [start[s.soldier_id, t.task_id, h - d] for d in range(t.duration) if h - d >= 0]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(starts_in_window))
                if h > num_hours - t.duration:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        schedule_data = []
        for s in soldiers:
            soldier_row = {"מספר_אישי": s.soldier_id, "שם": s.name}
            for h in range(num_hours):
                label = f"{h:02d}:00" if h < 24 else "00:00 (מחר)"
                active_tasks = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                soldier_row[label] = " + ".join(active_tasks) if active_tasks else "-"
            schedule_data.append(soldier_row)
        return pd.DataFrame(schedule_data)
    return None


# ==========================================
# 6. ממשק האתר - לשוניות (Tabs)
# ==========================================
st.title("🪖 מערכת שיבוץ כוח אדם חכמה")

# יצירת 3 לשוניות ניווט
tab_main, tab_instructions, tab_templates = st.tabs([
    "🚀 עמוד עבודה ושיבוץ",
    "📖 הוראות מילוי למפקד",
    "📥 הורדת תבניות אקסל"
])

# --- לשונית 3: הורדת תבניות ---
with tab_templates:
    st.header("הורדת קבצי עבודה נקיים")
    st.write("אנא השתמש רק בפורמטים אלו כדי למנוע שגיאות במערכת.")
    s_temp, t_temp = create_template_files()
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.download_button("1️⃣ הורד תבנית לחיילים", data=to_excel(s_temp), file_name="Template_Soldiers.xlsx")
    with col_t2:
        st.download_button("2️⃣ הורד תבנית למשימות", data=to_excel(t_temp), file_name="Template_Tasks.xlsx")

# --- לשונית 2: הוראות מילוי ---
with tab_instructions:
    st.header("איך למלא את קבצי האקסל?")

    st.subheader("קובץ חיילים:")
    st.markdown("""
    * **מספר_אישי / שם:** שדות חובה. נא לא להשאיר ריק.
    * **הכשרות:** רשימת מילים מופרדות בפסיק (למשל: `נהג, חובש`). אם אין, השאר ריק.
    * **פטורים:** קודי המשימות שהחייל פטור מהן, מופרדים בפסיק (למשל: `101, 105`).
    * **מקסימום_שעות_ברצף:** מספר השעות הרצופות המקסימלי המותר לחייל זה (ברירת מחדל: 6).
    """)

    st.subheader("קובץ משימות:")
    st.markdown("""
    * **קוד_משימה / שם / כוח_אדם_נדרש / משך_זמן:** חובה למלא. מספרים שלמים.
    * **הכשרות_נדרשות:** דורש פורמט ספציפי - שם הכשרה וכמות מופרדים בנקודתיים (למשל: `נהג:1, קצין:2`).
    * **שעות_פעילות:** מתי המשימה קורית. ניתן לרשום `all` (לכל היממה) או שעות מופרדות בפסיק (למשל `0,1,2,22,23`).
    """)

# --- לשונית 1: העבודה עצמה ---
with tab_main:
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        soldiers_file = st.file_uploader("1️⃣ העלה את קובץ החיילים (Excel)", type="xlsx")
    with col_up2:
        tasks_file = st.file_uploader("2️⃣ העלה את קובץ המשימות (Excel)", type="xlsx")

    if soldiers_file and tasks_file:
        try:
            s_df = pd.read_excel(soldiers_file)
            t_df = pd.read_excel(tasks_file)

            # הפעלת הבודק החכם
            validation_errors = validate_excel_files(s_df, t_df)

            if validation_errors:
                st.error("התגלו שגיאות בקבצים שהועלו! אנא תקן והעלה מחדש:")
                for err in validation_errors:
                    st.warning(err)
            else:
                st.success("✅ הקבצים תקינים לחלוטין!")

                if st.button("⚙️ התחל שיבוץ אוטומטי", type="primary", use_container_width=True):
                    with st.spinner("מחשב את השיבוץ המיטבי (זה עשוי לקחת עד חצי דקה)..."):
                        soldiers_list = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'),
                                                 r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                        tasks_list = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get