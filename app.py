import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# ==========================================
# 1. הגדרות ועיצוב RTL (מימין לשמאל)
# ==========================================
st.set_page_config(page_title="מערכת שיבוץ חיילים", page_icon="🪖", layout="wide")

# הזרקת CSS לתיקון העברית והעיצוב
st.markdown("""
    <style>
    /* הגדרת כיוון כללי מימין לשמאל */
    .stApp, div[data-testid="stSidebar"], .stMarkdown, .stTable, .stDataFrame {
        direction: rtl;
        text-align: right;
    }
    /* תיקון רשימות ותפריטים */
    ul, ol {
        direction: rtl;
        padding-right: 2rem;
        padding-left: 0;
    }
    /* עיצוב טבלאות מדריך */
    .guide-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        background-color: #f8f9fa;
    }
    .guide-table th {
        background-color: #1f77b4;
        color: white;
        padding: 12px;
        text-align: right;
    }
    .guide-table td {
        padding: 10px;
        border: 1px solid #dee2e6;
    }
    /* התראות מעוצבות */
    .stAlert {
        direction: rtl;
        text-align: right;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. מחלקות הנתונים (OOP)
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
def to_excel_file(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    return output.getvalue()

def validate_data(s_df, t_df):
    errors = []
    req_s = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    req_t = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'אישור_חפיפה', 'שעות_פעילות']
    for col in req_s:
        if col not in s_df.columns: errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ חיילים.")
    for col in req_t:
        if col not in t_df.columns: errors.append(f"❌ שגיאת מבנה: חסרה העמודה '{col}' בקובץ משימות.")
    return errors

# ==========================================
# 4. אלגוריתם האופטימיזציה
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
            blocking_tasks = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking_tasks) <= 1)
        
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for s in soldiers:
            row = {"שם": s.name}
            for h in range(num_hours):
                lbl = f"{h:02d}:00"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[lbl] = " + ".join(active) if active else "-"
            res.append(row)
        return pd.DataFrame(res)
    return None

# ==========================================
# 5. ממשק המשתמש (Tabs)
# ==========================================
st.title("🪖 מערכת שיבוץ כוח אדם מקצועית")

tab1, tab2, tab3 = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך מפורט למשתמש", "📥 הורדת תבניות אקסל"])

# --- לשונית 3: תבניות ---
with tab3:
    st.header("הורדת תבניות עבודה")
    st.write("אנא השתמש בקבצים אלו כבסיס לנתונים שלך כדי למנוע שגיאות.")
    s_ex = pd.DataFrame([{'מספר_אישי': 1234567, 'שם': 'ישראל ישראלי', 'הכשרות': 'נהג', 'פטורים': '102', 'מקסימום_שעות_ברצף': 6}])
    t_ex = pd.DataFrame([
        {'קוד_משימה': 101, 'שם': 'אבטחה', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'},
        {'קוד_משימה': 102, 'שם': 'כיתת כוננות', 'כוח_אדם_נדרש': 8, 'משך_זמן': 25, 'אישור_חפיפה': True, 'שעות_פעילות': 'all'}
    ])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel_file(s_ex), file_name="Soldiers_Template.xlsx")
    with c2: st.download_button("📥 הורד תבנית משימות", data=to_excel_file(t_ex), file_name="Tasks_Template.xlsx")

# --- לשונית 2: מדריך למשתמש ---
with tab2:
    st.header("📖 מדריך למילוי נכון של קבצי המערכת")
    st.write("כדי שהאלגוריתם יצליח לשבץ את הכוחות בצורה מושלמת, יש להקפיד על הכללים הבאים:")

    st.subheader("1️⃣ קובץ חיילים (Soldiers)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר</th><th>דוגמה</th></tr>
        <tr><td><b>מספר_אישי</b></td><td>מספר זיהוי ייחודי לכל חייל (חובה)</td><td>1234567</td></tr>
        <tr><td><b>שם</b></td><td>שם החייל כפי שיופיע בלוח השיבוץ</td><td>אבי כהן</td></tr>
        <tr><td><b>הכשרות</b></td><td>מיומנויות מופרדות בפסיק (אם אין, השאר ריק)</td><td>נהג, חובש</td></tr>
        <tr><td><b>פטורים</b></td><td>קודי משימות (מספרים) שהחייל <b>לא יכול</b> לבצע</td><td>101, 105</td></tr>
        <tr><td><b>מקסימום_שעות_ברצף</b></td><td>מספר שעות מקסימלי שחייל יכול לבצע במשימה</td><td>6</td></tr>
    </table>
    """, unsafe_allow_html=True)

    st.subheader("2️⃣ קובץ משימות (Tasks)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר</th><th>ערכים אפשריים</th></tr>
        <tr><td><b>קוד_משימה</b></td><td>מספר מזהה (חייב להתאים למה שכתוב ב'פטורים')</td><td>101</td></tr>
        <tr><td><b>שם</b></td><td>שם המשימה (אבטחה, סיור וכו')</td><td>שמירת שער</td></tr>
        <tr><td><b>כוח_אדם_נדרש</b></td><td>כמות חיילים שחייבים להיות במשימה בכל רגע</td><td>2</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td>האם חייל יכול לעשות משימה נוספת במקביל?</td><td><b>True</b> (כן) / <b>False</b> (לא)</td></tr>
        <tr><td><b>שעות_פעילות</b></td><td>מתי המשימה קורית (שעות מופרדות בפסיק)</td><td>all / 8,9,10,11</td></tr>
    </table>
    """, unsafe_allow_html=True)

    st.info("💡 **דוגמה לכיתת כוננות:** הגדר 'אישור_חפיפה' כ-`True` ו'שעות_פעילות' כ-`all`. כך החייל יהיה כונן במקביל לכל משימה אחרת.")

# --- לשונית 1: העלאה ושיבוץ ---
with tab1:
    st.subheader("🚀 העלאת קבצים והרצת האלגוריתם")
    up1, up2 = st.columns(2)
    with up1: sf = st.file_uploader("📂 העלה קובץ חיילים (Excel)", type="xlsx")
    with up2: tf = st.file_uploader("📂 העלה קובץ משימות (Excel)", type="xlsx")

    if sf and tf:
        try:
            s_df = pd.read_excel(sf)
            t_df = pd.read_excel(tf)
            errs = validate_data(s_df, t_df)
            if errs:
                for e in errs: st.error(e)
                st.warning("⚠️ לא ניתן להמשיך עד ששמות העמודות באקסל יתוקנו (ראה מדריך).")
            else:
                st.success("✅ הקבצים תקינים!")
                if st.button("⚙️ צור שיבוץ אופטימלי", type="primary", use_container_width=True):
                    with st.spinner("המערכת מחשבת את הפתרון הטוב ביותר..."):
                        s_list = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                        t_list = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                        res = solve_scheduling(s_list, t_list)
                        if res is not None:
                            st.balloons()
                            st.subheader("🗓️ לוח השיבוץ הסופי")
                            st.dataframe(res, use_container_width=True)
                            excel_res = to_excel_file(res, sheet_name='Schedule')
                            st.download_button("📥 הורד לוח שיבוץ סופי (Excel)", data=excel_res, file_name="Final_Schedule.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
                        else:
                            st.error("❌ לא נמצא פתרון. ייתכן שיש חוסר בכוח אדם למשימות שהגדרת.")
        except Exception as ex:
            st.error(f"שגיאה בלתי צפויה: {ex}")
