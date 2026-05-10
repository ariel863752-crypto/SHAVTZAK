import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. עיצוב חמ"ל מבצעי (Military UI CSS)
# ==========================================
st.set_page_config(page_title="חמ''ל שיבוץ חכם", page_icon="⚔️", layout="wide")

st.markdown("""
    <style>
    /* רקע וצבעים צבאיים */
    .stApp {
        background-color: #3E442B; /* ירוק זית עמוק */
        color: #F1F1F1;
        direction: rtl;
    }
    [data-testid="stSidebar"] {
        background-color: #2B2F1B !important;
        border-left: 2px solid #4B5320;
    }
    h1, h2, h3, h4 { color: #D4C5A1 !important; text-align: right; }
    
    /* עיצוב כפתור כתום מבצעי */
    div.stButton > button:first-child {
        background-color: #FF6E40 !important;
        color: white !important;
        border: none !important;
        font-weight: bold !important;
        width: 100%;
        height: 3em;
        border-radius: 8px;
    }
    
    /* טבלאות ויישור לימין */
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    .stDataFrame { background-color: #2B2F1B; border-radius: 5px; }
    
    /* עיצוב המדריך */
    .guide-table { width: 100%; border-collapse: collapse; background-color: #2B2F1B; color: #D4C5A1; }
    .guide-table th { background-color: #4B5320; padding: 10px; border: 1px solid #6B7340; }
    .guide-table td { padding: 8px; border: 1px solid #4B5320; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. מחלקות הנתונים (Logic Objects)
# ==========================================
class Soldier:
    def __init__(self, s_id, name, qual="", restr="", max_h=6):
        self.soldier_id = str(s_id)
        self.name = name
        self.qualifications = [str(q).strip() for q in str(qual).split(',')] if pd.notna(qual) else []
        self.restricted_tasks = [int(float(t)) for t in str(restr).split(',') if str(t).strip().replace('.0','').isdigit()] if pd.notna(restr) else []
        self.max_h = int(max_h) if pd.notna(max_h) and str(max_h).isdigit() else 6

class Task:
    def __init__(self, t_id, name, req_p, dur, overlap=False, hours="all"):
        self.task_id = int(t_id)
        self.name = name
        self.required_personnel = int(req_p)
        self.duration = int(dur)
        self.allow_overlap = str(overlap).lower() == 'true'
        if pd.isna(hours) or str(hours).lower() in ['all', '']:
            self.active_hours = list(range(25))
        else:
            self.active_hours = [int(x.strip()) for x in str(hours).split(',') if str(x).strip().isdigit()]

# ==========================================
# 3. פונקציות עזר (Excel & Charts)
# ==========================================
def to_excel_file(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#4B5320', 'font_color': 'white'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
    return output.getvalue()

# ==========================================
# 4. אלגוריתם אופטימיזציה (The Engine)
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=25):
    model = cp_model.CpModel()
    x = {}
    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")

    # כיסוי משימות
    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers: model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אילוץ משימה אחת (ללא חפיפה) ופטורים
    for s in soldiers:
        for h in range(num_hours):
            model.Add(sum(x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap) <= 1)
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אופטימיזציה: חלוקת עומס שוויונית
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_loads: model.Add(max_load >= load)
    model.Minimize(max_load)

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
            row["סך שעות"] = sum(1 for h in range(num_hours) if any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks))
            res.append(row)
        return pd.DataFrame(res)
    return None

# ==========================================
# 5. ממשק המשתמש (UI Structure)
# ==========================================
st.markdown("<h1 style='text-align: center;'>⚔️ חמ''ל שיבוץ כוחות מבצעי</h1>", unsafe_allow_html=True)

# סרגל צדי להדרכה ותבניות
with st.sidebar:
    st.header("⚙️ הגדרות ומדריכים")
    with st.expander("📥 הורדת תבניות Excel"):
        s_tmp = pd.DataFrame([{'מספר_אישי': 101, 'שם': 'חייל א', 'הכשרות': '', 'פטורים': '', 'מקסימום_שעות_ברצף': 6}])
        t_tmp = pd.DataFrame([{'קוד_משימה': 1, 'שם': 'שמירה', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'}])
        st.download_button("תבנית חיילים", data=to_excel_file(s_tmp), file_name="Soldiers_Template.xlsx")
        st.download_button("תבנית משימות", data=to_excel_file(t_tmp), file_name="Tasks_Template.xlsx")
    
    st.markdown("---")
    st.subheader("📖 דגשים למילוי")
    st.info("כיתת כוננות: הגדר 'אישור_חפיפה' כ-True ומשך זמן 25 שעות.\nפטורים: רשום את קוד המשימה.")

# עמוד ראשי - העלאה
tab_work, tab_guide = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך מלא"])

with tab_guide:
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>הסבר</th></tr>
        <tr><td>אישור_חפיפה</td><td><b>True</b> למשימות שניתן לבצע במקביל לאחרות (כוננות)</td></tr>
        <tr><td>פטורים</td><td>מספרי המשימות שהחייל לא יכול לבצע (מופרד בפסיק)</td></tr>
        <tr><td>שעות_פעילות</td><td><b>all</b> לכל היום, או שעות ספציפיות: 8,9,10</td></tr>
    </table>
    """, unsafe_allow_html=True)

with tab_work:
    col1, col2 = st.columns(2)
    with col1: sf = st.file_uploader("📂 העלה רשימת סד''כ (חיילים)", type="xlsx")
    with col2: tf = st.file_uploader("📂 העלה רשימת משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        st.success("📊 נתונים נקלטו. מוכן להרצה.")
        
        if st.button("🚀 צור שיבוץ אופטימלי והוגן"):
            with st.spinner("האלגוריתם מחשב חלוקת עומס מיטבית..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                
                final_df = solve_scheduling(soldiers, tasks)
                
                if final_df is not None:
                    # --- מד עומס (Dashboard) ---
                    total_available = len(soldiers) * 25
                    total_needed = final_df["סך שעות"].sum()
                    utilization = (total_needed / total_available) * 100
                    
                    st.divider()
                    st.subheader("📈 לוח מחוונים מבצעי")
                    m1, m2 = st.columns(2)
                    m1.metric("ניצולת סד''כ כוללת", f"{utilization:.1f}%")
                    m2.metric("ממוצע שעות לחייל", f"{final_df['סך שעות'].mean():.1f}")
                    
                    st.progress(utilization / 100)
                    
                    # הצגת התוצאה
                    st.subheader("🗓️ לוח שיבוץ סופי")
                    st.dataframe(final_df, use_container_width=True)
                    
                    # הורדת התוצאה
                    ex_data = to_excel_file(final_df, sheet_name='Shavtzak')
                    st.download_button("📥 הורד לוח שיבוץ סופי (Excel)", data=ex_data, file_name="Final_Schedule.xlsx", use_container_width=True)
                    
                    # גרף עומסים
                    fig = px.bar(final_df, x="שם", y="סך שעות", title="חלוקת עומס בין החיילים", color="סך שעות", color_continuous_scale="Greens")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.error("❌ לא נמצא פתרון. נסה להוסיף חיילים או להפחית משימות.")
