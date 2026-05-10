import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. עיצוב ממשק נקי וקריא (Professional & Clean)
# ==========================================
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    /* הגדרות כיוון ויישור עברית */
    .stApp { direction: rtl; text-align: right; background-color: #f4f7f6; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    
    /* כותרות מקצועיות */
    .main-title { color: #2d5a27; font-size: 45px; font-weight: bold; border-bottom: 4px solid #2d5a27; padding-bottom: 10px; margin-bottom: 20px; }
    h2, h3 { color: #2d5a27; }

    /* עיצוב טבלאות מדריך - קריאות מקסימלית */
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; text-align: right; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 16px; }
    
    /* עיצוב כפתור כתום בולט */
    div.stButton > button:first-child {
        background-color: #e67e22 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 5px !important;
        width: 100%;
        height: 3.5em;
        border: none;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. לוגיקה ואלגוריתם (אותו מוח חזק)
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
        self.active_hours = list(range(25)) if str(hours).lower() in ['all', ''] else [int(x.strip()) for x in str(hours).split(',') if str(x).strip().isdigit()]

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
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # חלוקת עומס שוויונית
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

def to_excel_file(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='שבצ"ק סופי')
    return output.getvalue()

# ==========================================
# 3. ממשק האתר (Structure)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת שיבוץ כוחות חכמה</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי נכון", "📥 תבניות אקסל"])

# --- לשונית תבניות ---
with tab_templates:
    st.subheader("הורדת תבניות עבודה")
    st.write("הורידו את הקבצים הבאים, מלאו אותם והעלו בלשונית 'ביצוע שיבוץ'.")
    s_tmp = pd.DataFrame([{'מספר_אישי': 101, 'שם': 'חייל דוגמה', 'הכשרות': 'נהג', 'פטורים': '5', 'מקסימום_שעות_ברצף': 6}])
    t_tmp = pd.DataFrame([{'קוד_משימה': 1, 'שם': 'שמירה', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'}])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel_file(s_tmp), file_name="Template_Soldiers.xlsx")
    with c2: st.download_button("📥 הורד תבנית משימות", data=to_excel_file(t_tmp), file_name="Template_Tasks.xlsx")

# --- לשונית מדריך מפורט ---
with tab_guide:
    st.subheader("📖 איך למלא את הקבצים נכון?")
    
    st.markdown("#### 1. קובץ חיילים (Soldiers)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר</th><th>דוגמה</th></tr>
        <tr><td>מספר_אישי</td><td>מספר ייחודי לכל חייל</td><td>1234567</td></tr>
        <tr><td>שם</td><td>שם החייל שיופיע בלוח</td><td>יוסי כהן</td></tr>
        <tr><td>הכשרות</td><td>מיומנויות מופרדות בפסיק</td><td>נהג, חובש</td></tr>
        <tr><td>פטורים</td><td>מספר קוד המשימה שהחייל <b>לא</b> מבצע</td><td>1, 5</td></tr>
        <tr><td>מקסימום_שעות_ברצף</td><td>הגבלת שעות עבודה רצופות</td><td>6</td></tr>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 2. קובץ משימות (Tasks)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר</th><th>ערכים</th></tr>
        <tr><td>קוד_משימה</td><td>מספר המשימה (חובה להתאים ל'פטורים')</td><td>1</td></tr>
        <tr><td>שם</td><td>שם המשימה</td><td>סיור שער</td></tr>
        <tr><td>כוח_אדם_נדרש</td><td>כמות חיילים שחייבים להיות במשימה בכל רגע</td><td>2</td></tr>
        <tr><td>אישור_חפיפה</td><td>האם משימה נוספת יכולה לקרות במקביל?</td><td>True / False</td></tr>
        <tr><td>שעות_פעילות</td><td>מתי המשימה קורית?</td><td>all (כל היום) / 8,9,10</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    st.info("💡 **כדי להגדיר כיתת כוננות:** הגדירו אישור_חפיפה כ-True ושעות_פעילות כ-all.")

# --- לשונית ביצוע שיבוץ ---
with tab_run:
    col1, col2 = st.columns(2)
    with col1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with col2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        st.success("✅ הקבצים נקלטו. מוכן להרצה.")
        
        if st.button("⚙️ צור שבצ''ק אופטימלי והוגן"):
            with st.spinner("מחשב חלוקת עומסים הוגנת..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('הכשרות'), r.get('פטורים'), r.get('מקסימום_שעות_ברצף')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r['משך_זמן'], r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                
                final_df = solve_scheduling(soldiers, tasks)
                
                if final_df is not None:
                    st.balloons()
                    st.subheader("🗓️ שבצ''ק סופי")
                    st.dataframe(final_df, use_container_width=True)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.download_button("📥 הורד לוח סופי (Excel)", data=to_excel_file(final_df), file_name="Final_Shavtzak.xlsx", use_container_width=True)
                    
                    st.divider()
                    st.subheader("📊 ניתוח חלוקת עומס")
                    fig = px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="שעות עבודה לכל חייל", color_continuous_scale="Greens")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.error("❌ לא נמצא פתרון חוקי. בדוק אם יש מספיק חיילים.")
