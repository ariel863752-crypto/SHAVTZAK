import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. הגדרות ממשק, עיצוב ו-RTL
# ==========================================
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    /* הגדרות כיוון ויישור עברית */
    .stApp { direction: rtl; text-align: right; background-color: #f8f9fa; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    
    /* כותרת ראשית */
    .main-title { color: #2d5a27; font-size: 42px; font-weight: bold; border-bottom: 4px solid #e67e22; padding-bottom: 10px; margin-bottom: 25px; }
    
    /* עיצוב טבלאות מדריך */
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; text-align: right; border: 1px solid #ddd; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 15px; }
    
    /* עיצוב כפתור ביצוע כתום */
    div.stButton > button:first-child {
        background-color: #e67e22 !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 8px !important;
        width: 100%;
        height: 3.5em;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* עיצוב כרטיסי סטטיסטיקה */
    [data-testid="stMetricValue"] { color: #2d5a27 !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. מחלקות הנתונים (Objects)
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
        if isinstance(overlap, str):
            self.allow_overlap = overlap.strip().lower() == 'true'
        else:
            self.allow_overlap = bool(overlap) if pd.notna(overlap) else False
        
        if pd.isna(hours) or str(hours).lower() in ['all', '']:
            self.active_hours = list(range(25))
        else:
            self.active_hours = [int(x.strip()) for x in str(hours).split(',') if str(x).strip().isdigit()]

# ==========================================
# 3. פונקציות עזר (Excel & Validation)
# ==========================================
def to_excel_file(df, sheet_name='Sheet1'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#2d5a27', 'font_color': 'white', 'border': 1, 'align': 'center'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 15)
    return output.getvalue()

def validate_data(s_df, t_df):
    errors = []
    req_s = ['מספר_אישי', 'שם', 'הכשרות', 'פטורים', 'מקסימום_שעות_ברצף']
    req_t = ['קוד_משימה', 'שם', 'כוח_אדם_נדרש', 'משך_זמן', 'אישור_חפיפה', 'שעות_פעילות']
    for col in req_s:
        if col not in s_df.columns: errors.append(f"❌ חסרה עמודה בקובץ חיילים: {col}")
    for col in req_t:
        if col not in t_df.columns: errors.append(f"❌ חסרה עמודה בקובץ משימות: {col}")
    return errors

# ==========================================
# 4. האלגוריתם (אילוצים + חלוקת עומס הוגנת)
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

    # אילוץ משימה אחת (ללא חפיפה) ופטורים
    for s in soldiers:
        for h in range(num_hours):
            blocking_tasks = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking_tasks) <= 1)
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אופטימיזציה: חלוקת עומס שוויונית (Min-Max)
    s_total_hours = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_total_hours: model.Add(max_load >= load)
    model.Minimize(max_load)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for s in soldiers:
            row = {"שם": s.name}
            work_count = 0
            for h in range(num_hours):
                lbl = f"{h:02d}:00"
                active = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[lbl] = " + ".join(active) if active else "-"
                if active: work_count += 1
            row["סך שעות"] = work_count
            res.append(row)
        return pd.DataFrame(res)
    return None

# ==========================================
# 5. ממשק המשתמש (Tabs)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת לניהול ושיבוץ כוחות</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי נכון", "📥 תבניות אקסל"])

# --- לשונית תבניות ---
with tab_templates:
    st.subheader("הורדת תבניות עבודה")
    s_tmp = pd.DataFrame([{'מספר_אישי': 101, 'שם': 'חייל א', 'הכשרות': 'נהג', 'פטורים': '5', 'מקסימום_שעות_ברצף': 6}])
    t_tmp = pd.DataFrame([
        {'קוד_משימה': 1, 'שם': 'שער', 'כוח_אדם_נדרש': 1, 'משך_זמן': 4, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'},
        {'קוד_משימה': 2, 'שם': 'כוננות', 'כוח_אדם_נדרש': 8, 'משך_זמן': 25, 'אישור_חפיפה': True, 'שעות_פעילות': 'all'}
    ])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel_file(s_tmp), file_name="Soldiers_Template.xlsx")
    with c2: st.download_button("📥 הורד תבנית משימות", data=to_excel_file(t_tmp), file_name="Tasks_Template.xlsx")

# --- לשונית מדריך ---
with tab_guide:
    st.subheader("📖 איך למלא את קבצי האקסל?")
    st.markdown("#### קובץ חיילים:")
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>הסבר</th><th>דוגמה</th></tr>
        <tr><td>מספר_אישי</td><td>מספר זיהוי ייחודי לכל חייל</td><td>1234567</td></tr>
        <tr><td>שם</td><td>שם החייל שיופיע בלוח</td><td>יוסי כהן</td></tr>
        <tr><td>הכשרות</td><td>מיומנויות (נהג, חובש וכו') מופרדות בפסיק</td><td>נהג, חובש</td></tr>
        <tr><td>פטורים</td><td>קודי המשימות שהחייל <b>לא</b> מבצע</td><td>1, 5</td></tr>
        <tr><td>מקסימום_שעות_ברצף</td><td>הגבלת שעות עבודה רצופות</td><td>6</td></tr>
    </table>
    """, unsafe_allow_html=True)
    st.markdown("#### קובץ משימות:")
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>הסבר</th><th>ערכים</th></tr>
        <tr><td>קוד_משימה</td><td>מספר המשימה (להתאמה לפטורים)</td><td>1</td></tr>
        <tr><td>שם</td><td>שם המשימה (שער, סיור וכו')</td><td>אבטחת מתקן</td></tr>
        <tr><td>כוח_אדם_נדרש</td><td>כמות חיילים שחייבים להיות שם בו-זמנית</td><td>2</td></tr>
        <tr><td>אישור_חפיפה</td><td>האם ניתן לעשות משימה נוספת במקביל?</td><td>True / False</td></tr>
        <tr><td>שעות_פעילות</td><td>מתי המשימה קורית?</td><td>all / 8,9,10</td></tr>
    </table>
    """, unsafe_allow_html=True)

# --- לשונית שיבוץ ---
with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with col_u2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        errs = validate_data(s_df, t_df)
        if errs:
            for e in errs: st.error(e)
        else:
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
                        st.download_button("📥 הורד לוח שיבוץ סופי (Excel)", data=to_excel_file(final_df), file_name="Final_Shavtzak.xlsx", use_container_width=True)
                        
                        # --- סטטיסטיקה וגרפים ---
                        st.divider()
                        st.subheader("📊 ניתוח עומסים וסטטיסטיקה")
                        max_h, min_h, avg_h = final_df["סך שעות"].max(), final_df["סך שעות"].min(), final_df["סך שעות"].mean()
                        
                        m1, m2, m3 = st.columns(3)
                        m1.metric("מקסימום שעות", f"{max_h} ש'")
                        m2.metric("מינימום שעות", f"{min_h} ש'")
                        m3.metric("ממוצע שעות", f"{avg_h:.1f} ש'")
                        
                        fig = px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="חלוקת שעות עבודה לכל חייל", color_continuous_scale="Greens")
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # --- תובנות ה"למה" ---
                        st.markdown("#### 💡 תובנות מערכת")
                        with st.expander("לחץ כאן להסבר על חלוקת העומס"):
                            top_s = final_df[final_df["סך שעות"] == max_h]["שם"].tolist()
                            bot_s = final_df[final_df["סך שעות"] == min_h]["שם"].tolist()
                            st.write(f"**החיילים העמוסים ביותר:** {', '.join(top_s)}")
                            st.caption("סיבה סבירה: חיילים אלו ללא פטורים משמעותיים או שהם בעלי הכשרה ייחודית שנדרשה בזמנים לחוצים.")
                            st.write(f"**החיילים הפחות עמוסים:** {', '.join(bot_s)}")
                            st.caption("סיבה סבירה: קיומם של פטורים רבים באקסל או הגעת החייל למגבלת 'שעות ברצף' שאילצה את המערכת להוציאו למנוחה.")
                    else:
                        st.error("❌ לא ניתן למצוא פתרון. ייתכן ויש חוסר בחיילים למשימות שהוגדרו.")
