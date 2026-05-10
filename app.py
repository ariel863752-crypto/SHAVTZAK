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
    /* הגדרות כיוון ויישור עברית */
    .stApp { direction: rtl; text-align: right; background-color: #fcfcfc; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    
    /* כותרת ראשית בסטייל שאהבת */
    .main-title { color: #2d5a27; font-size: 45px; font-weight: bold; border-bottom: 4px solid #e67e22; padding-bottom: 10px; margin-bottom: 25px; }
    
    /* עיצוב טבלאות מדריך - קריאות מקסימלית */
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; text-align: right; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 16px; }
    
    /* עיצוב כפתור כתום בולט */
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
# 3. המנוע המתמטי - ניהול משמרות, מנוחה ושינה
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=24):
    model = cp_model.CpModel()
    x = {}      # חייל s במשימה t בשעה h
    start = {}  # חייל s מתחיל משמרת t בשעה h

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"start_{s.soldier_id}_{t.task_id}_{h}")

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                # רציפות המשמרת (הנעילה)
                relevant_starts = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))
                
                # מנוחה לאחר משימה (השליטה שלך)
                if t.rest_duration > 0:
                    total_busy = t.shift_duration + t.rest_duration
                    for next_h in range(h + 1, min(h + total_busy, num_hours)):
                        for other_t in tasks:
                            if not other_t.allow_overlap:
                                model.AddImplication(start[s.soldier_id, t.task_id, h], x[s.soldier_id, other_t.task_id, next_h].Not())

                if h + t.shift_duration > num_hours:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

    # אילוץ אייש עמדות
    for h in range(num_hours):
        for t in tasks:
            if h in t.active_hours:
                model.Add(sum(x[s.soldier_id, t.task_id, h] for s in soldiers) == t.required_personnel)
            else:
                for s in soldiers: model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אילוצי חייל (משימה אחת ופטורים)
    for s in soldiers:
        for h in range(num_hours):
            blocking = [x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap]
            model.Add(sum(blocking) <= 1)
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אופטימיזציה: הוגנות + שינת לילה (22:00 עד 08:00)
    night_hours = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7]
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_loads: model.Add(max_load >= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in night_hours)
    model.Minimize(20 * max_load + night_work)

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
# 4. ממשק המשתמש (UI)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת שיבוץ כוחות חכמה</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות אקסל"])

with tab_templates:
    st.subheader("הורדת תבניות עבודה")
    s_tmp = pd.DataFrame([{'מספר_אישי': 1001, 'שם': 'ישראל ישראלי', 'פטורים': '5'}])
    t_tmp = pd.DataFrame([
        {'קוד_משימה': 1, 'שם': 'שמירת שער', 'כוח_אדם_נדרש': 1, 'משך_משמרת': 4, 'שעות_מנוחה_אחרי': 8, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'},
        {'קוד_משימה': 2, 'שם': 'כוננות', 'כוח_אדם_נדרש': 8, 'משך_משמרת': 24, 'שעות_מנוחה_אחרי': 0, 'אישור_חפיפה': True, 'שעות_פעילות': 'all'}
    ])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel(s_tmp), file_name="Template_Soldiers.xlsx")
    with c2: st.download_button("📥 הורד תבנית משימות", data=to_excel(t_tmp), file_name="Template_Tasks.xlsx")

with tab_guide:
    st.subheader("📖 מדריך למילוי שבצ''ק - צעד אחר צעד")
    
    st.markdown("#### 1️⃣ קובץ חיילים (Soldiers)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>מה למלא?</th><th>דוגמה</th></tr>
        <tr><td><b>מספר_אישי</b></td><td>מספר ייחודי לכל חייל (למניעת כפילויות)</td><td>1234567</td></tr>
        <tr><td><b>שם</b></td><td>שם החייל שיופיע בלוח הסופי</td><td>יוסי כהן</td></tr>
        <tr><td><b>פטורים</b></td><td>מספרי 'קוד משימה' שהחייל <b>לא יכול</b> לבצע</td><td>1, 5</td></tr>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 2️⃣ קובץ משימות (Tasks)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>מה למלא?</th><th>הסבר מבצעי</th></tr>
        <tr><td><b>קוד_משימה</b></td><td>מספר המשימה (למשל 1)</td><td>חובה להתאים למה שכתוב ב'פטורים'</td></tr>
        <tr><td><b>שם</b></td><td>שם המשימה</td><td>שמירת שער, מטבח, סיור</td></tr>
        <tr><td><b>כוח_אדם_נדרש</b></td><td>כמה חיילים בעמדה בו-זמנית</td><td>2 (לשמירה בזוגות)</td></tr>
        <tr><td><b>משך_משמרת</b></td><td>כמה שעות רצופות חייל עושה (הנעילה)</td><td>4 (החייל לא יוחלף באמצע 4 השעות)</td></tr>
        <tr><td><b>שעות_מנוחה_אחרי</b></td><td>כמה מנוחה חייב החייל בסיום</td><td>8 (למשימה מעייפת) או 0 (למשימה קלה)</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td>האם ניתן לעשות משימה נוספת במקביל?</td><td><b>True</b> לכוננות, <b>False</b> לכל השאר</td></tr>
        <tr><td><b>שעות_פעילות</b></td><td>מתי המשימה קורית</td><td><b>all</b> ל-24/7 או שעות מופרדות בפסיק (8,9,10)</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    st.info("💡 **ניהול שינה:** המערכת תנסה אוטומטית לתת לכל חייל 7 שעות שינה רצופות בלילה (בין 22:00 ל-08:00) ככל שהסד''כ מאפשר.")

with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with col_u2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק אופטימלי"):
            with st.spinner("מחשב שינה, מנוחה, משמרות וחלוקת עומס..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'), r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                final_df = solve_scheduling(soldiers, tasks)
                
                if final_df is not None:
                    st.balloons()
                    st.subheader("🗓️ שבצ''ק סופי")
                    st.table(final_df)
                    st.download_button("📥 הורד לוח סופי (Excel)", data=to_excel(final_df), file_name="Final_Shavtzak.xlsx", use_container_width=True)
                    
                    st.divider()
                    st.subheader("📊 ניתוח עומסים ושינת לילה")
                    m1, m2 = st.columns(2)
                    m1.metric("ממוצע שעות לחייל", f"{final_df['סך שעות'].mean():.1f}")
                    m2.metric("חיילים שעבדו בלילה", f"{final_df[final_df['שעות לילה'] > 0].shape[0]}")
                    
                    st.plotly_chart(px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="חלוקת עומס כללית", color_continuous_scale="Greens"), use_container_width=True)
                    st.plotly_chart(px.bar(final_df, x="שם", y="שעות לילה", title="פגיעה בשינה (עבודה בין 22:00 ל-08:00)", color="שעות לילה", color_continuous_scale="Reds"), use_container_width=True)
                else:
                    st.error("❌ לא נמצא פתרון. בדוק אם יש מספיק חיילים למשימות ולמנוחות הנדרשות.")
