import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. עיצוב ממשק RTL וסטיילינג מקצועי
# ==========================================
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    /* הגדרות כיוון ויישור עברית */
    .stApp { direction: rtl; text-align: right; background-color: #fcfcfc; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    
    /* כותרת ראשית */
    .main-title { color: #2d5a27; font-size: 45px; font-weight: bold; border-bottom: 4px solid #e67e22; padding-bottom: 10px; margin-bottom: 25px; }
    
    /* עיצוב טבלאות מדריך */
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 25px; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; text-align: right; border: 1px solid #ddd; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 15px; }
    
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
# 3. המנוע המתמטי - שינה, מנוחה, רציפות והוגנות
# ==========================================
def solve_scheduling(soldiers, tasks, num_hours=24):
    model = cp_model.CpModel()
    x, start = {}, {}

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"x_{s.soldier_id}_{t.task_id}_{h}")
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f"start_{s.soldier_id}_{t.task_id}_{h}")

    for s in soldiers:
        for t in tasks:
            for h in range(num_hours):
                # רציפות המשמרת
                relevant_starts = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))
                
                # מנוחה לאחר משימה
                if t.rest_duration > 0:
                    total_busy = t.shift_duration + t.rest_duration
                    for next_h in range(h + 1, min(h + total_busy, num_hours)):
                        for other_t in tasks:
                            if not other_t.allow_overlap:
                                model.AddImplication(start[s.soldier_id, t.task_id, h], x[s.soldier_id, other_t.task_id, next_h].Not())

                if h + t.shift_duration > num_hours:
                    model.Add(start[s.soldier_id, t.task_id, h] == 0)

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
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(num_hours): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # אופטימיזציה: הוגנות + שינת לילה (22:00-08:00)
    night_hours = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7]
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_loads: model.Add(max_load >= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in night_hours)
    model.Minimize(30 * max_load + night_work)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 25.0
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
# 4. ממשק המשתמש (Tabs)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת שיבוץ כוחות חכמה</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות אקסל"])

with tab_templates:
    st.subheader("📥 הורדת תבניות עבודה")
    s_tmp = pd.DataFrame([{'מספר_אישי': 1001, 'שם': 'ישראל ישראלי', 'פטורים': '5'}])
    t_tmp = pd.DataFrame([{'קוד_משימה': 1, 'שם': 'שמירה', 'כוח_אדם_נדרש': 1, 'משך_משמרת': 4, 'שעות_מנוחה_אחרי': 8, 'אישור_חפיפה': False, 'שעות_פעילות': 'all'}])
    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 הורד תבנית חיילים", data=to_excel(s_tmp), file_name="Template_Soldiers.xlsx")
    with c2: st.download_button("📥 הורד תבנית משימות", data=to_excel(t_tmp), file_name="Template_Tasks.xlsx")

with tab_guide:
    st.subheader("📖 מדריך למילוי שבצ''ק - הגדרות מבצעיות")
    
    st.markdown("#### 1️⃣ קובץ חיילים (Soldiers)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>מה למלא?</th><th>למה זה חשוב?</th></tr>
        <tr><td><b>מספר_אישי</b></td><td>מספר זיהוי ייחודי לכל חייל.</td><td>מונע כפילויות במערכת השיבוץ.</td></tr>
        <tr><td><b>שם</b></td><td>שם החייל כפי שיופיע בלוח.</td><td>זיהוי ברור בלוח השיבוץ הסופי.</td></tr>
        <tr><td><b>פטורים</b></td><td>מספרי 'קוד משימה' (למשל 1, 5).</td><td>מונע שיבוץ חייל למשימה שאסור לו לבצע.</td></tr>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 2️⃣ קובץ משימות (Tasks)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>מה למלא?</th><th>למה זה חשוב?</th></tr>
        <tr><td><b>קוד_משימה</b></td><td>מספר המשימה (למשל 1).</td><td>חובה להתאמה מול עמודת ה'פטורים'.</td></tr>
        <tr><td><b>שם</b></td><td>שם המשימה (שער, סיור, מטבח).</td><td>יוצג בלוח השיבוץ לכל שעה.</td></tr>
        <tr><td><b>כוח_אדם_נדרש</b></td><td>כמות חיילים בעמדה בו-זמנית.</td><td>מבטיח שהעמדה מאוישת לפי הצורך.</td></tr>
        <tr><td><b>משך_משמרת</b></td><td>שעות רצופות (למשל 4).</td><td>החייל יינעל למשימה ולא יוחלף באמצע.</td></tr>
        <tr><td><b>שעות_מנוחה_אחרי</b></td><td>שעות הפסקה חובה (למשל 8).</td><td>המערכת תמנע שיבוץ נוסף עד לסיום המנוחה.</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td><b>True</b> לכוננות, <b>False</b> לאחרות.</td><td>מאפשר ביצוע המשימה במקביל למשימה אחרת.</td></tr>
        <tr><td><b>שעות_פעילות</b></td><td><b>all</b> ל-24/7 או שעות (8,9,10).</td><td>קובע מתי העמדה חייבת להיות מאוישת.</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    st.info("💡 **ניהול שינה:** המערכת מתעדפת אוטומטית 7 שעות שינה רצופות בלילה (בין 22:00 ל-08:00) לכל חייל.")

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
                    st.subheader("📊 ניתוח עומסים ותובנות מערכת")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("ממוצע שעות מחלקתי", f"{final_df['סך שעות'].mean():.1f}")
                    m2.metric("חיילים שעבדו בלילה", f"{final_df[final_df['שעות לילה'] > 0].shape[0]}")
                    m3.metric("פער עומס (מקס'-מינ')", f"{(final_df['סך שעות'].max() - final_df['סך שעות'].min()):.0f} ש'")
                    
                    st.plotly_chart(px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="חלוקת עומס כללית", color_continuous_scale="Greens"), use_container_width=True)
                    st.plotly_chart(px.bar(final_df, x="שם", y="שעות לילה", title="פגיעה בשינה (22:00 עד 08:00)", color="שעות לילה", color_continuous_scale="Reds"), use_container_width=True)
                    
                    with st.expander("💡 למה הלו\"ז נראה ככה?"):
                        night_req = t_df[t_df['שעות_פעילות'].astype(str).str.contains('all|22|23|0|1|2|3|4|5|6|7')]['כוח_אדם_נדרש'].sum()
                        st.write(f"**עבודה בלילה:** יש דרישה ל-{night_req} חיילים בלילה. המערכת חייבת לשבץ אותם למרות השאיפה לשינה.")
                        max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                        st.write(f"**העמוסים ביותר:** {', '.join(max_s)}. אלו חיילים ללא פטורים שזמינים למילוי 'חורים' שנוצרו עקב מנוחות של אחרים.")
                else:
                    st.error("❌ לא נמצא פתרון. בדוק אם יש מספיק חיילים למשימות ולמנוחות הנדרשות.")
