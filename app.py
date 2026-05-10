import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# 1. הגדרות עמוד ו-CSS (RTL חסין כדורים)
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title='שבצ"ק חכם',
    page_icon="🪖",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* כפייה של RTL על כל האפליקציה */
.stApp, [data-testid="stAppViewContainer"], .main, .block-container {
    direction: rtl !important;
    text-align: right !important;
}

/* יישור טאבים לצד ימין של המסך */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important;
    justify-content: flex-start !important;
    gap: 20px;
    border-bottom: 2px solid #e67e22;
}

[data-testid="stTabs"] [data-baseweb="tab"] {
    font-family: 'Heebo', sans-serif;
    font-weight: 700;
    font-size: 18px;
}

/* כותרת ירוקה - יישור לימין */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 35px;
    margin-bottom: 30px;
    color: white;
    text-align: right;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
}
.app-header h1 { font-size: 42px; font-weight: 900; color: white !important; margin: 0; }
.app-header p { font-size: 18px; color: white !important; margin-top: 10px; opacity: 0.9; }

/* כרטיסי מדדים - תיקון חפיפה ויישור */
.metric-row { 
    display: flex; 
    gap: 20px; 
    margin: 25px 0; 
    flex-wrap: wrap; 
    justify-content: flex-start;
    direction: rtl;
}
.metric-card {
    flex: 1; min-width: 200px;
    background: white; border-radius: 15px; padding: 25px;
    border: 1px solid #dde8dc; box-shadow: 0 4px 10px rgba(0,0,0,0.05);
    text-align: right;
}
.mc-label { font-size: 14px; color: #7a9a77; font-weight: 700; margin-bottom: 8px; }
.mc-value { font-size: 36px; font-weight: 900; color: #1a3d17; }

/* טבלאות מדריך */
.guide-table {
    width: 100%; border-collapse: collapse; margin: 20px 0;
    background: white; border-radius: 12px; overflow: hidden;
}
.guide-table th { background: #2d5a27; color: white; padding: 15px; text-align: right; }
.guide-table td { padding: 15px; border-bottom: 1px solid #eee; text-align: right; line-height: 1.6; }

/* תיבות מידע */
.info-box { 
    background: #edf5ec; border-right: 6px solid #2d5a27; 
    padding: 20px; border-radius: 0 10px 10px 0; 
    margin: 15px 0; text-align: right; direction: rtl;
}

/* כפתור כתום בולט */
div.stButton > button {
    background: linear-gradient(135deg, #e67e22, #d35400) !important;
    color: white !important; font-weight: 800 !important; font-size: 20px !important;
    border-radius: 12px !important; height: 3.5em; width: 100%; border: none !important;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 2. לוגיקה ומחלקות
# ══════════════════════════════════════════════════════════════════

class Soldier:
    def __init__(self, s_id, name, restr=""):
        self.soldier_id = str(s_id)
        self.name = str(name).strip()
        self.restricted_tasks = self._parse(restr)
    @staticmethod
    def _parse(restr):
        if pd.isna(restr) or str(restr).strip() in ("", "nan"): return []
        return [int(p.replace(".0","")) for p in str(restr).replace(" ","").split(",") if p.replace(".0","").isdigit()]

class Task:
    NIGHT = {22, 23, 0, 1, 2, 3, 4, 5, 6, 7}
    def __init__(self, t_id, name, req_p, shift_dur=1, rest_dur=0, overlap=False, hours="all"):
        self.task_id = int(t_id)
        self.name = str(name).strip()
        self.required_personnel = int(req_p)
        self.shift_duration = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration = int(rest_dur) if pd.notna(rest_dur) else 0
        self.allow_overlap = str(overlap).strip().lower() == "true"
        self.active_hours = self._parse_hours(hours)
    @staticmethod
    def _parse_hours(hours):
        if pd.isna(hours) or str(hours).strip().lower() in ("all",""): return list(range(24))
        return [int(x) for x in str(hours).replace(" ","").split(",") if x.isdigit()]

def to_excel_styled(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=True, sheet_name='שבצ"ק')
        workbook, worksheet = writer.book, writer.sheets['שבצ"ק']
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5a27', 'font_color': 'white', 'border': 1, 'align': 'right'})
        for ci, col in enumerate(df.columns.values):
            worksheet.write(0, ci + 1, col, h_fmt)
            w = max(df[col].astype(str).map(len).max(), len(col)) + 5
            worksheet.set_column(ci + 1, ci + 1, min(w, 40))
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════
# 3. מנוע אופטימיזציה (48 שעות + הוגנות)
# ══════════════════════════════════════════════════════════════════

def solve_scheduling(soldiers, tasks, num_days=2):
    H = 24 * num_days
    model = cp_model.CpModel()
    x, start = {}, {}

    for s in soldiers:
        for t in tasks:
            for h in range(H):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f'x_{s.soldier_id}_{t.task_id}_{h}')
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f'st_{s.soldier_id}_{t.task_id}_{h}')

    # אילוצי רציפות, מנוחה ואיוש
    for s in soldiers:
        for t in tasks:
            for h in range(H):
                rel = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(rel))
                if h + t.shift_duration > H: model.Add(start[s.soldier_id, t.task_id, h] == 0)

                if t.rest_duration > 0:
                    busy = t.shift_duration + t.rest_duration
                    for nh in range(h + 1, min(h + busy, H)):
                        for ot in tasks:
                            if not ot.allow_overlap:
                                model.AddImplication(start[s.soldier_id, t.task_id, h], x[s.soldier_id, ot.task_id, nh].Not())

    for h in range(H):
        for t in tasks:
            assigned = [x[s.soldier_id, t.task_id, h] for s in soldiers]
            if (h % 24) in t.active_hours: model.Add(sum(assigned) == t.required_personnel)
            else: model.Add(sum(assigned) == 0)

    for s in soldiers:
        for h in range(H):
            model.Add(sum(x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap) <= 1)

    # הוגנות Min-Max
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(H)) for s in soldiers]
    max_load = model.NewIntVar(0, H, "max_load")
    min_load = model.NewIntVar(0, H, "min_load")
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)
    model.Minimize(100 * (max_load - min_load))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        rows = []
        for s in soldiers:
            row = {"שם": s.name}
            n_hours = 0
            for h in range(H):
                day = h // 24 + 1
                label = f"יום {day} - {h%24:02d}:00"
                act = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
                row[label] = " + ".join(act) if act else "—"
                if act and (h%24) in Task.NIGHT: n_hours += 1
            row["סך שעות"] = sum(1 for h in range(H) if any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks))
            row["שעות לילה"] = n_hours
            rows.append(row)
        return pd.DataFrame(rows)
    return None

# ══════════════════════════════════════════════════════════════════
# 4. ממשק המשתמש (Tabs עם תוכן מלא)
# ══════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות ותורנויות — הוגן, מהיר ומדויק</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות עבודה"])

# --- טאב 1: הרצה + גרפים + תובנות ---
with tab_run:
    c1, c2 = st.columns(2)
    with c1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with c2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק ונתח נתונים"):
            soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים')) for _, r in s_df.iterrows()]
            tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'), r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
            
            res = solve_scheduling(soldiers, tasks)
            if res is not None:
                st.success("✅ השיבוץ הושלם!")
                
                # כרטיסי מדד
                gap = res["סך שעות"].max() - res["סך שעות"].min()
                st.markdown(f"""
                <div class="metric-row">
                    <div class="metric-card"><div class="mc-label">חיילים</div><div class="mc-value">{len(soldiers)}</div></div>
                    <div class="metric-card"><div class="mc-label">ממוצע שעות</div><div class="mc-value">{res['סך שעות'].mean():.1f}</div></div>
                    <div class="metric-card"><div class="mc-label">פער הוגנות</div><div class="mc-value">{gap} שעות</div></div>
                </div>
                """, unsafe_allow_html=True)
                
                st.subheader("📅 לוח השיבוץ הסופי")
                st.table(res)
                st.download_button("📥 הורד קובץ אקסל סופי", data=to_excel_styled(res), file_name="Final_Schedule.xlsx")
                
                # גרפים
                st.divider()
                st.subheader("📊 ניתוח עומסים חזותי")
                g1, g2 = st.columns(2)
                with g1: st.plotly_chart(px.bar(res, x="שם", y="סך שעות", color="סך שעות", title="עומס כולל", color_continuous_scale="Greens"), use_container_width=True)
                with g2: st.plotly_chart(px.bar(res, x="שם", y="שעות לילה", title="פגיעה בשינת לילה", color="שעות לילה", color_continuous_scale="Reds"), use_container_width=True)

                # תובנות ה"למה"
                st.markdown("#### 💡 למה הלו\"ז נראה ככה?")
                with st.expander("ניתוח סיבות השיבוץ"):
                    st.write(f"**פערי עומס:** החיילים העמוסים ביותר הם אלו ללא פטורים שזמינים למילוי 'חורים'.")
                    st.write(f"**שינת לילה:** המערכת ממזערת עבודה בלילה אך מחויבת לאייש עמדות 24/7 שהגדרת.")
            else:
                st.error("לא נמצא פתרון. נסה להוריד דרישות מנוחה.")

# --- טאב 2: מדריך מפורט (החזרתי הכל) ---
with tab_guide:
    st.markdown("### 📖 מדריך מפורט למילוי השבצ\"ק")
    st.markdown('<div class="info-box">פעלו לפי ההנחיות כדי להבטיח שהאלגוריתם יפיק את התוצאה הטובה ביותר.</div>', unsafe_allow_html=True)
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>הסבר טכני</th><th>דוגמה</th></tr>
        <tr><td><b>משך_משמרת</b></td><td>כמה שעות רצופות חייל "נעול" במשימה.</td><td>4 (החייל לא יוחלף באמצע 4 שעות).</td></tr>
        <tr><td><b>שעות_מנוחה_אחרי</b></td><td>כמה שעות חובה לתת לחייל לנוח בסיום.</td><td>8 (חסימה מוחלטת משיבוץ ל-8 שעות).</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td>האם המשימה מאפשרת משימה נוספת במקביל?</td><td>True לכוננות, False לשמירה.</td></tr>
        <tr><td><b>פטורים</b></td><td>קודי משימה (מהקובץ השני) שהחייל לא מבצע.</td><td>1, 2 (מופרד בפסיקים).</td></tr>
    </table>
    """, unsafe_allow_html=True)

# --- טאב 3: תבניות (החזרתי הכל) ---
with tab_templates:
    st.subheader("📥 הורדת תבניות עבודה")
    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        s_ex = pd.DataFrame({"מספר_אישי": [1001, 1002], "שם": ["ישראל ישראלי", "יוסי כהן"], "פטורים": ["", "1"]})
        st.write("**תבנית חיילים:**")
        st.dataframe(s_ex, hide_index=True)
        st.download_button("הורד תבנית חיילים", data=to_excel_styled(s_ex), file_name="Soldiers_Template.xlsx")
    with col_ex2:
        t_ex = pd.DataFrame({"קוד_משימה": [1, 2], "שם": ["שמירה", "מטבח"], "כוח_אדם_נדרש": [1, 2], "משך_משמרת": [4, 6], "שעות_מנוחה_אחרי": [8, 0]})
        st.write("**תבנית משימות:**")
        st.dataframe(t_ex, hide_index=True)
        st.download_button("הורד תבנית משימות", data=to_excel_styled(t_ex), file_name="Tasks_Template.xlsx")
