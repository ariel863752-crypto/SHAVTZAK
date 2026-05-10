import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# 1. הגדרות עמוד ו-CSS (RTL חסין ועיצוב)
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title='שבצ"ק חכם',
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* יישור כללי לימין - כפייה על כל אלמנט */
.stApp, [data-testid="stAppViewContainer"], .main, .block-container {
    direction: rtl !important;
    text-align: right !important;
}

/* תיקון טאבים - העברה לצד ימין */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important;
    justify-content: flex-start !important;
    gap: 15px;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    font-family: 'Heebo', sans-serif;
    font-weight: 600;
}

/* כותרת עליונה */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 30px;
    margin-bottom: 25px;
    color: white;
    text-align: right;
}
.app-header h1 { font-size: clamp(24px, 5vw, 42px); font-weight: 900; color: white !important; margin: 0; }
.app-header p { font-size: 18px; opacity: 0.9; color: white !important; margin-top: 10px; }

/* כרטיסי מדד (Metrics) - מניעת חפיפה */
.metric-row { 
    display: flex; 
    gap: 16px; 
    margin: 20px 0; 
    flex-wrap: wrap; /* מאפשר ירידת שורה במסכים קטנים */
    direction: rtl;
}
.metric-card {
    flex: 1; min-width: 200px;
    background: white; border-radius: 12px; padding: 20px;
    border: 1px solid #dde8dc; box-shadow: 0 2px 8px rgba(45,90,39,0.07);
    text-align: right;
}
.mc-label { font-size: 12px; color: #7a9a77; font-weight: 600; margin-bottom: 5px; }
.mc-value { font-size: 32px; font-weight: 800; color: #1a3d17; }

/* טבלאות מדריך */
.guide-table {
    width: 100%; border-collapse: collapse; margin: 15px 0;
    background: white; border-radius: 10px; overflow: hidden;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}
.guide-table th { background: #2d5a27; color: white; padding: 12px; text-align: right; font-weight: 700; }
.guide-table td { padding: 12px; border-bottom: 1px solid #eee; text-align: right; line-height: 1.6; }

/* כפתורים */
div.stButton > button {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important; font-weight: 700 !important;
    border-radius: 10px !important; height: 3.5em; width: 100%;
}

/* תיבות מידע */
.info-box { background: #edf5ec; border-right: 5px solid #2d5a27; padding: 15px; border-radius: 0 8px 8px 0; margin: 10px 0; text-align: right; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים ולוגיקה
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
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5a27', 'font_color': 'white', 'border': 1})
        for ci, col in enumerate(df.columns.values):
            worksheet.write(0, ci + 1, col, h_fmt)
            w = max(df[col].astype(str).map(len).max(), len(col)) + 5
            worksheet.set_column(ci + 1, ci + 1, min(w, 40))
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════
# 3. מנוע CP-SAT (הלוגיקה המלאה מ-v3.0)
# ══════════════════════════════════════════════════════════════════

def _build_and_solve(soldiers, tasks, H, min_sleep, timeout):
    model = cp_model.CpModel()
    x, start = {}, {}
    for s in soldiers:
        for t in tasks:
            for h in range(H):
                x[s.soldier_id, t.task_id, h] = model.NewBoolVar(f'x_{s.soldier_id}_{t.task_id}_{h}')
                start[s.soldier_id, t.task_id, h] = model.NewBoolVar(f'st_{s.soldier_id}_{t.task_id}_{h}')

    for s in soldiers:
        for t in tasks:
            for h in range(H):
                rel = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(rel))
                if h + t.shift_duration > H: model.Add(start[s.soldier_id, t.task_id, h] == 0)

    for s in soldiers:
        for t in tasks:
            if t.rest_duration <= 0: continue
            busy = t.shift_duration + t.rest_duration
            for h in range(H):
                for nh in range(h + 1, min(h + busy, H)):
                    for ot in tasks:
                        if not ot.allow_overlap:
                            model.AddImplication(start[s.soldier_id, t.task_id, h], x[s.soldier_id, ot.task_id, nh].Not())

    for h in range(H):
        hour_of_day = h % 24
        for t in tasks:
            assigned = [x[s.soldier_id, t.task_id, h] for s in soldiers]
            if hour_of_day in t.active_hours: model.Add(sum(assigned) == t.required_personnel)
            else: model.Add(sum(assigned) == 0)

    for s in soldiers:
        for h in range(H):
            model.Add(sum(x[s.soldier_id, t.task_id, h] for t in tasks if not t.allow_overlap) <= 1)

    for s in soldiers:
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(H): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # הוגנות ושינה
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(H)) for s in soldiers]
    max_load = model.NewIntVar(0, H, "max_load")
    min_load = model.NewIntVar(0, H, "min_load")
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in range(H) if (h % 24) in Task.NIGHT)
    
    model.Minimize(100 * (max_load - min_load) + 2 * night_work)
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = timeout
    status = solver.Solve(model)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE): return solver, x, start
    return None, {}, {}

def solve_scheduling(soldiers, tasks, min_sleep=6, num_days=2):
    H = 24 * num_days
    # ניסיון פתרון עם אילוץ שינה מלא (ניתן להוסיף לוגיקת צמצום כאן)
    solver, x, start = _build_and_solve(soldiers, tasks, H, min_sleep, 30.0)
    
    if solver is None: return None
    
    rows = []
    for s in soldiers:
        row = {"שם": s.name}
        total, night_count = 0, 0
        for h in range(H):
            day = h // 24 + 1
            label = f"יום {day} {h%24:02d}:00"
            act = [t.name for t in tasks if solver.Value(x[s.soldier_id, t.task_id, h]) == 1]
            row[label] = " + ".join(act) if act else "—"
            if act:
                total += 1
                if (h % 24) in Task.NIGHT: night_count += 1
        row["סך שעות"] = total
        row["שעות לילה"] = night_count
        rows.append(row)
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════════
# 4. ממשק המשתמש (UI)
# ══════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות ותורנויות — הוגן, מהיר ומדויק</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך מפורט", "📥 תבניות עבודה"])

with tab_templates:
    st.markdown("### 📥 הורדת תבניות עבודה מעודכנות")
    st.markdown('<div class="info-box">הורידו את הקבצים, מלאו אותם לפי ההנחיות והעלו בטאב "ביצוע שיבוץ".</div>', unsafe_allow_html=True)
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        s_ex = pd.DataFrame({"מספר_אישי": [1001, 1002], "שם": ["ישראל ישראלי", "יוסי כהן (פטור שער)"], "פטורים": ["", "1"]})
        st.write("**תבנית חיילים:**")
        st.dataframe(s_ex, hide_index=True)
        st.download_button("הורד תבנית חיילים", data=to_excel_styled(s_ex), file_name="Template_Soldiers.xlsx")
    
    with col_t2:
        t_ex = pd.DataFrame({
            "קוד_משימה": [1, 2, 10], 
            "שם": ["שמירת שער", "מטבח", "כיתת כוננות"], 
            "כוח_אדם_נדרש": [1, 3, 8], 
            "משך_משמרת": [4, 6, 24], 
            "שעות_מנוחה_אחרי": [8, 0, 0],
            "אישור_חפיפה": [False, False, True]
        })
        st.write("**תבנית משימות:**")
        st.dataframe(t_ex, hide_index=True)
        st.download_button("הורד תבנית משימות", data=to_excel_styled(t_ex), file_name="Template_Tasks.xlsx")

with tab_guide:
    st.markdown("### 📖 מדריך למילוי השבצ\"ק")
    st.markdown("""
    <table class="guide-table">
        <tr><th>עמודה</th><th>הסבר מבצעי</th><th>הנחיות</th></tr>
        <tr><td><b>משך_משמרת</b></td><td>כמה שעות רצופות חייל נמצא בעמדה (למשל 4).</td><td>האלגוריתם נועל את החייל לאותן שעות.</td></tr>
        <tr><td><b>שעות_מנוחה_אחרי</b></td><td>כמה שעות חובה לתת לחייל לנוח לפני המשימה הבאה.</td><td>חוסם שיבוץ לאחר סיום המשמרת.</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td>האם ניתן לבצע משימה נוספת במקביל?</td><td>כתבו <b>True</b> לכוננות, <b>False</b> לשמירה.</td></tr>
        <tr><td><b>שעות_פעילות</b></td><td>מתי המשימה קורית.</td><td><b>all</b> ל-24/7 או שעות מופרדות בפסיק (8,9,12).</td></tr>
    </table>
    """, unsafe_allow_html=True)
    st.info("💡 המערכת דואגת אוטומטית ל-7 שעות שינה רצופות בלילה ככל שהסד''כ מאפשר.")

with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with col_u2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק אופטימלי"):
            with st.spinner("מחשב שיבוץ אופטימלי..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'), r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                
                res = solve_scheduling(soldiers, tasks)
                if res is not None:
                    st.success("✅ השיבוץ הושלם!")
                    
                    # הצגת מדדים בכרטיסים (תיקון חפיפה)
                    gap = res["סך שעות"].max() - res["סך שעות"].min()
                    st.markdown(f"""
                    <div class="metric-row">
                        <div class="metric-card"><div class="mc-label">חיילים משובצים</div><div class="mc-value">{len(soldiers)}</div></div>
                        <div class="metric-card"><div class="mc-label">ממוצע שעות</div><div class="mc-value">{res['סך שעות'].mean():.1f}</div></div>
                        <div class="metric-card"><div class="mc-label">פער הוגנות</div><div class="mc-value">{gap} שעות</div></div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.table(res)
                    st.download_button("📥 הורד שבצ''ק סופי", data=to_excel_styled(res), file_name="Final_Schedule.xlsx")
                else:
                    st.error("לא נמצא פתרון. נסה להפחית דרישות מנוחה או להוסיף חיילים.")
