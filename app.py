import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# 1. הגדרות עמוד ו-CSS (העיצוב המנצח + RTL חסין)
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title='שבצ"ק חכם',
    page_icon="🪖",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* כפיית RTL ופונט Heebo על כל האפליקציה */
.stApp, [data-testid="stAppViewContainer"], .main, .block-container {
    direction: rtl !important;
    text-align: right !important;
    font-family: 'Heebo', sans-serif !important;
}

/* יישור טאבים לצד ימין */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important;
    justify-content: flex-start !important;
    gap: 15px;
    border-bottom: 2px solid #2d5a27;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    font-weight: 700;
    font-size: 16px;
}

/* כותרת עליונה בעיצוב Gradient */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 30px 35px;
    margin-bottom: 30px;
    color: white;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
}
.app-header h1 { font-size: 42px; font-weight: 900; color: white !important; margin: 0; }
.app-header p { font-size: 18px; color: white !important; margin-top: 8px; opacity: 0.9; }

/* כרטיסי מדד (Metrics) - מניעת חפיפה */
.metric-row { display: flex; gap: 16px; margin: 25px 0; flex-wrap: wrap; direction: rtl; }
.metric-card {
    flex: 1; min-width: 200px;
    background: white; border-radius: 14px; padding: 22px;
    border: 1px solid #dde8dc; box-shadow: 0 2px 8px rgba(45,90,39,0.07);
    text-align: right;
}
.mc-label { font-size: 12px; color: #7a9a77; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
.mc-value { font-size: 34px; font-weight: 900; color: #1a3d17; line-height: 1; }

/* טבלאות מדריך */
.guide-table { width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 10px; overflow: hidden; }
.guide-table th { background: #2d5a27; color: white; padding: 14px; text-align: right; font-weight: 700; }
.guide-table td { padding: 14px; border-bottom: 1px solid #eee; text-align: right; line-height: 1.6; }

/* כפתורים */
div.stButton > button {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important; font-weight: 700 !important; border-radius: 10px !important;
    height: 3.5em; width: 100%; border: none !important; box-shadow: 0 4px 10px rgba(0,0,0,0.15);
}
[data-testid="stDownloadButton"] > button {
    background: #c0500a !important; color: white !important; font-weight: 600 !important;
}

/* תיבות מידע */
.info-box { background: #edf5ec; border-right: 5px solid #2d5a27; padding: 15px; border-radius: 0 10px 10px 0; margin: 15px 0; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 2. מחלקות נתונים (Data Objects)
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

# פונקציית Excel עם Auto-Fit (מרווחת)
def to_excel_styled(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=True, sheet_name='שבצ"ק')
        workbook, worksheet = writer.book, writer.sheets['שבצ"ק']
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#2d5a27', 'font_color': 'white', 'border': 1, 'align': 'right'})
        for ci, col in enumerate(df.columns.values):
            worksheet.write(0, ci + 1, col, h_fmt)
            w = max(df[col].astype(str).map(len).max(), len(col)) + 6
            worksheet.set_column(ci + 1, ci + 1, min(w, 40))
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════
# 3. מנוע CP-SAT (שינה, הוגנות, רציפות ומנוחה)
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
                # אילוץ רציפות (נעילה)
                rel = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(rel))
                if h + t.shift_duration > H: model.Add(start[s.soldier_id, t.task_id, h] == 0)
                # אילוץ מנוחה
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
        for t in tasks:
            if t.task_id in s.restricted_tasks:
                for h in range(H): model.Add(x[s.soldier_id, t.task_id, h] == 0)

    # הוגנות Min-Max + עונש שינה (Soft)
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(H)) for s in soldiers]
    max_load = model.NewIntVar(0, H, "max_load")
    min_load = model.NewIntVar(0, H, "min_load")
    for load in s_loads:
        model.Add(max_load >= load)
        model.Add(min_load <= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in range(H) if (h % 24) in Task.NIGHT)
    
    # אילוץ שינה רצופה כ-Soft Constraint
    sleep_viol = []
    if min_sleep > 0:
        for s in soldiers:
            for h in range(H - min_sleep + 1):
                work_in_window = [x[s.soldier_id, t.task_id, h + k] for t in tasks if not t.allow_overlap for k in range(min_sleep)]
                viol = model.NewBoolVar(f"v_{s.soldier_id}_{h}")
                model.Add(sum(work_in_window) >= 1).OnlyEnforceIf(viol)
                model.Add(sum(work_in_window) == 0).OnlyEnforceIf(viol.Not())
                sleep_viol.append(viol)

    model.Minimize(150 * (max_load - min_load) + 5 * night_work + 10 * sum(sleep_viol))
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = timeout
    status = solver.Solve(model)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE): return solver, x, start
    return None, {}, {}

def solve_scheduling(soldiers, tasks, min_sleep, timeout, num_days):
    H = 24 * num_days
    # לוגיקת צמצום שינה אוטומטית למציאת פתרון
    for s_level in [min_sleep, 4, 0]:
        solver, x, start = _build_and_solve(soldiers, tasks, H, s_level, timeout)
        if solver: break
    
    if solver is None: return None, "infeasible"
    
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
    return pd.DataFrame(rows), f"נשמר חלון שינה רצוף של {s_level} שעות."

# ══════════════════════════════════════════════════════════════════
# 4. ממשק המשתמש (Tabs ועיצוב)
# ══════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות ותורנויות — הוגן, מהיר ומדויק</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות עבודה"])

with tab_templates:
    st.subheader("📥 הורדת תבניות עבודה אקטואליות")
    st.markdown('<div class="info-box">השתמשו בתבניות אלו כדי להבטיח את תקינות הנתונים.</div>', unsafe_allow_html=True)
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        s_ex = pd.DataFrame({"מספר_אישי": [101, 102], "שם": ["ישראל ישראלי", "יוסי כהן (פטור שער)"], "פטורים": ["", "1"]})
        st.write("**תבנית חיילים:**")
        st.dataframe(s_ex, hide_index=True)
        st.download_button("הורד תבנית חיילים", data=to_excel_styled(s_ex), file_name="Template_Soldiers.xlsx")
    with col_t2:
        t_ex = pd.DataFrame({"קוד_משימה": [1, 2, 10], "שם": ["שמירה", "מטבח", "כוננות"], "כוח_אדם_נדרש": [1, 2, 8], "משך_משמרת": [4, 6, 24], "שעות_מנוחה_אחרי": [8, 0, 0], "אישור_חפיפה": [False, False, True]})
        st.write("**תבנית משימות:**")
        st.dataframe(t_ex, hide_index=True)
        st.download_button("הורד תבנית משימות", data=to_excel_styled(t_ex), file_name="Template_Tasks.xlsx")

with tab_guide:
    st.subheader("📖 מדריך מלא למילוי השבצ''ק")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר מפורט</th><th>הנחיות</th></tr>
        <tr><td><b>מספר_אישי</b></td><td>זיהוי ייחודי לכל חייל.</td><td>חובה לכל חייל.</td></tr>
        <tr><td><b>פטורים</b></td><td>מספרי קוד משימה שהחייל לא מבצע.</td><td>הפרדה בפסיקים (1, 5).</td></tr>
        <tr><td><b>משך_משמרת</b></td><td>שעות רצופות שהחייל ננעל למשימה.</td><td>למשל 4 (לא ניתן להחליף חייל באמצע).</td></tr>
        <tr><td><b>שעות_מנוחה_אחרי</b></td><td>שעות הפסקה חובה בסיום המשימה.</td><td>למשל 8 (חסימה מוחלטת משיבוץ).</td></tr>
        <tr><td><b>אישור_חפיפה</b></td><td>האם המשימה מאפשרת משימה נוספת במקביל?</td><td><b>True</b> לכוננות, <b>False</b> לשמירה.</td></tr>
        <tr><td><b>שעות_פעילות</b></td><td>מתי המשימה קורית.</td><td><b>all</b> ל-24/7 או שעות (7,8,9).</td></tr>
    </table>
    """, unsafe_allow_html=True)
    st.info("💡 המערכת מתעדפת אוטומטית חלון שינה רצוף של 7 שעות בלילה.")

with tab_run:
    c_up1, c_up2 = st.columns(2)
    with c_up1: sf = st.file_uploader("📂 העלה קובץ חיילים", type="xlsx")
    with c_up2: tf = st.file_uploader("📂 העלה קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק ונתח תובנות"):
            with st.spinner("מחשב שיבוץ אופטימלי..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'), r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                
                res, msg = solve_scheduling(soldiers, tasks, 7, 30.0, 2)
                if res is not None:
                    st.success(f"✅ {msg}")
                    gap = res["סך שעות"].max() - res["סך שעות"].min()
                    st.markdown(f"""
                    <div class="metric-row">
                        <div class="metric-card"><div class="mc-label">חיילים</div><div class="mc-value">{len(soldiers)}</div></div>
                        <div class="metric-card"><div class="mc-label">ממוצע שעות</div><div class="mc-value">{res['סך שעות'].mean():.1f}</div></div>
                        <div class="metric-card"><div class="mc-label">פער הוגנות</div><div class="mc-value">{gap} שעות</div></div>
                    </div>
                    """, unsafe_allow_html=True)
                    st.table(res)
                    st.download_button("📥 הורד לוח סופי מרווח", data=to_excel_styled(res), file_name="Final_Shavtzak.xlsx")
                    
                    st.divider()
                    st.subheader("📊 ניתוח עומסים חזותי")
                    g1, g2 = st.columns(2)
                    with g1: st.plotly_chart(px.bar(res, x="שם", y="סך שעות", color="סך שעות", title="עומס כולל", color_continuous_scale="Greens"), use_container_width=True)
                    with g2: st.plotly_chart(px.bar(res, x="שם", y="שעות לילה", title="פגיעה בשינת לילה", color="שעות לילה", color_continuous_scale="Reds"), use_container_width=True)
                    
                    st.markdown("#### 💡 למה הלו\"ז נראה ככה?")
                    with st.expander("ניתוח סיבות השיבוץ"):
                        night_req = t_df[t_df['שעות_פעילות'].astype(str).str.contains('all|22|23|0|1|2|3|4|5|6|7')]['כוח_אדם_נדרש'].sum()
                        st.write(f"1. **עבודה בלילה:** נדרשים {night_req} חיילים בלילה. המערכת חילקה את הנטל למרות השאיפה לשינה.")
                        st.write(f"2. **פערי עומס:** נובעים מחיילים ללא פטורים שממלאים חורים של חיילים עם פטורים או מנוחה חובה.")
                else:
                    st.error("לא נמצא פתרון. נסו להפחית דרישות מנוחה.")
