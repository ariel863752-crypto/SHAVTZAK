import io
import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model

# ══════════════════════════════════════════════════════════════════
# 1. עיצוב ממשק (UI/UX) - העיצוב המנצח עם RTL חסין
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

/* בסיס RTL ופונט Heebo */
.stApp, [data-testid="stAppViewContainer"], .main, .block-container {
    direction: rtl !important;
    text-align: right !important;
    font-family: 'Heebo', sans-serif !important;
}

/* כותרת Gradient ירוקה */
.app-header {
    background: linear-gradient(135deg, #1a3d17 0%, #2d5a27 60%, #3d7a35 100%);
    border-radius: 16px;
    padding: 30px 35px;
    margin-bottom: 30px;
    color: white;
    text-align: right;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
}
.app-header h1 { font-size: 42px; font-weight: 900; color: white !important; margin: 0; }
.app-header p { font-size: 18px; color: white !important; margin-top: 8px; opacity: 0.9; }

/* טאבים מיושרים לימין */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    flex-direction: row-reverse !important;
    justify-content: flex-start !important;
    gap: 15px;
    border-bottom: 2px solid #2d5a27;
}
[data-testid="stTabs"] [data-baseweb="tab"] { font-weight: 700; font-size: 16px; }

/* כרטיסי מדד (Metrics) */
.metric-row { display: flex; gap: 16px; margin: 25px 0; flex-wrap: wrap; direction: rtl; }
.metric-card {
    flex: 1; min-width: 200px;
    background: white; border-radius: 14px; padding: 22px;
    border: 1px solid #dde8dc; box-shadow: 0 2px 8px rgba(45,90,39,0.07);
    text-align: right;
}
.mc-label { font-size: 12px; color: #7a9a77; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
.mc-value { font-size: 34px; font-weight: 900; color: #1a3d17; line-height: 1; }

/* כפתורים */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #2d5a27, #3d7a35) !important;
    color: white !important; font-weight: 700 !important; border-radius: 10px !important;
    height: 3.5em; width: 100%; border: none !important; box-shadow: 0 4px 10px rgba(0,0,0,0.15);
}
[data-testid="stDownloadButton"] > button {
    background: #c0500a !important; color: white !important; font-weight: 600 !important; border-radius: 10px !important;
}

/* טבלאות מדריך */
.guide-table { width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 10px; overflow: hidden; }
.guide-table th { background: #2d5a27; color: white; padding: 14px; text-align: right; font-weight: 700; }
.guide-table td { padding: 14px; border-bottom: 1px solid #eee; text-align: right; line-height: 1.6; }

/* תיבות מידע */
.info-box { background: #edf5ec; border-right: 5px solid #2d5a27; padding: 15px; border-radius: 0 10px 10px 0; margin: 15px 0; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 2. מחלקות הנתונים (Objects) - כולל תפקידים ופטורים
# ══════════════════════════════════════════════════════════════════
class Soldier:
    def __init__(self, s_id, name, restr="", roles=""):
        self.soldier_id = str(s_id)
        self.name = name
        # פטורים (קודי משימות)
        self.restricted_tasks = [int(float(t)) for t in str(restr).split(',') if
                                 str(t).strip().replace('.0', '').isdigit()] if pd.notna(restr) else []
        # תפקידים (נהג, קצין וכו')
        self.roles = [r.strip() for r in str(roles).split(',') if r.strip()] if pd.notna(roles) else []

class Task:
    def __init__(self, t_id, name, req_p, shift_dur, rest_dur=0, overlap=False, hours="all", req_role=""):
        self.task_id = int(t_id)
        self.name = name
        self.required_personnel = int(req_p)
        self.shift_duration = int(shift_dur) if pd.notna(shift_dur) else 1
        self.rest_duration = int(rest_dur) if pd.notna(rest_dur) else 0
        self.allow_overlap = str(overlap).lower() == 'true'
        self.required_role = str(req_role).strip() if pd.notna(req_role) and str(req_role).strip() != "" else None
        
        if pd.isna(hours) or str(hours).lower() in ['all', '']:
            self.active_hours = list(range(24))
        else:
            self.active_hours = [int(x.strip()) for x in str(hours).split(',') if str(x).strip().isdigit()]

# ══════════════════════════════════════════════════════════════════
# 3. פונקציית Excel עם Auto-Fit (מרווחת)
# ══════════════════════════════════════════════════════════════════
def to_excel_styled(df, sheet_name='שבצ"ק', include_index=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=include_index, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bold': True, 'fg_color': '#2d5a27', 'font_color': 'white', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            col_idx = col_num + (1 if include_index else 0)
            worksheet.write(0, col_idx, value, header_format)
            column_len = max(df[value].astype(str).map(len).max(), len(value)) + 5
            worksheet.set_column(col_idx, col_idx, column_len)
        if include_index: worksheet.set_column(0, 0, 5)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════
# 4. המנוע המתמטי - כולל תפקידים, שינה, מנוחה והוגנות
# ══════════════════════════════════════════════════════════════════
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
            # אילוץ תפקיד נדרש (Role Constraint)
            if t.required_role and t.required_role not in s.roles:
                for h in range(num_hours):
                    model.Add(x[s.soldier_id, t.task_id, h] == 0)

            for h in range(num_hours):
                # אילוץ רציפות (נעילה)
                relevant_starts = [start[s.soldier_id, t.task_id, i] for i in range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))

                # אילוץ מנוחה אחרי משימה
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

    night_hours = [22, 23, 0, 1, 2, 3, 4, 5, 6, 7]
    s_loads = [sum(x[s.soldier_id, t.task_id, h] for t in tasks for h in range(num_hours)) for s in soldiers]
    max_load = model.NewIntVar(0, num_hours, 'max_load')
    for load in s_loads: model.Add(max_load >= load)
    
    night_work = sum(x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in night_hours)
    model.Minimize(40 * max_load + night_work)

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

# ══════════════════════════════════════════════════════════════════
# 5. ממשק המשתמש (Tabs)
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <h1>🪖 שבצ"ק — מערכת שיבוץ כוחות חכמה</h1>
  <p>אופטימיזציה אוטומטית של שמירות ותורנויות — כולל ניהול תפקידים וכישורים</p>
</div>
""", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות אקסל"])

with tab_templates:
    st.subheader("📥 הורדת תבניות עבודה מרווחות")
    s_ex = pd.DataFrame({'מספר_אישי': [1001, 1002], 'שם': ['ישראל ישראלי', 'יוסי כהן'], 'פטורים': ['', '1'], 'תפקידים': ['נהג, מפקד', 'קלע']})
    t_ex = pd.DataFrame({
        'קוד_משימה': [1, 2, 10], 'שם': ['שמירת שער', 'סיור רכוב', 'כוננות'], 
        'כוח_אדם_נדרש': [1, 1, 8], 'משך_משמרת': [4, 4, 24], 
        'שעות_מנוחה_אחרי': [8, 4, 0], 'אישור_חפיפה': [False, False, True], 
        'שעות_פעילות': ['all', 'all', 'all'], 'תפקיד_נדרש': ['', 'נהג', '']
    })
    c1, c2 = st.columns(2)
    with c1: st.download_button("הורד תבנית חיילים", data=to_excel_styled(s_ex, "Soldiers", False), file_name="Shavtzak_Soldiers.xlsx")
    with c2: st.download_button("הורד תבנית משימות", data=to_excel_styled(t_ex, "Tasks", False), file_name="Shavtzak_Tasks.xlsx")

with tab_guide:
    st.subheader("📖 מדריך מלא - ניהול תפקידים וכישורים")
    st.markdown("### 👥 קובץ חיילים (Soldiers.xlsx)")
    st.markdown("""<table class="guide-table">
        <tr><th>עמודה</th><th>הסבר מפורט</th></tr>
        <tr><td><b>תפקידים</b></td><td>רשימת ההכשרות של החייל (נהג, חובש וכו'). המערכת תבדוק עמודה זו מול דרישת המשימה.</td></tr>
        <tr><td><b>פטורים</b></td><td>קודי משימות שהחייל חסום אליהן ללא קשר לתפקידו.</td></tr>
    </table>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 📋 קובץ משימות (Tasks.xlsx)")
    st.markdown("""<table class="guide-table">
        <tr><th>עמודה</th><th>הסבר מפורט</th></tr>
        <tr><td><b>תפקיד_נדרש</b></td><td>התפקיד שחייב להיות לחייל כדי לבצע משימה זו (למשל <b>נהג</b>). אם ריק - כולם יכולים.</td></tr>
        <tr><td><b>משך_משמרת</b></td><td>שעות רצופות שהחייל "נעול" במשימה.</td></tr>
    </table>""", unsafe_allow_html=True)

with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1: sf = st.file_uploader("📂 קובץ חיילים", type="xlsx")
    with col_u2: tf = st.file_uploader("📂 קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק ונתח תובנות"):
            with st.spinner("מחשב שיבוץ אופטימלי..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים'), r.get('תפקידים')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'), r.get('אישור_חפיפה'), r.get('שעות_פעילות'), r.get('תפקיד_נדרש')) for _, r in t_df.iterrows()]
                
                needed_roles = {t.required_role for t in tasks if t.required_role}
                available_roles = {role for s in soldiers for role in s.roles}
                missing = needed_roles - available_roles
                
                if missing:
                    st.error(f"❌ שגיאה: חסרים חיילים עם התפקידים: {', '.join(missing)}.")
                else:
                    final_df = solve_scheduling(soldiers, tasks)
                    if final_df is not None:
                        gap = final_df["סך שעות"].max() - final_df["סך שעות"].min()
                        st.markdown(f"""<div class="metric-row">
                            <div class="metric-card"><div class="mc-label">חיילים</div><div class="mc-value">{len(soldiers)}</div></div>
                            <div class="metric-card"><div class="mc-label">ממוצע שעות</div><div class="mc-value">{final_df['סך שעות'].mean():.1f}</div></div>
                            <div class="metric-card"><div class="mc-label">פער הוגנות</div><div class="mc-value">{gap} שעות</div></div>
                        </div>""", unsafe_allow_html=True)
                        st.table(final_df)
                        st.download_button("📥 הורד לוח סופי מרווח", data=to_excel_styled(final_df), file_name="Final_Shavtzak.xlsx", use_container_width=True)
                        st.plotly_chart(px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="חלוקת עומס כללית", color_continuous_scale="Greens"), use_container_width=True)
                        with st.expander("💡 למה הלו\"ז נראה ככה?"):
                            st.write(f"**התאמת תפקידים:** המערכת וידאה שרק בעלי הכשרה מתאימה שובצו למשימות הנדרשות.")
                            st.write(f"**פערי עומס:** נובעים מחיילים עם פטורים או מנוחה ארוכה שמוציאים אותם מהסבב.")
                    else:
                        st.error("לא נמצא פתרון. נסה להוסיף חיילים או להפחית מנוחות.")
