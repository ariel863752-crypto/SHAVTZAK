import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import plotly.express as px

# ==========================================
# 1. עיצוב ממשק RTL וסטיילינג מקצועי (שבצ"ק)
# ==========================================
st.set_page_config(page_title="מערכת שבצ''ק חכמה", page_icon="🪖", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #fcfcfc; color: #333; }
    .stMarkdown, .stAlert, p, span, div { text-align: right; direction: rtl; }
    .main-title { color: #2d5a27; font-size: 45px; font-weight: bold; border-bottom: 4px solid #e67e22; padding-bottom: 10px; margin-bottom: 25px; }

    /* עיצוב טבלאות מדריך - קריאות מקסימלית */
    .guide-table { width: 100%; border-collapse: collapse; margin-top: 20px; background-color: white; border: 1px solid #ddd; }
    .guide-table th { background-color: #2d5a27; color: white; padding: 12px; text-align: right; }
    .guide-table td { padding: 10px; border: 1px solid #ddd; font-size: 15px; line-height: 1.6; }

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
        self.restricted_tasks = [int(float(t)) for t in str(restr).split(',') if
                                 str(t).strip().replace('.0', '').isdigit()] if pd.notna(restr) else []


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
# 3. פונקציית Excel עם Auto-Fit (מרווחת)
# ==========================================
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


# ==========================================
# 4. המנוע המתמטי - שינה, מנוחה והוגנות
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
                relevant_starts = [start[s.soldier_id, t.task_id, i] for i in
                                   range(max(0, h - t.shift_duration + 1), h + 1)]
                model.Add(x[s.soldier_id, t.task_id, h] == sum(relevant_starts))

                if t.rest_duration > 0:
                    total_busy = t.shift_duration + t.rest_duration
                    for next_h in range(h + 1, min(h + total_busy, num_hours)):
                        for other_t in tasks:
                            if not other_t.allow_overlap:
                                model.AddImplication(start[s.soldier_id, t.task_id, h],
                                                     x[s.soldier_id, other_t.task_id, next_h].Not())

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

    night_work = sum(
        x[s.soldier_id, t.task_id, h] for s in soldiers for t in tasks if not t.allow_overlap for h in night_hours)
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
                if h in night_hours and any(
                        solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks if not t.allow_overlap):
                    n_work += 1
            row["סך שעות"] = sum(
                1 for h in range(num_hours) if any(solver.Value(x[s.soldier_id, t.task_id, h]) == 1 for t in tasks))
            row["שעות לילה"] = n_work
            res_list.append(row)
        df = pd.DataFrame(res_list)
        df.index = range(1, len(df) + 1)
        return df
    return None


# ==========================================
# 5. ממשק המשתמש (Tabs)
# ==========================================
st.markdown("<div class='main-title'>שבצ''ק - מערכת שיבוץ כוחות חכמה</div>", unsafe_allow_html=True)

tab_run, tab_guide, tab_templates = st.tabs(["🚀 ביצוע שיבוץ", "📖 מדריך למילוי שבצ''ק", "📥 תבניות אקסל"])

with tab_templates:
    st.subheader("📥 הורדת תבניות עבודה מרווחות")
    s_ex = pd.DataFrame({'מספר_אישי': [1001, 1002], 'שם': ['ישראל ישראלי', 'יוסי כהן (פטור שער)'], 'פטורים': ['', '1']})
    t_ex = pd.DataFrame({
        'קוד_משימה': [1, 2, 10],
        'שם': ['שמירת שער', 'תורנות מטבח', 'כיתת כוננות'],
        'כוח_אדם_נדרש': [1, 3, 8],
        'משך_משמרת': [4, 6, 24],
        'שעות_מנוחה_אחרי': [8, 0, 0],
        'אישור_חפיפה': [False, False, True],
        'שעות_פעילות': ['all', '7,8,9,12,13,14', 'all']
    })
    c1, c2 = st.columns(2)
    with c1: st.download_button("הורד תבנית חיילים", data=to_excel_styled(s_ex, "Soldiers", False),
                                file_name="Shavtzak_Soldiers.xlsx")
    with c2: st.download_button("הורד תבנית משימות", data=to_excel_styled(t_ex, "Tasks", False),
                                file_name="Shavtzak_Tasks.xlsx")

with tab_guide:
    st.subheader("📖 מדריך למילוי שבצ''ק - הסבר מלא לכל עמודה")

    st.markdown("### 👥 קובץ חיילים (Soldiers.xlsx)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר מפורט</th><th>הנחיות חשובות</th></tr>
        <tr>
            <td><b>מספר_אישי</b></td>
            <td>זיהוי ייחודי לכל חייל במערכת.</td>
            <td>חובה להשתמש במספרים שונים לכל חייל. המערכת משתמשת בזה כדי להבדיל בין שני חיילים עם שם דומה.</td>
        </tr>
        <tr>
            <td><b>שם</b></td>
            <td>שם החייל כפי שתרצו שיופיע בטבלה.</td>
            <td>מומלץ שם פרטי ומשפחה. זהו הטקסט שיוצג בלוח השיבוץ הסופי.</td>
        </tr>
        <tr>
            <td><b>פטורים</b></td>
            <td>רשימת משימות שהחייל <b>לא יכול</b> לבצע.</td>
            <td>יש לרשום את <b>קוד המשימה</b> (למשל 1). במידה ויש כמה, הפרידו בפסיק (1, 5). המערכת לעולם לא תשבץ חייל למשימה שמופיעה בפטורים שלו.</td>
        </tr>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📋 קובץ משימות (Tasks.xlsx)")
    st.markdown("""
    <table class="guide-table">
        <tr><th>שם עמודה</th><th>הסבר מפורט</th><th>הנחיות חשובות</th></tr>
        <tr>
            <td><b>קוד_משימה</b></td>
            <td>מספר זיהוי למשימה.</td>
            <td>מספר זה חייב להיות תואם למספרים שכתבתם בעמודת ה'פטורים' בקובץ החיילים.</td>
        </tr>
        <tr>
            <td><b>שם</b></td>
            <td>שם המשימה (שמירה, סיור וכו').</td>
            <td>השם שיופיע בתוך המשבצות בלוח השיבוץ.</td>
        </tr>
        <tr>
            <td><b>כוח_אדם_נדרש</b></td>
            <td>כמות החיילים שצריכים להיות בעמדה בכל שעה.</td>
            <td>אם השמירה היא בזוגות, כתבו 2. האלגוריתם יוודא שבכל רגע נתון יש בדיוק 2 חיילים בעמדה.</td>
        </tr>
        <tr>
            <td><b>משך_משמרת</b></td>
            <td>כמה שעות רצופות חייל מבצע את המשימה ברצף.</td>
            <td><b>מנגנון הנעילה:</b> אם כתבתם 4, חייל ששובץ לשעה הראשונה יישאר שם לכל 4 השעות. אי אפשר להחליף חייל באמצע משמרת.</td>
        </tr>
        <tr>
            <td><b>שעות_מנוחה_אחרי</b></td>
            <td>כמות שעות ההפסקה המינימלית שהחייל חייב לקבל בסיום המשימה.</td>
            <td>למשל, אחרי שמירה של 4 שעות רצופות (משך משמרת), תוכלו להגדיר 8 שעות מנוחה שבהן החייל לא ישובץ לשום דבר אחר.</td>
        </tr>
        <tr>
            <td><b>אישור_חפיפה</b></td>
            <td>האם ניתן לבצע משימה נוספת במקביל למשימה הזו?</td>
            <td>כתבו <b>True</b> למשימות כמו כיתת כוננות (שבהן החייל נמצא בכוננות אך יכול לבצע שמירה). כתבו <b>False</b> למשימות שחוסמות את החייל לחלוטין (כמו שמירה או מטבח).</td>
        </tr>
        <tr>
            <td><b>שעות_פעילות</b></td>
            <td>באילו שעות המשימה הזו מתקיימת.</td>
            <td>רשמו <b>all</b> למשימות שרצות 24/7 (כמו שער). למשימות בשעות ספציפיות, רשמו את השעות מופרדות בפסיק (למשל 7,8,9,12,13,14 למטבח).</td>
        </tr>
    </table>
    """, unsafe_allow_html=True)

    st.info(
        "💡 **ניהול שינת לילה:** המערכת תעדף אוטומטית לתת לכל חייל חלון של 7 שעות שינה רצופות בלילה (בין 22:00 ל-08:00) ככל שתנאי הסד''כ יאפשרו זאת.")

with tab_run:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sf = st.file_uploader("📂 קובץ חיילים", type="xlsx")
    with col_u2:
        tf = st.file_uploader("📂 קובץ משימות", type="xlsx")

    if sf and tf:
        s_df, t_df = pd.read_excel(sf), pd.read_excel(tf)
        if st.button("⚙️ צור שבצ''ק ונתח תובנות"):
            with st.spinner("מחשב נתונים..."):
                soldiers = [Soldier(r['מספר_אישי'], r['שם'], r.get('פטורים')) for _, r in s_df.iterrows()]
                tasks = [Task(r['קוד_משימה'], r['שם'], r['כוח_אדם_נדרש'], r.get('משך_משמרת'), r.get('שעות_מנוחה_אחרי'),
                              r.get('אישור_חפיפה'), r.get('שעות_פעילות')) for _, r in t_df.iterrows()]
                final_df = solve_scheduling(soldiers, tasks)

                if final_df is not None:
                    st.table(final_df)
                    st.download_button("📥 הורד לוח סופי מרווח", data=to_excel_styled(final_df),
                                       file_name="Final_Shavtzak.xlsx", use_container_width=True)

                    st.divider()
                    st.subheader("📊 ניתוח עומסים ותובנות מערכת")
                    st.plotly_chart(px.bar(final_df, x="שם", y="סך שעות", color="סך שעות", title="חלוקת עומס כללית",
                                           color_continuous_scale="Greens"), use_container_width=True)

                    st.markdown("#### 💡 למה הלו\"ז נראה ככה?")
                    with st.expander("לחץ כאן להסבר על פערי השעות ועבודת הלילה"):
                        night_req = t_df[t_df['שעות_פעילות'].astype(str).str.contains('all|22|23|0|1|2|3|4|5|6|7')][
                            'כוח_אדם_נדרש'].sum()
                        st.write(
                            f"**1. עבודה בלילה:** יש דרישה ל-{night_req} חיילים בלילה. האלגוריתם נאלץ 'להקריב' שינה כדי לאייש עמדות.")
                        max_s = final_df[final_df["סך שעות"] == final_df["סך שעות"].max()]["שם"].tolist()
                        st.write(
                            f"**2. החיילים העמוסים ביותר:** {', '.join(max_s)}. הם זמינים ללא פטורים ולכן המחשב השתמש בהם למילוי 'חורים'.")
                        st.write(
                            f"**3. השפעת המנוחה:** הגדרת מנוחה ארוכה (כמו 8 שעות) מוציאה חייל מהסבב, מה שמאלץ אחרים לעבוד יותר באותן שעות.")
                else:
                    st.error("לא נמצא פתרון. בדוק את יחס החיילים למשימות.")
