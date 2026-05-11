from typing import List, Dict

class Soldier:
    def __init__(self, name: str, roles: List[str]):
        self.name = name
        self.roles = roles  # רשימת הסמכות: ["driver", "commander", "medic"]
        self.assignments_count = {}  # מעקב שוויוניות: { "מטבח": 0, "רס"ר": 0 }
        self.is_available = True
        self.current_tasks_today = 0

    def __repr__(self):
        return f"{self.name} ({', '.join(self.roles)})"

class Task:
    def __init__(self, name: str, total_needed: int, 
                 required_roles: Dict[str, int] = None, 
                 is_fairness_task: bool = False, 
                 max_per_soldier: int = 2):
        """
        name: שם המשימה (למשל 'מטבח')
        total_needed: סך הכל חיילים למשימה
        required_roles: דיקט של תפקידים ספציפיים, למשל {"driver": 1, "commander": 1}
        is_fairness_task: האם זו משימה שצריכה להיות שוויונית
        max_per_soldier: רלוונטי למשימות שוויוניות - כמה פעמים מותר לחייל לבצע
        """
        self.name = name
        self.total_needed = total_needed
        self.required_roles = required_roles if required_roles else {}
        self.is_fairness_task = is_fairness_task
        self.max_per_soldier = max_per_soldier
        self.assigned_soldiers = []

    def get_assigned_roles_count(self) -> Dict[str, int]:
        # סופר כמה בעלי תפקידים כבר שיבצנו למשימה הזו
        counts = {}
        for s in self.assigned_soldiers:
            for role in self.required_roles:
                if role in s.roles:
                    counts[role] = counts.get(role, 0) + 1
        return counts

# --- הגדרת נתונים לדוגמה (מבוסס על המדריך המלא) ---

 soldiers_list = [
    Soldier("יוסי", ["driver"]),
    Soldier("דנה", ["commander"]),
    Soldier("רוני", ["driver", "medic"]),
    Soldier("אריאל", []), # חייל רגיל
    Soldier("נועה", []),
    Soldier("עידו", ["commander"]),
    Soldier("מיכל", []),
    Soldier("עומר", [])
]

tasks_list = [
    # משימה שחייבת נהג אחד ומפקד אחד, השאר (3) יכולים להיות כל אחד
    Task("סיור חודר", total_needed=5, required_roles={"driver": 1, "commander": 1}),
    
    # משימת שוויונית - מטבח (עד 2 פעמים לחייל)
    Task("מטבח", total_needed=3, is_fairness_task=True, max_per_soldier=2)
] 
def assign_tasks(soldiers: List[Soldier], tasks: List[Task]):
    # עוברים משימה-משימה
    for task in tasks:
        print(f"\n--- משבץ למשימה: {task.name} ---")
        
        # 1. שלב א': שיבוץ "עוגנים" (בעלי תפקידים הכרחיים)
        # למשל: אם המשימה צריכה נהג אחד, נמצא נהג אחד ונשבץ אותו
        for role, count_needed in task.required_roles.items():
            assigned_count = 0
            # מחפשים חיילים שיש להם את התפקיד והם פנויים
            eligible_specialists = [s for s in soldiers if s.is_available and role in s.roles]
            
            for specialist in eligible_specialists:
                if assigned_count < count_needed:
                    task.assigned_soldiers.append(specialist)
                    specialist.is_available = False # חייל ששובץ לא זמין למשימות אחרות באותו זמן
                    assigned_count += 1
                    print(f"שיבוץ עוגן: {specialist.name} שובץ כ-{role}")
        
        # 2. שלב ב': השלמת מכסה (חיילים רגילים)
        # כאן נכנס עניין השוויוניות (למשל למטבח)
        needed_to_complete = task.total_needed - len(task.assigned_soldiers)
        
        if needed_to_complete > 0:
            # סינון חיילים פנויים (שלא שובצו כעוגנים)
            potential_soldiers = [s for s in soldiers if s.is_available]
            
            # אם זו משימת שוויוניות - נמיין אותם לפי מי שעשה הכי פחות מהמשימה הזו
            if task.is_fairness_task:
                potential_soldiers.sort(key=lambda s: s.assignments_count.get(task.name, 0))
                
                # בדיקת חסם עליון (למשל: לא יותר מ-2 מטבחים)
                potential_soldiers = [
                    s for s in potential_soldiers 
                    if s.assignments_count.get(task.name, 0) < task.max_per_soldier
                ]

            # שיבוץ עד להשלמת המכסה
            for i in range(min(needed_to_complete, len(potential_soldiers))):
                chosen = potential_soldiers[i]
                task.assigned_soldiers.append(chosen)
                chosen.is_available = False
                
                # עדכון מונה שוויוניות במידה וזו משימה כזו
                if task.is_fairness_task:
                    chosen.assignments_count[task.name] = chosen.assignments_count.get(task.name, 0) + 1
                
                print(f"שיבוץ כללי: {chosen.name} שובץ למשימה")

        # בדיקה - האם הצלחנו למלא את המשימה?
        if len(task.assigned_soldiers) < task.total_needed:
            print(f"אזהרה: לא נמצאו מספיק חיילים למשימת {task.name}!") 
def print_shabtzak_results(tasks: List[Task]):
    print("\n" + "="*30)
    print("תוצאות שיבוץ שבצ\"ק")
    print("="*30)
    
    for task in tasks:
        names = [s.name for s in task.assigned_soldiers]
        print(f"📌 {task.name} ({len(names)}/{task.total_needed}):")
        print(f"   חיילים: {', '.join(names)}")
        
        # בדיקה אם חסרים אנשים
        if len(names) < task.total_needed:
            print(f"   ⚠️ שים לב: חסרים {task.total_needed - len(names)} חיילים!")
    print("="*30)

# --- הרצת המערכת ---

# 1. קריאה לפונקציית השיבוץ (מחלק 2)
assign_tasks(soldiers_list, tasks_list)

# 2. הדפסת התוצאות
print_shabtzak_results(tasks_list)

# 3. הדפסת מצב שוויוניות (למעקב)
print("\nמונה משימות שוויוניות (מטבח):")
for s in soldiers_list:
    count = s.assignments_count.get("מטבח", 0)
    print(f"- {s.name}: {count} פעמים")
