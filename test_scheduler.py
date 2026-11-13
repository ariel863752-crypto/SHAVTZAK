"""
test_scheduler.py — בדיקות אוטומטיות למנוע השיבוץ.

הרצה:   python test_scheduler.py        (או:  pytest test_scheduler.py)

כל בדיקה מתארת תרחיש קטן שבו ברור מראש מה התוצאה הנכונה.
חלק מהבדיקות הן בדיוק התרחישים שבהם הגרסה הקודמת נכשלה.
"""
import scheduler as sc
from scheduler import Config, Directive


def S(sid, name=None, exempt='', roles='', blocked=''):
    s, _ = sc.soldier_from_row({'personal_id': sid, 'full_name': name or f'חייל {sid}',
                                'exemptions': exempt, 'roles': roles, 'blocked_hours': blocked})
    return s


def T(tid, name, req, shift, rest, shared, hours='all', roles='', intensity=5, block='', priority=5):
    t, _ = sc.task_from_row({'task_no': tid, 'name': name, 'required': req, 'shift_hours': shift,
                             'rest_hours': rest, 'shared': shared, 'active_hours': hours,
                             'required_roles': roles, 'intensity': intensity,
                             'blocked_roles': block, 'priority': priority})
    return t


def squad(n, start=1):
    return [S(1000 + i) for i in range(start, start + n)]


def run(soldiers, tasks, directives=None, time_limit=10, **kw):
    cfg = Config(time_limit=time_limit, **kw)
    res = sc.solve_schedule(soldiers, tasks, cfg, directives)
    assert res['metrics']['hard_violations'] == [], res['metrics']['hard_violations']
    return res


def who(res, task_name):
    """{(התחלה, סיום): [שמות]} למשימה"""
    out = {}
    for sid, k, _ in res['assign']:
        sh = res['shifts'][k]
        if sh.task.name == task_name:
            key = (sc.fmt_clock(sh.start, res['cfg'].day_start), sc.fmt_clock(sh.end, res['cfg'].day_start))
            out.setdefault(key, []).append(sid)
    return out


# ── פענוח זמנים ───────────────────────────────────────────────────
def test_parse_end_is_exclusive():
    iv, bad = sc.parse_ranges('8-20', True)
    assert iv == [(480, 1200)] and not bad          # 12 שעות, לא 13
    iv, _ = sc.parse_ranges('08:30-16:00', False)
    assert iv == [(510, 960)]                        # בדיוק 08:30 עד 16:00
    iv, _ = sc.parse_ranges('22-6', True)
    assert iv == [(1320, 1800)]                      # חוצה חצות, 8 שעות
    assert sc.parse_ranges('', True)[0] == [(0, 1440)]
    assert sc.parse_ranges('', False)[0] == []
    assert sc.parse_ranges('all', False)[0] == [(0, 1440)]
    assert sc.to_horizon(sc.parse_ranges('6,7,8', True)[0], 0) == [(360, 540)]
    assert sc.parse_ranges('8-abc', True)[1] == ['8-abc']


def test_horizon_keeps_night_in_one_piece():
    night, _ = sc.parse_ranges('22-6', True)
    assert sc.to_horizon(night, 6 * 60) == [(960, 1440)]          # יממה מ-06:00: רצף אחד
    assert sc.to_horizon(night, 0) == [(0, 360), (1320, 1440)]    # יממה מחצות: נחתך


# ── הבאגים מהגרסה הקודמת ─────────────────────────────────────────
def test_night_task_crossing_midnight():
    res = run(squad(6), [T(1, 'שמירת לילה', 1, 4, 8, False, '22-6')])
    w = who(res, 'שמירת לילה')
    assert sorted(w) == [('02:00', '06:00'), ('22:00', '02:00')]
    assert res['metrics']['short_minutes'] == 0
    assert w[('22:00', '02:00')] != w[('02:00', '06:00')]         # לא אותו חייל ברצף


def test_no_phantom_shortage_on_8_to_20():
    res = run(squad(6), [T(1, 'מטבח', 1, 4, 0, False, '8-20')])
    assert sorted(who(res, 'מטבח')) == [('08:00', '12:00'), ('12:00', '16:00'), ('16:00', '20:00')]
    assert res['metrics']['short_minutes'] == 0
    assert not res['warnings']


def test_block_until_half_past():
    # חסום 08:30–16:00: יכול לשמור 06:30–08:30 ו-16:00–18:00, לא יכול 08:30–10:30
    s = [S(1, blocked='08:30-16:00')]
    res = run(s, [T(1, 'א', 1, 2, 0, False, '06:30-08:30'),
                  T(2, 'ב', 1, 2, 0, False, '16:00-18:00'),
                  T(3, 'ג', 1, 2, 0, False, '08:30-10:30')])
    assert who(res, 'א') and who(res, 'ב') and not who(res, 'ג')
    assert res['metrics']['short_minutes'] == 120


def test_same_soldier_never_twice_in_one_shift():
    # כוננות דורשת 2, רק חייל אחד כשיר → תקן אחד מאויש ואחד מדווח כחוסר
    soldiers = [S(1), S(2, exempt='1')]
    res = run(soldiers, [T(1, 'כוננות', 2, 24, 0, True), T(2, 'שמירה', 1, 4, 0, False)])
    w = who(res, 'כוננות')
    assert sum(len(v) for v in w.values()) == 1
    assert res['metrics']['short_slots'] == 1
    assert 'כוננות' == res['shortage_reasons'][0]['משימה']


def test_standby_allowed_in_parallel_to_exclusive():
    res = run([S(1)], [T(1, 'שמירה', 1, 4, 0, False, '8-12'), T(2, 'כוננות', 1, 24, 0, True)])
    assert res['metrics']['short_minutes'] == 0          # חייל אחד מחזיק את שתיהן


def test_two_exclusive_tasks_never_overlap():
    res = run([S(1)], [T(1, 'שמירה', 1, 4, 0, False, '8-12'), T(2, 'סיור', 1, 4, 0, False, '10-14')])
    assert res['metrics']['short_slots'] == 1


# ── סדר העדיפויות ────────────────────────────────────────────────
def test_coverage_beats_rest():
    # 2 חיילים, שמירה 24/7 במשמרות של 4 עם מנוחה 8: אי אפשר גם לכסות וגם לנוח
    res = run(squad(2), [T(1, 'שמירה', 1, 4, 8, False)])
    assert res['metrics']['short_minutes'] == 0
    assert len(res['metrics']['rest_violations']) > 0


def test_rest_kept_when_possible():
    # 3 חיילים: אפשר לכסות עם מנוחה מלאה (4 עבודה, 8 מנוחה)
    res = run(squad(3), [T(1, 'שמירה', 1, 4, 8, False)])
    assert res['metrics']['short_minutes'] == 0
    assert res['metrics']['rest_violations'] == []


def test_priority_decides_what_stays_empty():
    # חייל אחד, שתי משימות באותן שעות: המשימה הקריטית מאוישת
    res = run([S(1)], [T(1, 'מטבח', 1, 4, 0, False, '8-12', priority=2),
                       T(2, 'ש"ג', 1, 4, 0, False, '8-12', priority=9)])
    assert who(res, 'ש"ג') and not who(res, 'מטבח')


def test_certification_and_blocked_roles():
    soldiers = [S(1, roles='נהג'), S(2, roles='קצין'), S(3)]
    res = run(soldiers, [T(1, 'סיור', 2, 4, 0, False, '8-12', roles='נהג'),
                         T(2, 'מטבח', 1, 4, 0, False, '8-12', block='קצין')])
    patrol = [sid for v in who(res, 'סיור').values() for sid in v]
    assert '1' in patrol and '2' in patrol
    assert who(res, 'מטבח') == {('08:00', '12:00'): ['3']}
    assert res['metrics']['short_minutes'] == 0


def test_continuous_night_rest():
    # 8 חיילים, שמירה אחת 24/7: כולם אמורים לקבל 7 שעות רצופות, ורובם בלילה
    res = run(squad(8), [T(1, 'שמירה', 1, 4, 8, False)])
    m = res['metrics']
    assert m['short_minutes'] == 0 and m['full_rest'] == 8
    assert m['full_night_rest'] >= 6
    assert m['hours_gap'] <= 4


def test_fairness_ignores_who_holds_standby():
    # הכונן לא "נענש": כוננות בעצימות 1 שקולה למשמרת שמירה אחת, לא ל-24 שעות עבודה
    res = run(squad(4), [T(1, 'שמירה', 1, 4, 8, False, intensity=6),
                         T(2, 'כוננות', 1, 24, 0, True, intensity=1)])
    m = res['metrics']
    assert m['short_minutes'] == 0
    assert m['hours_gap'] <= 4 and m['load_gap'] <= 24


# ── הנחיות מפקד ──────────────────────────────────────────────────
def test_commander_directives():
    soldiers = squad(4)
    tasks = [T(1, 'שמירה', 1, 4, 8, False)]
    d = [Directive(sid='1001', tid=1, kind='must_not'),
         Directive(sid='1002', tid=1, kind='must', hours_clock=[(22 * 60, 26 * 60)])]
    res = run(soldiers, tasks, d)
    w = who(res, 'שמירה')
    assert all('1001' not in v for v in w.values())
    assert w[('22:00', '02:00')] == ['1002']


def test_contradicting_must_directives_are_reported():
    tasks = [T(1, 'שמירה', 1, 4, 0, False, '8-12'), T(2, 'סיור', 1, 4, 0, False, '8-12')]
    d = [Directive('1001', 1, 'must'), Directive('1001', 2, 'must')]
    res = run(squad(2), tasks, d)
    assert any('סותרות' in w for w in res['warnings'])
    assert res['metrics']['short_minutes'] == 0


# ── עקביות בין הפותר לפונקציית ההערכה ───────────────────────────
def platoon(n=30):
    soldiers = []
    for i in range(n):
        roles = []
        if i % 6 == 0:
            roles.append('מפקד')
        if i % 5 == 1:
            roles.append('נהג')
        if i % 10 == 3:
            roles.append('חובש')
        if i % 8 == 4:
            roles.append('חמ"ל')
        blocked = '08:00-12:30' if i % 9 == 5 else ('14-18' if i % 11 == 7 else '')
        soldiers.append(S(2000 + i, f'חייל {i + 1}', roles=','.join(roles), blocked=blocked,
                          exempt='4' if i % 7 == 6 else ''))
    tasks = [
        T(1, 'ש"ג', 2, 4, 8, False, 'all', intensity=6, priority=9),
        T(2, 'סיור', 4, 6, 8, False, 'all', roles='מפקד,נהג', intensity=8, priority=8),
        T(3, 'תצפית', 1, 3, 6, False, 'all', intensity=5, priority=7),
        T(4, 'מטבח', 2, 7, 4, False, '06:00-20:00', intensity=4, priority=3, block='מפקד'),
        T(5, 'חמ"ל', 1, 8, 8, False, 'all', roles='חמ"ל', intensity=3, priority=8),
        T(6, 'כוננות', 5, 24, 0, True, 'all', roles='מפקד,נהג,חובש', intensity=1, priority=6),
    ]
    return soldiers, tasks


def test_objectives_match_evaluator():
    soldiers, tasks = platoon(30)
    res = run(soldiers, tasks, time_limit=20)
    m, stages = res['metrics'], {s['שלב'][0]: s for s in res['stages']}
    assert all(s['_status'] in ('OPTIMAL', 'FEASIBLE') for s in res['stages'])
    assert stages['1']['ערך'] == m['short_weighted']            # שלב 1 = מדד הכיסוי
    assert m['wellbeing'] <= stages['2']['ערך']                 # שלב 2 נשמר גם אחרי שלב 3
    assert m['load_std'] < res['baseline_metrics']['load_std']  # ושלב 3 משפר את ההוגנות
    b = res['baseline_metrics']
    assert m['short_weighted'] <= b['short_weighted']
    assert (m['short_weighted'], m['wellbeing']) <= (b['short_weighted'], b['wellbeing'])


def test_shortage_is_minimal_and_explained():
    # 10 חיילים למחלקה של 30: חוסר בלתי נמנע, והמערכת מסבירה כל חוסר
    soldiers, tasks = platoon(10)
    res = run(soldiers, tasks, time_limit=15)
    m = res['metrics']
    assert m['short_minutes'] > 0
    assert len(res['shortage_reasons']) == len(m['shortages'])
    assert m['short_weighted'] <= res['baseline_metrics']['short_weighted']


def test_evaluator_catches_a_bad_manual_schedule():
    # שבצ"ק "ידני" שגוי: אותו חייל בשתי משימות בלעדיות באותו זמן + בלי הסמכה
    soldiers = [S(1), S(2)]
    tasks = [T(1, 'שמירה', 1, 4, 0, False, '8-12'), T(2, 'סיור', 1, 4, 0, False, '8-12', roles='נהג')]
    cfg = Config()
    shifts, _ = sc.prepare(soldiers, tasks, cfg)
    bad = [('1', 0, None), ('1', 1, 'נהג')]
    hard = sc.evaluate(soldiers, shifts, bad, cfg)['hard_violations']
    assert any(h.startswith('H4') for h in hard) and any(h.startswith('H2') for h in hard)


if __name__ == '__main__':
    import sys
    import time
    import traceback
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith('test_') and callable(f)]
    failed = 0
    for name, fn in tests:
        t0 = time.time()
        try:
            fn()
            print(f'✅ {name}  ({time.time() - t0:.1f}s)')
        except Exception:
            failed += 1
            print(f'❌ {name}')
            traceback.print_exc()
    print(f'\n{len(tests) - failed}/{len(tests)} בדיקות עברו')
    sys.exit(1 if failed else 0)
