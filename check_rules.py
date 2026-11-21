"""Rule check for update_inbound.py - run after EVERY change:  python check_rules.py

Compares the numbers in update_inbound.py with the NUMBERS REGISTER below (the same table as CLAUDE.md 'Numbers register').
  USER    = the planner's decision. A mismatch = FAIL: the change needs the planner's OK first, then update this file + CLAUDE.md.
  DEFAULT = not confirmed by the planner. A mismatch = FAIL too (so no number moves silently); ask the planner to confirm it.
The script is not run and no input file is needed (it reads the code only), so the check takes a second.
"""
import ast, sys, os

SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'update_inbound.py')

# name: (expected value, owner, meaning)
CONSTANTS = {
    'COV_RED_DAYS':       (7,   'USER',    "dashboards red under 7 days (12-Nov-2026 'Red is under 7 days')"),
    'DASH_HTML_MAX_DAYS': (21,  'USER',    'web page lists Inv. Days today <= 21 (21-Nov-2026)'),
    'CHAT_SPLIT_DAYS':    (14,  'USER',    'chat pictures split at Inv. Days today <= 14 / > 14 (21-Nov-2026)'),
    'DASH_HTML_WEEKS':    (3,   'USER',    'web page weeks, at most 3 (21-Nov-2026 "Make it show 3 weeks")'),
    'WSA_HTML_WEEKS':     (3,   'USER',    'WSA web page weeks, 3 (21-Nov-2026)'),
    'CONSUMABLE_CODES':   ({'RMTAP00000018', 'RMTAP00000025', 'RMTTT00000018', 'RMTTT00000009',
                            'RMLAB00000097', 'RMLAB00000098', 'RMLAB00000121'}, 'USER', "the planner's consumables list"),
    'LEAD_TIME_RM':       (49,  'USER',    'fallback RM lead time when Settings has no row (21-Nov-2026: 49)'),
    'LEAD_TIME_CSPM':     (14,  'USER', 'fallback packaging / consumables lead time'),
    'ALERT_HORIZON':      (112, 'USER', "look-ahead of 'below lead time later' (16 weeks)"),
    'RISK_DAYS':          (5,   'USER', 'dashboard risk order 2nd step'),
    'RELEASE_WINDOW':     (7,   'USER', "STOCK-OUT LATER / 'ORDER NOW' window"),
    'RESCHED_OUT_X':      (3,   'USER', "'RESCHEDULE OUT' when stock > 3 x target"),
    'ETD_SOON':           (7,   'USER', "ETD alert 'ships within 7 days'"),
    'ATA_SHOW':           (4,   'USER', 'arrival dates shown before one total'),
    'MOVE_DAYS':          (14,  'USER', 'Movement Log days'),
    'SHORT_LEAD':         (20,  'USER', 'Inbound short lead flag (ETD -> plant days)'),
    'LONG_LEAD':          (75,  'USER', 'Inbound long lead flag'),
    'SHORT_HORIZON':      (60,  'USER', 'shortage measured over 60 days'),
    'DASH_STEP2':         (14,  'USER', 'MRP_Daily dark colours step 2'),
    'DASH_BLUE1':         (14,  'USER', 'MRP_Daily dark colours blue 1'),
    'DASH_BLUE2':         (39,  'USER', 'MRP_Daily dark colours blue 2'),
    'DASH_MIN_LT':        (14,  'USER', 'minimum lead time for colour steps'),
    'IT_STATUS_ONLY':     (False, 'USER', 'in transit = every PPC status, not only IN-TRANSIT'),
    'UNLIMITED_CODES':    ({'PMPKM00000338', 'PMPKM00000523'}, 'USER', 'never short (Settings can add more)'),
    'CUST_TARGET_DEFAULT': ([('Pepper Solar Farm LLC', 'L1', 3000, 'No'), ('Lucky 7 Solar Farm LLC', 'L2', 3000, 'No'),
                             ('Clenera Crimson', 'L3', 3600, 'No'), ('SB Energy G12R', 'L4', 3600, 'No'),
                             ('Clenera Snowflake', 'OEM', 4000, 'Yes')], 'USER',
                            "customer daily targets when Settings has no 'CUSTOMER DAILY TARGET' table (confirmed 21-Nov-2026)"),
}

# numbers written inside functions: (exact code text that must exist, owner, meaning)
SNIPPETS = [
    ("PLAN_AVG_DAYS=globals().get('PLAN_AVG_DAYS',49)", 'USER', 'Action List daily plan = average of next 49 production days (NOT a Settings row) (confirmed 21-Nov-2026)'),
    ("(e_+pd.Timedelta(days=5))",                       'USER', 'PPC line without ATA: Revised ETA Port + 5 days'),
    ("PPC_IT_NEED=('ETD','ETA','BL#','AtA')",           'USER', 'PPC In_Transit line counts only with ETD + ETA + BL# (HBL / MBL) + Final AtA (21-Nov-2026)'),
    ("below<=(lead or LEAD_TIME_RM)+14",               'USER', 'status WATCH: below-target day within lead time + 14'),
    ("LEAD_TIME_RM)*2/3]",                              'USER', 'line CRITICAL: runs under 2/3 of its longest lead time'),
    ("sc[best]-sc[fc]<2.0",                             'USER', 'shipment tracing: Final RM Code replaced only when another code is 2+ points better'),
    ("abs(dd_)>=3)",                                    'USER', 'Inbound AQ: PPC date 3+ days off the WSA date'),
    ("if empty>(500 if raw else 2000): break",          'USER', 'PPC In_Transit reading stops after 500 empty rows'),
    ("v<7?\"s\":v<21?\"l\":\"b\"",                     'USER',    "web colours red <7 (incl. <1) / 7-21 / 21+ (21-Nov-2026 '1-7 day is also red')"),
    ("Math.max(34,Math.floor(innerWidth/3/",            'USER',    'web week columns = 1/3 of the screen (material column 2/3), 34 px minimum (21-Nov-2026)'),
]


def read_constants(path):
    tree = ast.parse(open(path, encoding='utf-8').read())
    out = {}
    for node in tree.body:                                   # top level only: the CONFIG constants
        for t in (node.targets if isinstance(node, ast.Assign) else []):
            if isinstance(t, ast.Name) and t.id in CONSTANTS and t.id not in out:
                try: out[t.id] = ast.literal_eval(node.value)
                except ValueError: out[t.id] = '<not a plain value>'
    return out


def main():
    src = open(SCRIPT, encoding='utf-8').read()
    got = read_constants(SCRIPT)
    fails = 0
    print(f"{'RESULT':6} {'OWNER':7} {'NAME / CODE':34} {'EXPECTED':>10} {'IN CODE':>10}  MEANING")
    for name, (exp, owner, why) in CONSTANTS.items():
        val = got.get(name, '<missing>')
        ok = val == exp
        fails += not ok
        short = lambda v: (str(v) if len(str(v)) <= 10 else str(v)[:7] + '...')
        print(f"{'OK' if ok else 'FAIL':6} {owner:7} {name:34} {short(exp):>10} {short(val):>10}  {why}")
    for text, owner, why in SNIPPETS:
        ok = text in src
        fails += not ok
        print(f"{'OK' if ok else 'FAIL':6} {owner:7} {text[:34]:34} {'':>10} {'found' if ok else 'NOT FOUND':>10}  {why}")
    print()
    if fails:
        print(f"FAIL: {fails} number(s) differ from the register. USER = needs the planner's OK; DEFAULT = ask the planner to confirm.")
        print("If the change was agreed: update check_rules.py AND the 'Numbers register' in CLAUDE.md in the same commit.")
        sys.exit(1)
    print(f"PASS: all {len(CONSTANTS) + len(SNIPPETS)} registered numbers match.")


if __name__ == '__main__':
    main()
