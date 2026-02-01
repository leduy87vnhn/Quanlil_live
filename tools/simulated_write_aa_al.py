import os, sys, pickle, csv, datetime
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)

pkl = os.path.join(root, 'src', 'ui_state.pkl')
out_csv = os.path.join(root, 'tools', 'simulated_sheet_AA_AL.csv')
logf = os.path.join(root, 'src', 'vmix_debug.log')

if not os.path.exists(pkl):
    print('ui_state.pkl not found:', pkl)
    sys.exit(1)

with open(pkl, 'rb') as f:
    s = pickle.load(f)

table = s.get('table', []) or []
rows = []
# columns AA..AL (12 columns)
cols = ['AA_TenA','AB_TenB','AC_DiemA','AD_DiemB','AE_Lco','AF_AVGA','AG_AVGB','AH_HR1A','AI_HR2A','AJ_HR1B','AK_HR2B','AL_note']

for r in table:
    # r is list of first 6 columns saved earlier
    ten_a = r[0] if len(r) > 0 else ''
    ten_b = r[1] if len(r) > 1 else ''
    diem_a = r[4] if len(r) > 4 else ''
    # best-effort computed fields left blank (no live vMix fetch in simulation)
    diem_b = ''
    lco = ''
    avga = ''
    avgb = ''
    hr1a = ''
    hr2a = ''
    hr1b = ''
    hr2b = ''
    note = f'simulated {datetime.datetime.now().isoformat()}'
    rows.append([ten_a, ten_b, diem_a, diem_b, lco, avga, avgb, hr1a, hr2a, hr1b, hr2b, note])

with open(out_csv, 'w', newline='', encoding='utf-8') as cf:
    writer = csv.writer(cf)
    writer.writerow(cols)
    for r in rows:
        writer.writerow(r)

print('Wrote simulated AA..AL CSV:', out_csv)
print('Rows written:', len(rows))
for i, r in enumerate(rows[:8]):
    print(i, r)

try:
    with open(logf, 'a', encoding='utf-8') as lf:
        lf.write(f"[{datetime.datetime.now().isoformat()}] SIMULATED_SHEET_WRITE -> {out_csv} rows={len(rows)}\n")
except Exception:
    pass
