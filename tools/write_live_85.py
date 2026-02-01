import sys, os, csv
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from gsheet_client import GSheetClient

SID='1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
CREDS='credentials.json'
CSV_PATH=os.path.join(os.path.dirname(os.path.dirname(__file__)), 'src', 'vmix_input1_temp.csv')

# read vmix csv
rows=[]
if os.path.exists(CSV_PATH):
    with open(CSV_PATH, 'r', encoding='utf-8') as f:
        reader=csv.DictReader(f)
        for r in reader:
            rows.append(r)
else:
    print('No vmix csv at', CSV_PATH)

# pick row for match 85: prefer index 84 (0-based)
if len(rows) >= 85:
    vm = rows[84]
    print('Using vmix row index 84 for match 85')
elif rows:
    vm = rows[min(0, len(rows)-1)]
    print('Not enough rows; using first vmix row')
else:
    print('No vmix rows available; abort')
    sys.exit(1)

# prepare values
extra_vals = [
    vm.get('TenA',''), vm.get('TenB',''), vm.get('DiemA',''), vm.get('DiemB',''), vm.get('Lco',''),
    vm.get('HR1A',''), vm.get('HR2A',''), vm.get('HR1B',''), vm.get('HR2B',''), vm.get('AvgA',''), vm.get('AvgB','')
]

gs = GSheetClient(SID, CREDS)
rownum = 86
# write AA86:AK86
range_name = f"Kết quả!AA{rownum}:AK{rownum}"
print('Writing', range_name, extra_vals)
res = gs.batch_update([{'range': range_name, 'values': [extra_vals]}])
print('batch_update response:', res)
# write TenA -> col D, TenB -> col F
# read headers to get actual header names
hdrs = gs.read_table('Kết quả!A1:Z1')
if hdrs:
    headers = list(hdrs[0].keys())
else:
    headers = []
# col letters
def idx_to_letter(i):
    n=i+1
    s=''
    while n>0:
        n, rem = divmod(n-1,26)
        s = chr(65+rem) + s
    return s
# write D and F by column index 3 and 5
d_cell = f"Kết quả!{idx_to_letter(3)}{rownum}"
f_cell = f"Kết quả!{idx_to_letter(5)}{rownum}"
print('Writing D,F cells', d_cell, f_cell)
gs.batch_update([{'range': d_cell, 'values': [[vm.get('TenA','')]]}, {'range': f_cell, 'values': [[vm.get('TenB','')]]}])
# try to find Điểm A / Điểm B column names
possible_diem_a = ['Điểm A','DiemA','ĐiểmA','Điểm A']
possible_diem_b = ['Điểm B','DiemB','ĐiểmB','Điểm B']
found_diem_a=None
found_diem_b=None
for h in headers:
    hn = h.replace(' ','').lower()
    if any(x.replace(' ','').lower()==hn for x in possible_diem_a):
        found_diem_a=h
    if any(x.replace(' ','').lower()==hn for x in possible_diem_b):
        found_diem_b=h
if found_diem_a:
    cell = f"Kết quả!{idx_to_letter(headers.index(found_diem_a))}{rownum}"
    gs.batch_update([{'range': cell, 'values': [[vm.get('DiemA','')]]}])
if found_diem_b:
    cell = f"Kết quả!{idx_to_letter(headers.index(found_diem_b))}{rownum}"
    gs.batch_update([{'range': cell, 'values': [[vm.get('DiemB','')]]}])
print('Done')
