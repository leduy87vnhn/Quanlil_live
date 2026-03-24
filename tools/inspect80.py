import sys, os
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient
sid='1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
gs=GSheetClient(sid, 'credentials.json')
rows=gs.read_table('Kết quả!A1:Z2000')
if not rows:
    print('no rows')
else:
    for i,r in enumerate(rows):
        if any(v and '80' in str(v) for v in r.values()):
            rownum=i+2
            print('found row', rownum)
            vals=gs.read_table(f'Kết quả!AA{rownum}:AK{rownum}')
            print('AA..AK =', vals)
            headers=list(rows[0].keys())
            dcol=headers[3] if len(headers)>3 else None
            fcol=headers[5] if len(headers)>5 else None
            print('D =', r.get(dcol))
            print('F =', r.get(fcol))
            break
