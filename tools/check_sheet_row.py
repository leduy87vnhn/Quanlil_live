import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient
sid='1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
creds='credentials.json'
gs=GSheetClient(sid, creds)
rows=gs.read_table('Kết quả!A1:Z2000')
if not rows:
    print('No rows')
else:
    for i,r in enumerate(rows):
        # try matching '80' in Trận or normalized
        import sys, os
        sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
        from src.gsheet_client import GSheetClient

        sid = '1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
        creds = 'credentials.json'

        gs = GSheetClient(sid, creds)
        rows = gs.read_table('Kết quả!A1:Z2000')
        if not rows:
            print('No rows')
        else:
            for i, r in enumerate(rows):
                # try matching '80' in Trận or normalized
                found = False
                for k, v in r.items():
                    if v and '80' in str(v):
                        found = True
                        break
                if not found:
                    continue
                print('Found at index', i, 'rownum', i + 2)
                aa = 'AA' + str(i + 2)
                ak = 'AK' + str(i + 2)
                vals = gs.read_table(f'Kết quả!{aa}:{ak}')
                print('AA..AK:', vals)
                headers = list(rows[0].keys())
                dcol = headers[3] if len(headers) > 3 else None
                fcol = headers[5] if len(headers) > 5 else None
                print('D col name:', dcol, 'F col name:', fcol)
                if dcol:
                    print('D value:', r.get(dcol))
                if fcol:
                    print('F value:', r.get(fcol))
                break
