import sys, os
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient
sid='1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
creds='credentials.json'

gs=GSheetClient(sid, creds)
row=86
vals=gs.read_table(f'Kết quả!AA{row}:AK{row}')
print('AA86:AK86 =', vals)
headers = gs.read_table('Kết quả!A1:Z1')
print('Headers:', headers)
