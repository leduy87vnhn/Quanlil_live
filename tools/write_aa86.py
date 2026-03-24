import sys, os
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient

sid = '1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
creds = 'credentials.json'

gs = GSheetClient(sid, creds)
row = 86
values = [
    ['TEST_TenA', 'TEST_TenB', '10', '20', 'LCO', 'HR1A', 'HR2A', 'HR1B', 'HR2B', 'AVGA', 'AVGB']
]
range_name = f'Kết quả!AA{row}:AK{row}'
print('Writing to', range_name, 'values=', values)
res = gs.batch_update([{'range': range_name, 'values': values}])
print('batch_update result:', res)
read = gs.batch_get(range_name)
print('Read back:', read)
