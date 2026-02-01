import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from gsheet_client import GSheetClient
SPREADSHEET_ID = '1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
CREDS = 'credentials.json'
try:
    gs = GSheetClient(SPREADSHEET_ID, CREDS)
    r = gs.batch_get('Kết quả!AA91:AK91')
    print('AA91:AK91 =', r)
except Exception as ex:
    print('ERROR', ex)
