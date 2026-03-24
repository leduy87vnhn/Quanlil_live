
import sys
import os
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient

# Thay đổi đường dẫn credentials và spreadsheet_id tại đây
CREDENTIALS_PATH = os.path.abspath('vmix-youtube-auto-de998d3481dd.json')
SPREADSHEET_ID = '1vqYeIELbZEiz6ryIGCZ5PKjehWBjN98pPv9_84Hghj4'
RANGE = 'Kết quả!A1:Z1000'

def test_gsheet():
    print('Credentials:', CREDENTIALS_PATH)
    print('Spreadsheet ID:', SPREADSHEET_ID)
    try:
        gs = GSheetClient(SPREADSHEET_ID, CREDENTIALS_PATH)
        rows = gs.read_table(RANGE)
        print('Read success! Rows:', len(rows))
        if rows:
            print('First row:', rows[0])
        else:
            print('No data found in range.')
    except Exception as e:
        print('ERROR:', e)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_gsheet()
