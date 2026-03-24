import os
import sys
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient

SPREADSHEET_ID = '1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
CREDS = 'credentials.json'

def idx_to_col(i):
    n = i+1
    s = ''
    while n>0:
        n, rem = divmod(n-1, 26)
        s = chr(65+rem) + s
    return s

def main():
    rownum = 86
    test_vals = ['TEST_TenA', 'TEST_TenB', '10', '20', 'LCO', 'HR1A', 'HR2A', 'HR1B', 'HR2B', '1.111', '2.222']
    try:
        gs = GSheetClient(SPREADSHEET_ID, CREDS)
    except Exception as ex:
        print('ERROR creating GSheetClient:', ex)
        return
    batch = []
    # D86, F86
    batch.append({'range': f"Kết quả!{idx_to_col(3)}{rownum}", 'values': [[test_vals[0]] ]})
    batch.append({'range': f"Kết quả!{idx_to_col(5)}{rownum}", 'values': [[test_vals[1]] ]})
    # AA86:AK86
    aa = idx_to_col(26)
    ak = idx_to_col(26 + len(test_vals) - 1)
    batch.append({'range': f"Kết quả!{aa}{rownum}:{ak}{rownum}", 'values': [test_vals]})
    print('Writing test batch to', f'Kết quả!{aa}{rownum}:{ak}{rownum}')
    res = gs.batch_update(batch)
    print('batch_update response:', res)
    # read back
    try:
        read = gs.batch_get(f'Kết quả!{aa}{rownum}:{ak}{rownum}')
        print('Read back:', read)
    except Exception as ex:
        print('ERROR reading back:', ex)

if __name__ == '__main__':
    main()
