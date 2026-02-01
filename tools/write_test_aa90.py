# Test script: write to Kết quả!AA90:AK90 and read back
import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from gsheet_client import GSheetClient

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
    rownum = 91
    test_vals = ['TEST_TenA_90', 'TEST_TenB_90', '1', '2', 'LCO', 'HR1A', 'HR2A', 'HR1B', 'HR2B', '3.33', '4.44']
    try:
        gs = GSheetClient(SPREADSHEET_ID, CREDS)
    except Exception as ex:
        print('ERROR creating GSheetClient:', ex)
        return
    aa = idx_to_col(26)
    ak = idx_to_col(26 + len(test_vals) - 1)
    batch = [{'range': f"Kết quả!{aa}{rownum}:{ak}{rownum}", 'values': [test_vals]}]
    print('Writing test batch to', f'Kết quả!{aa}{rownum}:{ak}{rownum}')
    res = gs.batch_update(batch)
    print('batch_update response:', res)
    try:
        read = gs.batch_get(f'Kết quả!{aa}{rownum}:{ak}{rownum}')
        print('Read back:', read)
    except Exception as ex:
        print('ERROR reading back:', ex)

if __name__ == '__main__':
    main()
