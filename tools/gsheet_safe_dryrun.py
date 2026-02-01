"""Dry-run for GSheetClient.batch_update_safe filtering.
This script avoids calling Google API and only exercises the local filtering logic.
"""
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient


def main():
    gs = GSheetClient.__new__(GSheetClient)
    # Do not initialize real service. We'll directly use helper methods.
    sample = [
        {'range': 'Kết quả!AA3', 'values': [["v1"]]},
        {'range': 'Kết quả!AB3', 'values': [["v2"]]},
        {'range': 'Kết quả!A3', 'values': [["X"]]},
        {'range': 'Sheet1!AC10', 'values': [["v3"]]},
        {'range': 'Other!AZ5', 'values': [["skip"]]},
    ]

    print('Sample ranges:')
    for it in sample:
        print(' ', it['range'])

    allowed = ['AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL']
    filtered = []
    for item in sample:
        col = gs._extract_col_letters(item['range'])
        print(f"Extracted col from {item['range']}: '{col}'")
        if col in allowed:
            filtered.append(item)

    print('\nFiltered items to be sent:')
    for it in filtered:
        print(' ', it)

    print('\nDone (no network calls performed).')


if __name__ == '__main__':
    main()
