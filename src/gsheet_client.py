from typing import List, Dict
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

class GSheetClient:
    def __init__(self, spreadsheet_id: str, credentials_path: str):
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        self.service = build("sheets", "v4", credentials=creds)
        self.spreadsheet_id = spreadsheet_id

    def read_table(self, range_name: str) -> List[Dict[str, str]]:
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.spreadsheet_id, range=range_name).execute()
        values = result.get("values", [])
        if not values:
            return []
        headers = values[0]
        rows = []
        for row in values[1:]:
            obj = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
            rows.append(obj)
        return rows

    def get_metadata(self) -> Dict:
        sheet = self.service.spreadsheets()
        return sheet.get(spreadsheetId=self.spreadsheet_id).execute()

    def write_table(self, range_name: str, rows: List[List[str]], value_input_option: str = "USER_ENTERED") -> dict:
        # Write the provided rows to the given A1 range using the Sheets API.
        body = {"values": rows}
        sheet = self.service.spreadsheets()
        return sheet.values().update(spreadsheetId=self.spreadsheet_id, range=range_name,
                                      valueInputOption=value_input_option, body=body).execute()

    def batch_update(self, data: List[Dict], value_input_option: str = "USER_ENTERED") -> dict:
        """Perform a single batchUpdate for multiple ranges.

        Each item in `data` should be a dict with keys: `range` and `values` (list of lists).
        Example: [{"range": "Sheet1!A2:A2", "values": [["x"]]}, ...]
        """
        body = {"valueInputOption": value_input_option, "data": data}
        sheet = self.service.spreadsheets()
        return sheet.values().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

    def _extract_col_letters(self, a1_range: str) -> str:
        """Extract the column letters from an A1 range like 'Sheet1!AA12' or 'AA12'."""
        try:
            if '!' in a1_range:
                _, a1 = a1_range.split('!', 1)
            else:
                a1 = a1_range
            # remove row numbers and any range suffix
            col = ''
            for ch in a1:
                if ch.isalpha():
                    col += ch.upper()
                else:
                    break
            return col
        except Exception:
            return ''

    def batch_update_safe(self, data: List[Dict], allowed_cols: List[str] = None, value_input_option: str = "USER_ENTERED") -> dict:
        """Filter batch update items to only allowed column letters and perform the update.

        - `data` is a list of {"range": "Sheet!AA3", "values": [[...]]}
        - `allowed_cols` is list of column letter strings like ['AA','AB',...]. If None, defaults to ['AA'..'AL'].
        The method will silently skip any entries not in the allowed columns, and still return the API result
        for the permitted subset. If no items remain, returns an empty dict.
        """
        # default allowed AA..AL
        if allowed_cols is None:
            # generate AA..AL
            base = []
            for i in range(ord('A'), ord('Z')+1):
                pass
            # hardcode AA..AL
            allowed_cols = ['AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL']

        filtered = []
        for item in data:
            try:
                rng = item.get('range') or ''
                col = self._extract_col_letters(rng)
                if col and col in allowed_cols:
                    filtered.append(item)
            except Exception:
                continue

        if not filtered:
            return {}

        body = {"valueInputOption": value_input_option, "data": filtered}
        sheet = self.service.spreadsheets()
        return sheet.values().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

    def batch_get(self, range_name: str) -> List[List[str]]:
        """Return raw values (list of rows) for the specified range.

        This returns the same structure as the API `values.get` -> `values` field.
        """
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.spreadsheet_id, range=range_name, majorDimension='ROWS').execute()
        return result.get('values', [])
