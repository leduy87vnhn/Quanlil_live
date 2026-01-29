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
