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

	def write_table(self, range_name: str, rows: List[List[str]], value_input_option: str = "USER_ENTERED") -> dict:
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

	def batch_get(self, range_name: str) -> List[List[str]]:
		"""Return raw values (list of rows) for the specified range.

		This returns the same structure as the API `values.get` -> `values` field.
		"""
		sheet = self.service.spreadsheets()
		result = sheet.values().get(spreadsheetId=self.spreadsheet_id, range=range_name, majorDimension='ROWS').execute()
		return result.get('values', [])
