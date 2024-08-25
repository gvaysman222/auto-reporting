import gspread
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheetsClient:
    def __init__(self, json_keyfile_name, spreadsheet_id):
        self.json_keyfile_name = json_keyfile_name
        self.spreadsheet_id = spreadsheet_id
        self.client = self.setup_google_sheets_connection()

    def setup_google_sheets_connection(self):
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.json_keyfile_name, scope)
        client = gspread.authorize(creds)
        return client

    def get_worksheet(self, worksheet_index):
        spreadsheet = self.client.open_by_key(self.spreadsheet_id)
        worksheet = spreadsheet.get_worksheet(worksheet_index)
        return worksheet

    def update_google_sheet(self, worksheet, data):
        df_list = data.values.tolist()
        last_row = len(worksheet.get_all_values()) + 1
        worksheet.update(f'A{last_row}', df_list)

    def get_worksheet_by_name(self, worksheet_name):
        spreadsheet = self.client.open_by_key(self.spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet
