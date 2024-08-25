import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials


class RevenueProcessor:
    def __init__(self, file_path, credentials_path, spreadsheet_id, worksheet_name):
        self.file_path = file_path
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.worksheet_name = worksheet_name
        self.df = None
        self.date = None
        self.store_id = None
        self.revenue = None
        self.quantity = None
        self.checks = None
        self.average_check_sum = None
        self.upt = None
        self.entries = None
        self.conversion = None
        self.result_df = None
        self.client = None
        self.worksheet = None

    def load_initial_data(self):
        # Загрузка файла с учетом заголовков, пропуская 6 строк, чтобы получить заголовок столбца J
        self.df = pd.read_excel(self.file_path, skiprows=6)
        # Извлечение даты из названия столбца J
        self.date = self.df.columns[9]  # Столбец J имеет индекс 9

    def load_full_data(self):
        # Загрузка файла, пропуская первые 14 строк для извлечения всех нужных данных
        df_full = pd.read_excel(self.file_path, skiprows=14)
        # Извлечение ID магазина
        full_store_id = df_full['стоимость'].dropna().iloc[-1]
        self.store_id = full_store_id[:4]  # Извлекаем первые 4 символа для ID
        # Извлечение нужных данных
        self.revenue = df_full.iloc[:, 13].dropna().iloc[-1]
        self.quantity = df_full.iloc[:, 14].dropna().iloc[-1]
        self.checks = df_full.iloc[:, 15].dropna().iloc[-1]
        self.average_check_sum = df_full.iloc[:, 17].dropna().iloc[-1]
        self.upt = df_full.iloc[:, 19].dropna().iloc[-1]
        self.entries = df_full.iloc[:, 25].dropna().iloc[-1]
        self.conversion = df_full.iloc[:, 28].dropna().iloc[-2]

    def create_result_dataframe(self):
        # Создание DataFrame с одной записью
        self.result_df = pd.DataFrame({
            'ID': [self.store_id],
            'Дата': [self.date],
            'Выручка': [self.revenue],
            'Кол-во единиц товара': [self.quantity],
            'Кол-во чеков': [self.checks],
            'Сумма среднего чека': [self.average_check_sum],
            'UPT': [self.upt],
            'Входы': [self.entries],
            'Конверсия': [self.conversion]
        })

    def authenticate_google_sheets(self):
        # Аутентификация и доступ к Google Sheets API
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_path, scope)
        self.client = gspread.authorize(creds)

    def load_worksheet(self):
        # Открытие Google Таблицы и нужного листа
        spreadsheet = self.client.open_by_key(self.spreadsheet_id)
        self.worksheet = spreadsheet.worksheet(self.worksheet_name)

    def upload_to_google_sheets(self):
        # Определение последней заполненной строки
        last_row = len(self.worksheet.get_all_values()) + 1
        # Загрузка DataFrame в Google Таблицу без заголовков, начиная с последней заполненной строки
        set_with_dataframe(self.worksheet, self.result_df, row=last_row, include_column_header=False)

    def process(self):
        self.load_initial_data()
        self.load_full_data()
        self.create_result_dataframe()
        self.authenticate_google_sheets()
        self.load_worksheet()
        self.upload_to_google_sheets()
        print(self.result_df)  # Опционально для вывода результата


class SalesReportProcessor:
    def __init__(self, file_path, credentials_path, spreadsheet_id, worksheet_name):
        self.file_path = file_path
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.worksheet_name = worksheet_name
        self.df_full = None
        self.date = None
        self.sellers = None
        self.gross_sales = None
        self.quantity = None
        self.returns_cost = None
        self.returns_quantity = None
        self.discount = None
        self.revenue = None
        self.item_quantity = None
        self.checks = None
        self.average_check = None
        self.upt = None
        self.new_df = None

    def load_data(self):
        # Загрузка основного DataFrame, пропуская первые 14 строк
        self.df_full = pd.read_excel(self.file_path, skiprows=14)
        # Извлечение данных о продавцах из столбца "Продавец"
        self.sellers = self.df_full['Продавец'].dropna().iloc[:-2]

        # Извлечение даты из заголовка столбца J (нужная дата в ячейке J7)
        df_header = pd.read_excel(self.file_path, skiprows=5)
        self.date = df_header.iloc[0, 9]  # Извлечение даты из ячейки J7

    def process_data(self):
        # Извлечение валовой стоимости для соответствующих продавцов (столбец F)
        self.gross_sales = self.df_full.loc[self.sellers.index, 'стоимость']

        # Извлечение количества товара для соответствующих продавцов (столбец G)
        self.quantity = self.df_full.loc[self.sellers.index, 'ое кол-']

        # Извлечение стоимости возвратов для соответствующих продавцов (столбец H)
        self.returns_cost = self.df_full.loc[self.sellers.index, 'возврата']

        # Извлечение количества возвратов для соответствующих продавцов (столбец J)
        self.returns_quantity = self.df_full.loc[self.sellers.index, 'возврат']

        # Извлечение данных о скидках для соответствующих продавцов (столбец K)
        self.discount = self.df_full.loc[self.sellers.index, 'Unnamed: 10']

        # Извлечение данных о выручке для соответствующих продавцов (столбец 'стоимость.1')
        self.revenue = self.df_full.loc[self.sellers.index, 'стоимость.1']

        # Извлечение количества товаров (столбец O)
        self.item_quantity = self.df_full.loc[self.sellers.index, 'кол-во']

        # Извлечение количества чеков для соответствующих продавцов (столбец P)
        self.checks = self.df_full.loc[self.sellers.index, 'чеков']

        # Расчет среднего чека (выручка делится на количество чеков)
        self.average_check = self.revenue.astype(float) / self.checks.astype(float)

        # Извлечение данных о UPT для соответствующих продавцов (столбец 'во')
        self.upt = self.df_full.loc[self.sellers.index, 'во']

    def create_dataframe(self):
        # Проверка соответствия длин массивов и создание DataFrame
        if len(self.sellers) == len(self.gross_sales) == len(self.quantity) == len(self.returns_cost) == len(
                self.returns_quantity) == len(self.discount) == len(self.revenue) == len(self.item_quantity) == len(
            self.checks) == len(self.average_check) == len(self.upt):
            self.new_df = pd.DataFrame({
                'Продавец': self.sellers.values,
                'Дата': [self.date] * len(self.sellers),
                'Валовая стоимость': self.gross_sales.values,
                'Количество товара': self.quantity.values,
                'Стоимость возвратов': self.returns_cost.values,
                'Количество возвратов': self.returns_quantity.values,
                'Скидка': self.discount.values,
                'Выручка': self.revenue.values,
                'Кол-во товаров': self.item_quantity.values,
                'Количество чеков': self.checks.values,
                'Средний чек': self.average_check.values,
                'UPT': self.upt.values
            })
        else:
            raise ValueError("Lengths of data arrays do not match.")

    def upload_to_google_sheets(self):
        # Аутентификация с помощью сервисного аккаунта
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_path, scope)
        client = gspread.authorize(creds)

        # Открытие таблицы и нужного листа
        spreadsheet = client.open_by_key(self.spreadsheet_id)
        worksheet = spreadsheet.worksheet(self.worksheet_name)

        # Определение последней заполненной строки
        last_row = len(worksheet.get_all_values()) + 1

        # Загрузка DataFrame в Google Таблицу без заголовков, начиная с последней заполненной строки
        set_with_dataframe(worksheet, self.new_df, row=last_row, include_column_header=False)

    def process(self):
        self.load_data()
        self.process_data()
        self.create_dataframe()
        self.upload_to_google_sheets()
