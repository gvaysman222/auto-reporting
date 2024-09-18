import pandas as pd
import re


class RetailSalesProcessorShop:
    def __init__(self, google_sheets_client, worksheet_index):
        self.google_sheets_client = google_sheets_client
        self.worksheet_index = worksheet_index
        self.worksheet = self.google_sheets_client.get_worksheet(self.worksheet_index)

    @staticmethod
    def extract_store_name(file_name):
        match = re.search(r'(LACRO|Guess)', file_name, re.IGNORECASE)
        if match:
            return match.group(0)
        else:
            return "Unknown"

    @staticmethod
    def extract_date_from_data(df):
        for i, row in df.iterrows():
            if isinstance(row.iloc[0], str) and re.search(r'\d{2}\.\d{2}\.\d{4}', row.iloc[0]):
                return row.iloc[0]
        return "Unknown"

    @staticmethod
    def extract_total_revenue(df):
        itogo_row = df[df.iloc[:, 0].str.contains("Итого", na=False, case=False)]
        if not itogo_row.empty:
            total_revenue = itogo_row.iloc[0]['Выручка, ']
            return total_revenue
        else:
            return 0

    @staticmethod
    def extract_total_checks(df):
        itogo_row = df[df.iloc[:, 0].str.contains("Итого", na=False, case=False)]
        if not itogo_row.empty:
            total_checks = itogo_row.iloc[0]['Количество чеков']
            return total_checks
        else:
            return 0

    @staticmethod
    def extract_total_quantity(df):
        quantity = df[df.iloc[:, 0].str.contains("Итого", na=False, case=False)]
        if not quantity.empty:
            total_quantity = quantity.iloc[0]['Количество']
            return total_quantity
        else:
            return 0

    @staticmethod
    def extract_average_check(df):
        itogo_row = df[df.iloc[:, 0].str.contains("Итого", na=False, case=False)]
        if not itogo_row.empty:
            average_check = itogo_row.iloc[0]['Статистика продаж']
            return average_check
        else:
            return 0

    @staticmethod
    def extract_upt(df):
        itogo_row = df[df.iloc[:, 0].str.contains("Итого", na=False, case=False)]
        if not itogo_row.empty:
            try:
                total_units = float(itogo_row.iloc[0]['Количество'])
                total_checks = int(itogo_row.iloc[0]['Количество чеков'])
                if total_checks != 0:
                    upt = total_units / total_checks
                else:
                    upt = 0
                return upt
            except (ValueError, TypeError) as e:
                return 0
        else:
            return 0

    def process_and_update(self, excel_file_path):
        file_name = excel_file_path.split('/')[-1]
        store_name = self.extract_store_name(file_name)

        df = pd.read_excel(excel_file_path)
        df = df.fillna('')
        report_date = self.extract_date_from_data(df)
        total_revenue = self.extract_total_revenue(df)
        total_checks = self.extract_total_checks(df)
        total_quantity = self.extract_total_quantity(df)
        average_check = self.extract_average_check(df)
        upt_value = self.extract_upt(df)

        df_result = pd.DataFrame({
            'Магазин': [store_name],
            'Дата': [report_date],
            'Итоговая выручка': [total_revenue],
            'Количество товаров': [total_quantity],
            'Количество чеков': [total_checks],
            'Средний чек': [average_check],
            'UPT': [upt_value]
        })

        self.google_sheets_client.update_google_sheet(self.worksheet, df_result)
