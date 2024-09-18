import pandas as pd
import re

class RetailSalesProcessor1CSellers:
    def __init__(self, google_sheets_client, worksheet_index):
        self.google_sheets_client = google_sheets_client
        self.worksheet_index = worksheet_index
        self.worksheet = self.google_sheets_client.get_worksheet(self.worksheet_index)

    @staticmethod
    def extract_store_name(file_name):
        match = re.search(r'(LaCrO|Guess)', file_name, re.IGNORECASE)
        return match.group(0) if match else "Unknown"

    def read_and_filter_excel_data(self, excel_file_path):
        df = pd.read_excel(excel_file_path)

        filtered_names = df['Продавец'].apply(self.is_seller_name)
        filtered_dates = df['Продавец'].apply(self.is_date)

        sellers = df[filtered_names]['Продавец'].values
        dates = df[filtered_dates]['Продавец'].values

        return df, sellers, dates

    def process_data(self, df, sellers, dates, store_name):
        df = df[~df['Продавец'].str.strip().str.lower().isin(["итого"])]

        quantity_values = []
        check_values = []
        gross_values = []
        revenue_values = []
        average_check_values = []
        upt_values = []

        for i, name in enumerate(df[df['Продавец'].isin(sellers)].index):
            seller_rows = df.loc[name + 2:]

            next_seller = df.loc[seller_rows.index, 'Продавец'].apply(self.is_seller_name)
            if next_seller.any():
                next_seller_index = next_seller.idxmax()
                seller_rows = seller_rows.loc[:next_seller_index - 1]

            quantity_value = seller_rows['Количество'].sum()
            check_value = df.loc[name + 1, 'Количество чеков']
            gross_value = df.loc[name + 1, 'Получено картами']
            revenue_value = df.loc[name + 1, 'Выручка, ']
            average_check_value = df.loc[name + 1, 'Статистика продаж']

            upt_value = quantity_value / check_value if check_value else 0

            quantity_values.append(quantity_value)
            check_values.append(check_value)
            gross_values.append(gross_value)
            revenue_values.append(revenue_value)
            average_check_values.append(average_check_value)
            upt_values.append(upt_value)

        df_result = pd.DataFrame({
            'Продавец': sellers,
            'Дата': dates[:len(sellers)],
            'Валовая стоимость': gross_values,
            'Кол-во товаров': [''] * len(sellers),
            'Стоимость возвратов': [''] * len(sellers),
            'Кол-во возвратов': [''] * len(sellers),
            'Скидка': [''] * len(sellers),
            'Выручка': revenue_values,
            'Количество товаров': quantity_values,
            'Количество чеков': check_values,
            'Средний чек': average_check_values,
            'UPT': upt_values,
            'Магазин': [store_name] * len(sellers),
        })

        return df_result.fillna('')

    def process_and_update(self, excel_file_path):
        # Извлечение имени магазина из имени файла
        file_name = excel_file_path.split('/')[-1]
        store_name = self.extract_store_name(file_name)

        df, sellers, dates = self.read_and_filter_excel_data(excel_file_path)
        df_result = self.process_data(df, sellers, dates, store_name)
        self.google_sheets_client.update_google_sheet(self.worksheet, df_result)

    @staticmethod
    def is_seller_name(value):
        if not isinstance(value, str):
            return False
        return re.match(r'^[А-ЯЁ][а-яё]+ [А-ЯЁ]\.?([А-ЯЁ]\.?)?([а-яё]+)?$', value) is not None

    @staticmethod
    def is_date(value):
        if not isinstance(value, str):
            return False
        return re.match(r'^\d{2}\.\d{2}\.\d{4}$', value) is not None
