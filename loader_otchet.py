import pandas as pd
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Установка соединения с Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('mailapi-431104-8992c2888d0e.json', scope)
client = gspread.authorize(creds)
# Загрузка первой таблицы


# Открытие существующей Google Таблицы по ее ID
spreadsheet_id = '1T_xOCiDFiE8BWDsK_Iodo79fu2aLV_DqgZQrtbpwhU4'  # Замените на ID вашей Google Таблицы
spreadsheet = client.open_by_key(spreadsheet_id)

# Получение листа по индексу (например, первый лист имеет индекс 0)
worksheet_index = 3  # Замените на нужный индекс листа
worksheet = spreadsheet.get_worksheet(worksheet_index)

# Чтение Excel файла, пропуская первые несколько строк
df1 = pd.read_excel('downloads/Розничная выручка по продавцам (XLS).xls')


def is_seller_name(value):
    if not isinstance(value, str):
        return False
    return re.match(r'^[А-ЯЁ][а-яё]+ [А-ЯЁ]\.?([А-ЯЁ]\.?)?([а-яё]+)?$', value) is not None


def is_date(value):
    if not isinstance(value, str):
        return False
    return re.match(r'^\d{2}\.\d{2}\.\d{4}$', value) is not None


filtered_names = df1['Продавец'].apply(is_seller_name)
filtered_dates = df1['Продавец'].apply(is_date)

sellers = df1[filtered_names]['Продавец'].values
dates = df1[filtered_dates]['Продавец'].values

quantity_values = []
check_values = []
gross_values = []
revenue_values = []
average_check_values = []
upt_values = []

for i, name in enumerate(df1[filtered_names].index):
    seller_rows = df1.loc[name + 1:]

    next_seller = df1.loc[seller_rows.index, 'Продавец'].apply(is_seller_name)
    if next_seller.any():
        next_seller_index = next_seller.idxmax()
        seller_rows = seller_rows.loc[:next_seller_index - 1]

    quantity_value = seller_rows['Количество'].sum()
    check_value = df1.loc[name + 1, 'Количество чеков']
    gross_value = df1.loc[name + 1, 'Получено картами']
    revenue_value = df1.loc[name + 1, 'Выручка, ']
    average_check_value = df1.loc[name + 1, 'Статистика продаж']

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
})

print(df_result)

df_result = df_result.fillna('')

df_list = df_result.values.tolist()

# Найти последнюю заполненную строку в листе
last_row = len(worksheet.get_all_values()) + 1  # +1 чтобы начать запись с новой строки

# Запись данных в Google Таблицу с добавлением данных в конец
worksheet.update(range_name=f'A{last_row}', values=df_list)