import datetime
import telebot
from commons.google_sheets_worker import GoogleSheetsClient


class TelegramAlertBot:
    def __init__(self, bot_token, chat_id, json_keyfile_name, spreadsheet_id):
        self.bot = telebot.TeleBot(bot_token)
        self.chat_id = chat_id
        self.google_sheets_client = GoogleSheetsClient(json_keyfile_name, spreadsheet_id)

    def send_message(self, text):
        try:
            self.bot.send_message(self.chat_id, text)
        except Exception as e:
            print(f"Не удалось отправить сообщение: {e}")

    def get_data_from_google_sheets(self, sheet_name, date):
        worksheet = self.google_sheets_client.get_worksheet_by_name(sheet_name)

        # Получаем все данные из листа
        data = worksheet.get_all_records()

        # Фильтруем данные по дате
        for row in data:
            if row['Дата'] == date:
                return row

        return None

    def generate_and_send_report(self):
        # Получение текущей даты и расчет вчерашней даты
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        date_str = yesterday.strftime("%d.%m.%Y")

        # Извлечение данных из Google Sheets
        sheet_name = "ИСТОРИЯ МАГАЗИНОВ"  # Укажите нужный лист
        data = self.get_data_from_google_sheets(sheet_name, date_str)

        if not data:
            self.send_message(f"Нет данных для отчета на {date_str}")
            return

        for row in data:
            message = (
                f"Дата: {yesterday.strftime('%A (%V) %d.%m.%Y (B)')}\n\n"
                f"{row['ID']} (A) {row['Магазин/город']} (R)\n\n"
                f"История: {row['Выручка история']} (J)\n"
                f"План: {row['Выручка план день']} (Q)\n"
                f"Факт: {row['Выручка']} (C)\n\n"
                f"Ср. чек: {row['Сумма среднего чека']} (F) / {row['Средний история+1+2']} (M)\n"
                f"Входы: {row['Входы']} (H) / {row['Входы история+1+2']} (O)\n"
                f"Upt: {row['UPT']} (G) / {row['UPT история+1+2']} (N)\n"
                f"Cnr: {row['Конверсия']} (I) / {row['Конверсия история+1+2']} (P)\n\n"
                f"Выручка: {row['Прирост выручка']} (S)\n"
                f"Входы: {row['Прирост Входы']} (T)\n"
                f"Эффективность: {row['Эффективность']} (U)"
            )

            # Отправка сообщения через Telegram бот
            self.send_message(message)
