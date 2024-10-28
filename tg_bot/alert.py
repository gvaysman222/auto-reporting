import datetime
import json
import os
import telebot
from commons.google_sheets_worker import GoogleSheetsClient

class TelegramAlertBot:
    def __init__(self, bot_token, chat_id_alert, json_keyfile_name, spreadsheet_id, chat_id_otchet):
        self.bot = telebot.TeleBot(bot_token)
        self.chat_id_alert = chat_id_alert
        self.google_sheets_client = GoogleSheetsClient(json_keyfile_name, spreadsheet_id)
        self.chat_id_otchet = chat_id_otchet
        self.processed_reports_file = "processed_reports.json"
        self.load_processed_reports()

    def load_processed_reports(self):
        """Загружает обработанные отчеты из файла, если он существует."""
        if os.path.exists(self.processed_reports_file):
            with open(self.processed_reports_file, 'r') as file:
                self.processed_reports = json.load(file)
        else:
            self.processed_reports = []

    def save_processed_report(self, store_code, report_date):
        """Сохраняет новый обработанный отчет в формате {'store_code': 'IDY0', 'date': 'YYYY-MM-DD'}."""
        self.processed_reports.append({
            "store_code": store_code,
            "date": report_date
        })
        with open(self.processed_reports_file, 'w') as file:
            json.dump(self.processed_reports, file)

    def is_report_processed(self, store_code, report_date):
        """Проверяет, был ли отчет по данному магазину и дате уже обработан."""
        for report in self.processed_reports:
            if report["store_code"] == store_code and report["date"] == report_date:
                return True
        return False

    def send_message_alert(self, text):
        try:
            self.bot.send_message(self.chat_id_alert, text)
        except Exception as e:
            print(f"Не удалось отправить сообщение: {e}")

    def send_message_otchet(self, text):
        try:
            self.bot.send_message(self.chat_id_otchet, text)
        except Exception as e:
            print(f"Не удалось отправить сообщение: {e}")

    def get_data_from_google_sheets(self, sheet_name, date):
        try:
            worksheet = self.google_sheets_client.get_worksheet_by_name(sheet_name)
            data = worksheet.get_all_values()
        except Exception as e:
            self.send_message_alert(f"Ошибка при получении данных из Google Sheets: {e}")
            return []

        # Список для хранения всех строк, соответствующих заданной дате
        matching_rows = []
        for row in data:
            if row[1] == date:  # Предполагаем, что столбец "Дата" - это второй столбец, индекс 1
                matching_rows.append(row)

        return matching_rows

    def generate_and_send_report(self):
        days_of_week = {
            "Monday": "Понедельник",
            "Tuesday": "Вторник",
            "Wednesday": "Среда",
            "Thursday": "Четверг",
            "Friday": "Пятница",
            "Saturday": "Суббота",
            "Sunday": "Воскресенье"
        }

        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        date_str = yesterday.strftime("%d.%m.%Y")
        day_of_week = days_of_week[yesterday.strftime("%A")]

        # Извлечение данных из Google Sheets
        sheet_name = "ИСТОРИЯ МАГАЗИНОВ"
        rows = self.get_data_from_google_sheets(sheet_name, date_str)

        if not rows:
            self.send_message_alert(f"Нет данных для отчета на {date_str}")
            return

        # Обработка всех строк, соответствующих заданной дате
        for data in rows:
            store_code = data[0]  # Предполагаем, что код магазина в первом столбце

            # Проверяем, был ли отчет уже обработан
            if self.is_report_processed(store_code, date_str):
                continue  # Пропускаем отчет, если он уже был отправлен

            if len(data) < 21:
                self.send_message_alert(f"Ошибка: Недостаточно данных для отчета на {date_str}")
                continue

            # Формирование сообщения
            message = (
                "✅Отчет✅\n\n"
                f"Дата: {day_of_week} {yesterday.strftime('%d.%m.%Y')}\n\n"
                f"{data[0]} {data[17]}\n\n"
                f"История: {data[9]} \n"
                f"План: {data[16]} \n"
                f"Факт: {data[2]} \n\n"
                f"Ср. чек: {data[5]}  / {data[12]} \n"
                f"Входы: {data[7]}  / {data[14]} \n"
                f"Кол-во товаров: {data[3]}  / {data[10]} \n"
                f"Кол-во чеков: {data[4]} / {data[11]} \n"
                f"Upt: {data[6]}  / {data[13]} \n"
                f"Cnr: {data[8]} / {data[15]} \n\n"
                f"Выручка: {data[18]} \n"
                f"Входы: {data[19]} \n"
                f"Эффективность: {data[20]} "
            )

            # Отправка сообщения через Telegram бот
            self.send_message_otchet(message)

            # Сохраняем отчет как обработанный
            self.save_processed_report(store_code, date_str)

    def generate_and_send_report_1c(self):
        days_of_week = {
            "Monday": "Понедельник",
            "Tuesday": "Вторник",
            "Wednesday": "Среда",
            "Thursday": "Четверг",
            "Friday": "Пятница",
            "Saturday": "Суббота",
            "Sunday": "Воскресенье"
        }

        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        date_str = yesterday.strftime("%d.%m.%Y")
        day_of_week = days_of_week[yesterday.strftime("%A")]

        # Извлечение данных из Google Sheets
        sheet_name = "ИСТОРИЯ МАГАЗИНОВ 1С"
        rows = self.get_data_from_google_sheets(sheet_name, date_str)

        if not rows:
            self.send_message_alert(f"Нет данных для отчета на {date_str} по Guess | Lacro")
            return

        # Обработка всех строк, соответствующих заданной дате
        for data in rows:
            store_code = data[0]

            # Проверяем, был ли отчет уже обработан
            if self.is_report_processed(store_code, date_str):
                continue  # Пропускаем отчет, если он уже был отправлен

            if len(data) < 21:
                self.send_message_alert(f"Ошибка: Недостаточно данных для отчета на {date_str}")
                continue

            # Формирование сообщения
            message = (
                "✅Отчет✅\n\n"
                f"Дата: {day_of_week} {yesterday.strftime('%d.%m.%Y')}\n\n"
                f"{data[0]} {data[17]}\n\n"
                f"История: {data[9]} \n"
                f"План: {data[16]} \n"
                f"Факт: {data[2]} \n\n"
                f"Ср. чек: {data[5]}  / {data[12]} \n"
                f"Входы: {data[7]}  / {data[14]} \n"
                f"Кол-во товаров: {data[3]} / {data[10]} \n"
                f"Кол-во чеков: {data[4]} / {data[11]} \n"
                f"Upt: {data[6]}  / {data[13]} \n"
                f"Cnr: {data[8]} / {data[15]} \n\n"
                f"Выручка: {data[18]} \n"
                f"Входы: {data[19]} \n"
                f"Эффективность: {data[20]} "
            )

            # Отправка сообщения через Telegram бот
            self.send_message_otchet(message)

            # Сохраняем отчет как обработанный
            self.save_processed_report(store_code, date_str)


