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
        try:
            worksheet = self.google_sheets_client.get_worksheet_by_name(sheet_name)
            data = worksheet.get_all_values()
        except Exception as e:
            self.send_message(f"Ошибка при получении данных из Google Sheets: {e}")
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
        # Получение текущей даты и расчет вчерашней даты
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        date_str = yesterday.strftime("%d.%m.%Y")

        day_of_week = days_of_week[yesterday.strftime("%A")]

        # Извлечение данных из Google Sheets
        sheet_name = "ИСТОРИЯ МАГАЗИНОВ"  # Укажите нужный лист
        rows = self.get_data_from_google_sheets(sheet_name, date_str)

        if not rows:
            self.send_message(f"Нет данных для отчета на {date_str}")
            return

        # Обработка всех строк, соответствующих заданной дате
        for data in rows:
            # Проверка, что данные содержат нужное количество столбцов
            if len(data) < 21:  # Убедитесь, что есть как минимум 21 элемент
                self.send_message(f"Ошибка: Недостаточно данных для отчета на {date_str}")
                continue

            # Формирование сообщения, используя индексы столбцов
            message = (
                f"Дата: {day_of_week} {yesterday.strftime('%d.%m.%Y')}\n\n"
                f"{data[0]} {data[17]}\n\n"  # ID и Магазин&Город (например, это столбцы A и R)
                f"История: {data[9]} \n"  # Выручка история
                f"План: {data[16]} \n"  # Выручка план день
                f"Факт: {data[2]} \n\n"  # Выручка
                f"Ср. чек: {data[5]}  / {data[12]} \n"  # Сумма среднего чека и Средний история+1+2
                f"Входы: {data[7]}  / {data[14]} \n"  # Входы и Входы история+1+2
                f"Upt: {data[6]}  / {data[13]} \n"  # UPT и UPT история+1+2
                f"Cnr: {data[8]} / {data[15]} \n\n"  # Конверсия и Конверсия история+1+2
                f"Выручка: {data[18]} \n"  # Прирост выручка
                f"Входы: {data[19]} \n"  # Прирост Входы
                f"Эффективность: {data[20]} "  # Эффективность
            )

            # Отправка сообщения через Telegram бот
            self.send_message(message)
