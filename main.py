import os
import time
import logging
from commons.google_sheets_worker import GoogleSheetsClient
from report_processor.retail_sales_processor_1c_sellers import RetailSalesProcessor1CSellers
from report_processor.retail_sales_shop_1c import RetailSalesProcessorShop
from loader_reporting.loader import GmailAttachmentDownloader
from report_processor.ne_1c import RevenueProcessor, SalesReportProcessor
from tg_bot.alert import TelegramAlertBot

CONFIG = {
    'credentials_path': 'loader_reporting/credentials/credentials.json',
    'token_path': 'loader_reporting/credentials/token.json',
    'download_dir': 'downloads',
    'processed_files_path': 'loader_reporting/processed_files.json',
    'sheets_credentials_path': 'loader_reporting/credentials/mailapi-431104-8992c2888d0e.json',
    'spreadsheet_id_1': '',
    'worksheet_name_1': 'ИСТОРИЯ МАГАЗИНОВ',
    'worksheet_name_2': 'ИСТОРИЯ ПО ПК',
    'bot_token': "",
    'user_chat_id': '',  # Личный chat_id пользователя, куда будут отправляться сообщения
    'chat_id_otchet': ""  # Чат для отчетов
}

# Кастомный обработчик логов для отправки сообщений через Telegram в личные сообщения
class TelegramLoggingHandler(logging.Handler):
    def __init__(self, bot_token, chat_id):
        super().__init__()
        self.bot = TelegramAlertBot(CONFIG['bot_token'], CONFIG['user_chat_id'], CONFIG['sheets_credentials_path'], "1qqDlGHYDUy8Uv8Yf89S0aKnUngOOLVuERNBmfrHA3a0", CONFIG['chat_id_otchet'])

    def emit(self, record):
        log_entry = self.format(record)
        self.bot.send_message_alert(log_entry)

# Настройка логирования
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Добавляем кастомный обработчик для отправки логов через Telegram в личные сообщения
telegram_handler = TelegramLoggingHandler(CONFIG['bot_token'], CONFIG['user_chat_id'])
telegram_handler.setLevel(logging.INFO)

# Форматирование логов
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
telegram_handler.setFormatter(formatter)

# Добавляем обработчик в логгер
logger.addHandler(telegram_handler)

# Функции обработки файлов
def process_file_shved(file_path, alert_bot):
    try:
        processor_1 = RevenueProcessor(file_path, CONFIG['sheets_credentials_path'], CONFIG['spreadsheet_id_1'],
                                       CONFIG['worksheet_name_1'])
        processor_1.process()

        processor_2 = SalesReportProcessor(file_path, CONFIG['sheets_credentials_path'], CONFIG['spreadsheet_id_1'],
                                           CONFIG['worksheet_name_2'])
        processor_2.process()

        return True
    except Exception as e:
        handle_processing_error(file_path, e, alert_bot)
        return False

def process_file_1c(file_path, alert_bot):
    sheets_client = GoogleSheetsClient(CONFIG['sheets_credentials_path'], CONFIG['spreadsheet_id_1'])
    try:
        processor1cSellers = RetailSalesProcessor1CSellers(sheets_client, 3)
        processor1cSellers.process_and_update(file_path)

        processor1cShop = RetailSalesProcessorShop(sheets_client, 2)
        processor1cShop.process_and_update(file_path)

        return True
    except Exception as e:
        handle_processing_error(file_path, e, alert_bot)
        return False

def handle_processing_error(file_path, exception, alert_bot):
    file_name = os.path.basename(file_path)
    alert_message = f"❌Ошибка при обработке файла '{file_name}': {exception}"
    time.sleep(1)
    alert_bot.send_message_alert(alert_message)
    logging.error(alert_message)

def main():
    # Создание экземпляра Telegram бота для отправки алертов
    alert_bot = TelegramAlertBot(CONFIG['bot_token'], CONFIG['user_chat_id'], CONFIG['sheets_credentials_path'], "1qqDlGHYDUy8Uv8Yf89S0aKnUngOOLVuERNBmfrHA3a0", CONFIG['chat_id_otchet'])

    # Создание экземпляра класса для загрузки файлов
    downloader = GmailAttachmentDownloader(
        credentials_path=CONFIG['credentials_path'],
        token_path=CONFIG['token_path'],
        download_dir=CONFIG['download_dir'],
        processed_files_path=CONFIG['processed_files_path']
    )

    # Загрузка вложений
    downloader.download_attachments()

    # Обработка загруженных файлов
    for file_name in os.listdir(CONFIG['download_dir']):
        file_path = os.path.join(CONFIG['download_dir'], file_name)
        if os.path.isfile(file_path):
            time.sleep(1)  # Задержка перед обработкой каждого файла
            processed = False

            # Проверка имени файла и вызов соответствующей функции обработки
            if "LACRO" in file_name.upper() or "GUESS" in file_name.upper():
                processed = process_file_1c(file_path, alert_bot)
            else:
                processed = process_file_shved(file_path, alert_bot)

            if processed:
                os.remove(file_path)
                logging.info(f"Файл {file_name} успешно обработан и удален.")
            else:
                os.remove(file_path)
                logging.warning(f"Файл {file_name} не удалось обработать, он пропущен.")

    TelegramAlertBot.generate_and_send_report(alert_bot)
    TelegramAlertBot.generate_and_send_report_1c(alert_bot)

if __name__ == '__main__':
    main()
