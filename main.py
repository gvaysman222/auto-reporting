import os
import time  # Импортируем модуль time
from loader_reporting.loader import GmailAttachmentDownloader
from report_processor.ne_1c import RevenueProcessor, SalesReportProcessor  # Импортируем оба класса
from tg_bot.alert import TelegramAlertBot  # Подключаем наш новый модуль


def process_file(file_path, alert_bot):
    credentials_path = 'report_processor/credentials.json'
    spreadsheet_id_1 = '1T_xOCiDFiE8BWDsK_Iodo79fu2aLV_DqgZQrtbpwhU4'  # ID таблицы для RevenueProcessor
    worksheet_name_1 = 'ИСТОРИЯ МАГАЗИНОВ'

    spreadsheet_id_2 = '1T_xOCiDFiE8BWDsK_Iodo79fu2aLV_DqgZQrtbpwhU4'  # ID таблицы для SalesReportProcessor
    worksheet_name_2 = 'ИСТОРИЯ ПО ПК'

    try:
        # Используем RevenueProcessor для первой задачи
        processor_1 = RevenueProcessor(file_path, credentials_path, spreadsheet_id_1, worksheet_name_1)
        processor_1.process()

        # Используем SalesReportProcessor для второй задачи
        processor_2 = SalesReportProcessor(file_path, credentials_path, spreadsheet_id_2, worksheet_name_2)
        processor_2.process()

        return True
    except Exception as e:
        file_name = os.path.basename(file_path)
        alert_message = f"Ошибка при обработке файла '{file_name}': {e}"

        # Задержка перед отправкой сообщения
        time.sleep(1)

        alert_bot.send_message(alert_message)
        return False


def main():
    # Настройка параметров
    credentials_path = 'loader_reporting/credentials/credentials.json'
    token_path = 'loader_reporting/credentials/token.json'
    download_dir = os.path.join(os.path.dirname(__file__), 'downloads')
    processed_files_path = 'loader_reporting/processed_files.json'

    # Создание экземпляра класса для загрузки файлов
    downloader = GmailAttachmentDownloader(
        credentials_path=credentials_path,
        token_path=token_path,
        download_dir=download_dir,
        processed_files_path=processed_files_path
    )

    # Инициализация Telegram бота для отправки алертов
    bot_token = "5998611067:AAGAorkOfr0PRAn-vZWyUiKxWQ11MhsUUj8"
    chat_id = "-1002030942634"
    alert_bot = TelegramAlertBot(bot_token, chat_id)

    # Загрузка вложений
    downloader.download_attachments()

    # Обработка загруженных файлов
    for file_name in os.listdir(download_dir):
        file_path = os.path.join(download_dir, file_name)
        if os.path.isfile(file_path):
            # Задержка перед обработкой каждого файла
            time.sleep(1)

            processed = process_file(file_path, alert_bot)
            if processed:
                os.remove(file_path)
                print(f"Файл {file_name} успешно обработан и удален.")
            else:
                os.remove(file_path)
                print(f"Файл {file_name} не удалось обработать, он пропущен.")


if __name__ == '__main__':
    main()
