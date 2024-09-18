import telebot
from telebot import types
import subprocess
from loader_reporting.loader import GmailAttachmentDownloader  # Импорт класса для обновления токена
from main import main
import os

# Токен твоего бота
API_TOKEN = '5998611067:AAGAorkOfr0PRAn-vZWyUiKxWQ11MhsUUj8'

# Инициализация бота
bot = telebot.TeleBot(API_TOKEN)

CONFIG = {
    'credentials_path': 'loader_reporting/credentials/credentials.json',
    'token_path': 'loader_reporting/credentials/token.json',
    'download_dir': 'downloads',
    'processed_files_path': 'loader_reporting/processed_files.json',
    'sheets_credentials_path': 'loader_reporting/credentials/mailapi-431104-8992c2888d0e.json',
    'spreadsheet_id_1': '1kX591Zj4ZxdH4HI-G8eV7FM-YqsNMrzVJJiP9s4Ekro',
    'worksheet_name_1': 'ИСТОРИЯ МАГАЗИНОВ',
    'worksheet_name_2': 'ИСТОРИЯ ПО ПК',
    'bot_token': "5998611067:AAGAorkOfr0PRAn-vZWyUiKxWQ11MhsUUj8",
    'chat_id_alert': "-1002030942634",
    'chat_id_otchet': "-1002030942634"
}

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    # Создаем кнопки
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    update_token_button = types.KeyboardButton("Обновить токен")
    run_script_button = types.KeyboardButton("Запустить скрипт")
    markup.add(update_token_button, run_script_button)

    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)


# Обработка нажатий на кнопки
@bot.message_handler(func=lambda message: True)
def handle_buttons(message):
    if message.text == "Обновить токен":
        update_token(message)
    elif message.text == "Запустить скрипт":
        run_script(message)


# Функция для обновления токена
def update_token(message):
    try:
        # Удаление старого токена, если он существует
        if os.path.exists(CONFIG['token_path']):
            os.remove(CONFIG['token_path'])

        # Создание нового объекта GmailAttachmentDownloader перед обновлением токена
        downloader = GmailAttachmentDownloader(
            credentials_path=CONFIG['credentials_path'],
            token_path=CONFIG['token_path'],
            download_dir=CONFIG['download_dir'],
            processed_files_path=CONFIG['processed_files_path']
        )

        # Обновление токена
        downloader.authenticate_gmail()
        bot.send_message(message.chat.id, "Токен обновлен успешно!")
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка при обновлении токена: {str(e)}")


# Функция для запуска скрипта
def run_script(message):
    bot.send_message(message.chat.id, "Скрипт успешно запущен!")
    try:
        main()
        bot.send_message(message.chat.id, "Скрипт успешно отработал.")
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка при запуске скрипта: {str(e)}")


# Запуск бота
bot.polling(none_stop=True)
