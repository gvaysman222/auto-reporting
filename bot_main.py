import telebot
from telebot import types
import subprocess
from loader_reporting.loader import GmailAttachmentDownloader  # Импорт класса для обновления токена
from main import main
import os

# Токен твоего бота
API_TOKEN = ''

# Инициализация бота
bot = telebot.TeleBot(API_TOKEN)

# Список разрешённых пользователей
ALLOWED_USERS = []  # Замените на реальные Telegram ID

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
    'chat_id_otchet': "  # Чат для отчетов
}


# Функция проверки пользователя
def is_allowed_user(user_id):
    return user_id in ALLOWED_USERS


# Проверка, является ли чат личным
def is_private_chat(message):
    return message.chat.type == 'private'


# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    if not is_private_chat(message):
        return  # Игнорируем команды из групп и каналов

    if not is_allowed_user(message.from_user.id):
        bot.send_message(message.chat.id, "У вас нет доступа к этому боту.")
        return

    # Создаем кнопки
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    update_token_button = types.KeyboardButton("Обновить токен")
    run_script_button = types.KeyboardButton("Запустить скрипт")
    markup.add(update_token_button, run_script_button)

    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)


# Обработка нажатий на кнопки
@bot.message_handler(func=lambda message: True)
def handle_buttons(message):
    if not is_private_chat(message):
        return  # Игнорируем сообщения из групп и каналов

    if not is_allowed_user(message.from_user.id):
        bot.send_message(message.chat.id, "У вас нет доступа к этому боту.")
        return

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
