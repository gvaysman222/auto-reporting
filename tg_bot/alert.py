import telebot

class TelegramAlertBot:
    def __init__(self, bot_token, chat_id):
        self.bot = telebot.TeleBot(bot_token)
        self.chat_id = chat_id

    def send_message(self, text):
        try:
            self.bot.send_message(self.chat_id, text)
        except Exception as e:
            print(f"Не удалось отправить сообщение: {e}")