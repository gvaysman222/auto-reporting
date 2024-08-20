import os
import base64
import json
import time
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


class GmailAttachmentDownloader:
    def __init__(self, credentials_path, token_path, download_dir, processed_files_path, scopes=None):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.download_dir = download_dir
        self.processed_files_path = processed_files_path
        self.scopes = scopes or ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = None
        self.creds = None

        # Создание папки для загрузки файлов, если она не существует
        if not os.path.exists(self.download_dir):
            os.makedirs(self.download_dir)

        self.authenticate_gmail()

    def authenticate_gmail(self):
        # Загрузка токена, если он уже существует
        if os.path.exists(self.token_path):
            self.creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)

        # Если токен недействителен, или не существует, запускаем процесс авторизации
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.scopes)
                self.creds = flow.run_local_server(port=0)

            # Сохраняем новый токен в файл
            with open(self.token_path, 'w') as token_file:
                token_file.write(self.creds.to_json())

        # Инициализация сервиса Gmail API
        self.service = build('gmail', 'v1', credentials=self.creds)

    def load_processed_files(self):
        # Загрузка списка обработанных файлов
        if os.path.exists(self.processed_files_path):
            with open(self.processed_files_path, 'r') as file:
                return json.load(file)
        return []

    def save_processed_file(self, file_id):
        processed_files = self.load_processed_files()
        processed_files.append(file_id)
        with open(self.processed_files_path, 'w') as file:
            json.dump(processed_files, file)

    def save_attachment(self, part, msg_id, subject):
        if 'filename' in part and part['filename']:
            filename = part['filename']
            file_data = part['body'].get('data')
            if 'attachmentId' in part['body']:
                attachment = self.service.users().messages().attachments().get(
                    userId='me', messageId=msg_id, id=part['body']['attachmentId']
                ).execute()
                file_data = attachment['data']

            # Преобразуем тему письма в допустимое имя файла
            safe_subject = "".join([c if c.isalnum() or c in " ._-" else "_" for c in subject])

            # Создаем новое имя файла на основе темы письма
            new_filename = f"{safe_subject}_{filename}"
            file_path = os.path.join(self.download_dir, new_filename)

            if not os.path.exists(file_path):
                with open(file_path, 'wb') as f:
                    f.write(base64.urlsafe_b64decode(file_data))
                print(f"Файл {new_filename} сохранен в {file_path}")
                self.save_processed_file(msg_id)
            else:
                print(f"Файл {new_filename} уже был скачан, пропускаем.")

    def process_parts(self, parts, msg_id, subject):
        for part in parts:
            if part['mimeType'] == 'multipart/alternative' or part['mimeType'] == 'multipart/mixed':
                self.process_parts(part['parts'], msg_id, subject)
            else:
                self.save_attachment(part, msg_id, subject)

    def download_attachments(self):
        processed_files = self.load_processed_files()

        # Получаем список писем с вложениями
        results = self.service.users().messages().list(userId='me', q="has:attachment").execute()
        messages = results.get('messages', [])

        if not messages:
            print("Нет писем с вложениями.")
            return

        for message in messages:
            msg_id = message['id']

            # Проверяем, было ли это письмо уже обработано
            if msg_id in processed_files:
                print(f"Письмо с ID {msg_id} уже обработано, пропускаем.")
                continue

            msg = self.service.users().messages().get(userId='me', id=msg_id).execute()

            subject = ""
            headers = msg['payload'].get('headers', [])
            for header in headers:
                if header['name'].lower() == 'subject':
                    subject = header['value']
                    break

            if 'parts' in msg['payload']:
                self.process_parts(msg['payload']['parts'], msg_id, subject)
