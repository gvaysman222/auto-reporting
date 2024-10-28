import os
import base64
import json
import time
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

class GmailAttachmentDownloader:
    def __init__(self, credentials_path, token_path, download_dir, processed_files_path, scopes=None):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.download_dir = download_dir
        self.processed_files_path = processed_files_path
        self.scopes = scopes or ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = None
        self.creds = None

        if not os.path.exists(self.download_dir):
            os.makedirs(self.download_dir)

        self.authenticate_gmail()

    def open_browser(self, auth_url):
        """Открывает браузер через Selenium и автоматически вводит логин и пароль"""
        chrome_options = ChromeOptions()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        driver = webdriver.Chrome(options=chrome_options)

        # Открываем URL для авторизации
        driver.get(auth_url)

        # Вводим логин и пароль через Selenium
        time.sleep(5)
        email_input = driver.find_element(By.ID, "identifierId")
        email_input.send_keys("analyticsuu@gmail.com")
        driver.find_element(By.ID, "identifierNext").click()

        time.sleep(5)
        password_input = driver.find_element(By.NAME, "Passwd")
        password_input.send_keys("UUTim_24")
        driver.find_element(By.ID, "passwordNext").click()



        wait = WebDriverWait(driver, 3)
        continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Продолжить']")))
        actions = ActionChains(driver)
        actions.move_to_element(continue_button).click().perform()

        try:
            wait = WebDriverWait(driver, 5)
            continue_button_2 = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Продолжить']]")))

            driver.execute_script("arguments[0].scrollIntoView(true);", continue_button_2)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", continue_button_2)
            print("Успешно кликнули на вторую кнопку 'Продолжить' с помощью JavaScript.")

        except StaleElementReferenceException:
            print("Элемент был изменен, повторяем попытку поиска...")
            continue_button_2 = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Продолжить']]")))
            driver.execute_script("arguments[0].click();", continue_button_2)

        except Exception as e:
            print(f"Не удалось нажать на вторую кнопку 'Продолжить': {e}")

        # Ожидаем появления кода авторизации в URL
        while "code=" not in driver.current_url:
            time.sleep(1)

        # Получаем код авторизации из URL
        auth_code = driver.current_url.split("code=")[1]

        # Закрываем браузер
        driver.quit()

        return auth_code

    def authenticate_gmail(self):
        if os.path.exists(self.token_path):
            self.creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path,
                    self.scopes,
                    redirect_uri='http://localhost:8000/oauth2callback'  # Указываем redirect_uri
                )
                auth_url, _ = flow.authorization_url(prompt='consent')

                # Открываем браузер для автоматизации через Selenium
                auth_code = self.open_browser(auth_url)

                # Получаем токен используя код авторизации
                flow.fetch_token(code=auth_code)
                self.creds = flow.credentials  # Получаем credentials объект

            # Сохраняем токен в файл
            with open(self.token_path, 'w') as token_file:
                token_file.write(self.creds.to_json())

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
