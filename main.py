import os
from loader_reporting.loader import GmailAttachmentDownloader
def main():

    # Настройка параметров
    credentials_path = 'loader_reporting/credentials/credentials.json'
    token_path = 'loader_reporting/credentials/token.json'
    download_dir = os.path.join(os.path.dirname(__file__), 'downloads')
    processed_files_path = 'loader_reporting/processed_files.json'

    # Создание экземпляра класса
    downloader = GmailAttachmentDownloader(
        credentials_path=credentials_path,
        token_path=token_path,
        download_dir=download_dir,
        processed_files_path=processed_files_path
    )

    # Загрузка вложений
    downloader.download_attachments()

if __name__ == '__main__':
    main()
