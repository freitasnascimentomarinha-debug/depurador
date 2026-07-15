"""
Funções para autenticar e ler arquivos de uma pasta do Google Drive
usando uma Service Account (conta de serviço).
"""
import io
import re

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

SUPPORTED_EXTENSIONS = ('.pdf', '.docx', '.doc', '.xlsx', '.xls')


def get_drive_service(service_account_json_path):
    """Cria o cliente autenticado da API do Google Drive."""
    creds = service_account.Credentials.from_service_account_file(
        service_account_json_path, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


def extract_folder_id(folder_id_or_url: str) -> str:
    """Aceita tanto o ID puro da pasta quanto o link completo do Drive."""
    folder_id_or_url = folder_id_or_url.strip()
    if 'drive.google.com' in folder_id_or_url:
        match = re.search(r'/folders/([a-zA-Z0-9_-]+)', folder_id_or_url)
        if match:
            return match.group(1)
    return folder_id_or_url


def list_files_in_folder(service, folder_id: str):
    """Lista todos os arquivos suportados dentro da pasta (sem subpastas)."""
    query = f"'{folder_id}' in parents and trashed = false"
    files = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType, modifiedTime)',
            pageToken=page_token
        ).execute()
        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken')
        if page_token is None:
            break

    return [f for f in files if f['name'].lower().endswith(SUPPORTED_EXTENSIONS)]


def download_file(service, file_id: str, dest_path: str) -> str:
    """Baixa um arquivo do Drive para o caminho local informado."""
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(dest_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    return dest_path
