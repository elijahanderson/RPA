from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build
import os


def service_act_login():
    """ Log into gmail service account. """
    SCOPES = ['https://www.googleapis.com/auth/drive.metadata',
              'https://www.googleapis.com/auth/drive.file',
              'https://www.googleapis.com/auth/drive',
              ]
    SERVICE_ACCOUNT_FILE = 'src/config/service_key.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
    return service


def upload_file(filename, folder_id):
    print('Uploading file to drive...', end=' ')
    service = service_act_login()
    file_metadata = {
        'name': filename.split('/')[-1],
        'parents': [folder_id]
    }
    media = MediaFileUpload(filename, resumable=True)
    service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print('File uploaded.')


def upload_folder(folder_path, parent_folder_id):
    print('Uploading folder to drive...', end=' ')
    service = service_act_login()
    # create folder
    folder_metadata = {
        'name': folder_path.split('/')[-1],
        'parents': [parent_folder_id],
        'mimeType': 'application/vnd.google-apps.folder'
    }
    drive_folder = service.files().create(body=folder_metadata, fields='id').execute()
    folder_id = drive_folder.get('id')
    # upload its contents
    for file in os.listdir(folder_path):
        file_metadata = {
            'name': file,
            'parents': [folder_id]
        }
        media = MediaFileUpload(os.path.join(folder_path, file), resumable=True)
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print('Folder uploaded.')
