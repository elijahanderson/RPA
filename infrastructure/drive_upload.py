from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build


def service_act_login():
    """ Log into gmail service account. """
    SCOPES = ['https://www.googleapis.com/auth/drive.metadata',
              'https://www.googleapis.com/auth/drive.file',
              'https://www.googleapis.com/auth/drive',
              ]
    SERVICE_ACCOUNT_FILE = '../data/service_key.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
    return service


def upload_file(filename, folder_id):
    service = service_act_login()
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(filename, resumable=True)
    file = service.files().create(body=file_metadata,
                                  media_body=media,
                                  fields='id').execute()
