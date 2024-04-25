import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from responseReading import read
from dotenv import load_dotenv 
import json

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive']
load_dotenv()
client_secrets=os.getenv('client_secrets')
client_secrets=json.loads(client_secrets)

def authenticate():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        # flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
        flow = InstalledAppFlow.from_client_config(client_secrets, SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def download_file(file_id, dest_path, creds):
    service = build('drive', 'v3', credentials=creds)
    request = service.files().get_media(fileId=file_id)
    fh = open(dest_path, 'wb')
    downloader = request.execute()
    fh.write(downloader)
    fh.close()
    print(f"File downloaded to: {dest_path}")

def removeAllFiles(directory_path):
    try:
     files = os.listdir(directory_path)
     for file in files:
       file_path = os.path.join(directory_path, file)
       if os.path.isfile(file_path):
         os.remove(file_path)
     print("All files deleted successfully.")
    except OSError:
     print("Error occurred while deleting files.")

def downloadFiles():
    removeAllFiles(os.path.join("Excel"))
    removeAllFiles(os.path.join("Images"))
    creds = authenticate()
    dataList = read()
    channel = dataList[3]
    video = dataList[4]
    logo = dataList[5]
    download_file(channel, 'Excel/channel.xlsx', creds)
    download_file(video, 'Excel/video.xlsx', creds)
    download_file(logo, 'logo.png', creds)
