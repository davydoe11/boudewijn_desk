import io
import shutil
import time
import google_auth_oauthlib.flow
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd 

def download():

    f = open('id.txt','r')
    file_id= f.readline()
    request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%" % int(status.progress() * 100))

    # The file has been downloaded into RAM, now save it in a file
    fh.seek(0)
    with open('spreadsheat.xlsx', 'wb') as f:
        shutil.copyfileobj(fh, f, length=131072)
        
        
    


class Auth:
 
    def __init__(self, client_secret_filename, scopes):
        self.client_secret = client_secret_filename
        self.scopes = scopes
        self.flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(self.client_secret, self.scopes)
        self.flow.redirect_uri = 'http://localhost:8080/'
        self.creds = None
 
    def get_credentials(self):
        flow = InstalledAppFlow.from_client_secrets_file(self.client_secret, self.scopes)
        self.creds = flow.run_local_server(port=8080)
        return self.creds

    
# The scope you app will use. 
# (NEEDS to be among the enabled in your OAuth consent screen)
SCOPES = "https://www.googleapis.com/auth/drive.readonly"
CLIENT_SECRET_FILE = "credentialsmaxem.json"
 
credentials = Auth(client_secret_filename=CLIENT_SECRET_FILE, scopes=SCOPES).get_credentials()
 
drive_service = build('drive', 'v3', credentials=credentials)
while True:
    download()
    time.sleep(20)
    df = pd.DataFrame(pd.read_excel("spreadsheat.xlsx"))
    print(df)


        