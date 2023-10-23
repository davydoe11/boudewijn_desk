import io
import shutil
import time
import google_auth_oauthlib.flow
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import pendulum
from datetime import datetime, timedelta
import pytz

def get_table():
    index_begin = 0
    index_end = 0
    output_string = "test"
    today = pendulum.now()
    start = today.start_of('week')
    week_ago = start - timedelta(days=7)
    # Remove the timezone offset and format the pendulum datetime object
    output_string = start.to_iso8601_string().split("+")[0].replace('T', ' ')
    output_string = start.strftime('%Y-%m-%d %H:%M:%S')
    output_week = week_ago.to_iso8601_string().split("+")[0].replace('T', ' ')
    output_week = week_ago.strftime('%Y-%m-%d %H:%M:%S')

    df = pd.DataFrame(pd.read_excel("spreadsheat.xlsx"))
    df = df.dropna(how='all')
    df = df[['2 jan t/m 6 jan', 'Unnamed: 2', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']]

    df.rename(columns={'2 jan t/m 6 jan': 'Datum', 'Unnamed: 2': 'Persoon'}, inplace=True)

    date_format = "%Y-%m-%d %H:%M:%S"
    df['Datum'] = df['Datum'].apply(lambda x: x.strftime(date_format) if isinstance(x, datetime) else x)
    # Convert the datetime.datetime objects to strings with the specified format
    index_begin = df.index[df['Datum'] == output_week]
    print(index_begin)
    index_end = df.index[df['Datum'] == output_string]
    index_begin = index_begin.to_numpy()[0]
    index_end = index_end.to_numpy()[0]


    df = df.drop(df.loc[0:index_begin].index)
    df = df.drop(df.loc[index_end:df.tail(1).index[0]].index)
    df = df.loc[df['Persoon'] == "Boudewijn"]
    return df

def day_checker(df):
    now = datetime.datetime.now()
    day = now.strftime("%A")
    return day

def download():
    f = open('id.txt', 'r')
    file_id = f.readline()
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

# The scope your app will use.
# (NEEDS to be among the enabled in your OAuth consent screen)
SCOPES = "https://www.googleapis.com/auth/drive.readonly"
CLIENT_SECRET_FILE = "credentialsmaxem.json"

credentials = Auth(client_secret_filename=CLIENT_SECRET_FILE, scopes=SCOPES).get_credentials()

drive_service = build('drive', 'v3', credentials=credentials)
while True:
    download()
    time.sleep(20)
    df = get_table()
    print(df.head)
