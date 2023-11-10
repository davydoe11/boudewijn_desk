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
import discord_notify as dn
import schedule

#Url of discord webhook bot
URL = 
notifier = dn.Notifier(URL)
blacklist = ["Saturday", "Sunday"]
persoon = "Boudewijn"


#Funtion for cleaning table to only see this weeks boudewijn stats.
def get_table():
    index_begin = 0
    index_end = 0
    output_string = "test"
    
    #Automatting the selction for each week
    today = pendulum.now()
    start = today.start_of('week')
    week_ago = start - timedelta(days=7)
    # Remove the timezone offset and format the pendulum datetime object
    output_string = start.to_iso8601_string().split("+")[0].replace('T', ' ')
    output_string = start.strftime('%Y-%m-%d %H:%M:%S')
    output_week = week_ago.to_iso8601_string().split("+")[0].replace('T', ' ')
    output_week = week_ago.strftime('%Y-%m-%d %H:%M:%S')

    #reading the downloaded spreedsheet
    df = pd.DataFrame(pd.read_excel("spreadsheat.xlsx"))
    df = df.dropna(how='all')
    df = df[['2 jan t/m 6 jan', 'Unnamed: 2', 'Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag']]

    #renaming for easier data usage
    df.rename(columns={'2 jan t/m 6 jan': 'Datum', 'Unnamed: 2': 'Persoon', 'Maandag' : 'Monday' , 'Dinsdag' : 'Tuesday', 'Woensdag' : 'Wednesday' , 'Donderdag' : 'Thursday', 'Vrijdag' : 'Friday' }, inplace=True)
    date_format = "%Y-%m-%d %H:%M:%S"
    df['Datum'] = df['Datum'].apply(lambda x: x.strftime(date_format) if isinstance(x, datetime) else x)
    
    # Convert the datetime.datetime objects to strings with the specified format, and formatting table to only this week.
    index_begin = df.index[df['Datum'] == output_week]
    index_end = df.index[df['Datum'] == output_string]
    index_begin = index_begin.to_numpy()[0]
    index_end = index_end.to_numpy()[0]
    df = df.drop(df.loc[0:index_begin].index)
    df = df.drop(df.loc[index_end:df.tail(1).index[0]].index)
    df = df.loc[df['Persoon'] == persoon]
    return df

#fucntion for daily day checking, and printing the corresponding response
def day_checker(df, day):
    val = df[day].values[0]
    if (val == "Present"):
        notifier.send("Boudewijn is vandaag op Kantoor 🥲", print_message=False)
    else:      
        notifier.send("Bureau is vrij gek!!! 🦅", print_message=False)
        
#Using the google drive APi we download the most recent spreadsheet
def download():
    f = open('id.txt', 'r')
    file_id = f.readline()
    request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    # The file has been downloaded into RAM, now save it in a file
    fh.seek(0)
    with open('spreadsheat.xlsx', 'wb') as f:
        shutil.copyfileobj(fh, f, length=131072)

#The function to run once the time has come
def job():
    now = datetime.now()
    day = now.strftime("%A")
    if day not in blacklist:
        download()
        time.sleep(1)
        df = get_table()
        day_checker(df, day)
    

#one time setup for use with the Discord bot and Google API
class Auth:
    def __init__(self, client_secret_filename, scopes):
        schedule.every().day.at("08:50").do(job)
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

#Credentials saved in the specific file.
credentials = Auth(client_secret_filename=CLIENT_SECRET_FILE, scopes=SCOPES).get_credentials()

drive_service = build('drive', 'v3', credentials=credentials)
while True:
    schedule.run_pending()
    time.sleep(60) # wait one minute
   
