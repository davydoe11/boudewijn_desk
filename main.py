import time
from googleapiclient.discovery import build
import pandas as pd
import pendulum
from datetime import datetime, timedelta
import discord_notify as dn
import schedule
from google.oauth2 import service_account

# Url of discord webhook bot
person = "Boudewijn"
f = open('url.txt', 'r')
URL = f.readline()
notifier = dn.Notifier(URL)
blacklist = ["Saturday", "Sunday"]


# Funtion for cleaning table to only see this weeks boudewijn stats.
def get_table():
    index_begin = 0
    index_end = 0
    output_string = "test"

    # Automatting the selction for each week
    today = pendulum.now()
    start = today.start_of('week')
    print("start of week:", start)
    # Format the date as "DD MMM" (e.g., "11 Dec")
    formatted_date = start.format('DD MMM').lower()

    print("Formatted date:", formatted_date)

    week_ago = start - timedelta(days=7)
    print("week ago: ", week_ago)
    output_string = start.strftime('%Y-%m-%d %H:%M:%S')
    output_week = week_ago.strftime('%Y-%m-%d %H:%M:%S')

    # reading the downloaded spreedsheet
    df = pd.DataFrame(pd.read_excel("spreadsheat.xlsx"))
    df = df.dropna(how='all')
    df = df[['1 jan t/m 5 jan', 'Unnamed: 2', 'Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag']]
    # renaming for easier data usage
    df.rename(columns={'1 jan t/m 5 jan': 'Datum', 'Unnamed: 2': 'Persoon', 'Maandag': 'Monday', 'Dinsdag': 'Tuesday',
                       'Woensdag': 'Wednesday', 'Donderdag': 'Thursday', 'Vrijdag': 'Friday'}, inplace=True)

    date_format = "%Y-%m-%d %H:%M:%S"
    df['Datum'] = df['Datum'].apply(lambda x: x.strftime(date_format) if isinstance(x, datetime) else x)

    # Convert the datetime.datetime objects to strings with the specified format, and formatting table to only this week.
    print(df.to_string())
    print("formatted date", formatted_date)
    indexes_of_date = df[df['Datum'].notna() & df['Datum'].str.contains(formatted_date)].index.tolist()
    try:
        index_begin = indexes_of_date[-1]
    except:
        print("Start date not found")
        index_begin = 0
    print("index begin: ", index_begin)
    index_end = df.index[df['Datum'] == output_string]
    print("index end", index_end)
    index_end = index_end[-1]


    df = df.drop(df.loc[0:index_begin].index)
    df = df.drop(df.loc[index_end:df.tail(1).index[0]].index)
    df = df.loc[df['Persoon'] == person]
    return df


# fucntion for daily day checking, and printing the corresponding response
def day_checker(df, day, person):
    val = df[day].values[0]
    if (val == "Present"):
        notifier.send(f"{person} is vandaag op Kantoor 🥲", print_message=False)
    else:
        notifier.send(f"Bureau {person} is vrij gek!!! 🦅", print_message=False)


# Using the google drive APi we download the most recent spreadsheet
def download():
    f = open('id.txt', 'r')
    file_id = f.readline()
    credentials = service_account.Credentials.from_service_account_file('credentials-service.json', scopes=[
        'https://www.googleapis.com/auth/drive.readonly'])

    # Replace 'your_spreadsheet_id' with the actual ID of your spreadsheet
    spreadsheet_id = file_id

    # Create a Google Drive API service
    service = build('drive', 'v3', credentials=credentials)

    # Download the spreadsheet
    request = service.files().export_media(fileId=spreadsheet_id,
                                           mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response = request.execute()

    # Save the spreadsheet content to a file
    with open('spreadsheat.xlsx', 'wb') as f:
        f.write(response)


# The function to run once the time has come
def job():
    now = datetime.now()
    day = now.strftime("%A")
    if day not in blacklist:
        download()
        time.sleep(1)
        df = get_table()
        day_checker(df, day, person)
        print(f"Send Succesfull {now}")


# one time setup for use with the Discord bot and Google API
class Main:
    now = datetime.now()
    print(f"Just rebooted: {now}")
    schedule.every().day.at("10:22").do(job)
    # job()
    while True:
        schedule.run_pending()
        time.sleep(60)  # wait one minute
