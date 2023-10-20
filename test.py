import numpy
import pandas as pd  
import pendulum
from datetime import datetime, timedelta
index_begin = 0
index_end = 0
output_string = "test"
import pytz


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
df = df[['2 jan t/m 6 jan','Unnamed: 2', 'Maandag','Dinsdag','Woensdag','Donderdag','Vrijdag']]



df.rename(columns={'2 jan t/m 6 jan': 'Datum', 'Unnamed: 2': 'Persoon'}, inplace=True)

date_format = "%Y-%m-%d %H:%M:%S"
df['Datum'] = df['Datum'].apply(lambda x: x.strftime(date_format) if isinstance(x, datetime) else x)
# Convert the datetime.datetime objects to strings with the specified format
index_begin = df.index[df['Datum']== output_week]
print(index_begin)
index_end = df.index[df['Datum']==output_string]
index_begin= index_begin.to_numpy()
index_end = index_end.to_numpy()

df = df.drop(df.loc[0:(int(index_begin))].index)
df = df.drop(df.loc[int(index_end):(df.tail(1).index.item())].index)
df = df.loc[df['Persoon'] == "Boudewijn"]
print(df.to_markdown())



print(output_string)



