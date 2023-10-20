import numpy
import pandas as pd  
value = 0
index_end = 0

df = pd.DataFrame(pd.read_excel("spreadsheat.xlsx"))
print(df)
df = df.dropna(how='all')
print(df.to_markdown())
df = df[['2 jan t/m 6 jan','Unnamed: 2', 'Maandag','Dinsdag','Woensdag','Donderdag','Vrijdag']]
for col in df.columns:
    print(col)
df.rename(columns={'2 jan t/m 6 jan': 'Datum', 'Unnamed: 2': 'Persoon'}, inplace=True)
    


value = df.index[df['Datum']=='16 okt t/m 20 okt']
index_end = df.index[df['Datum']=='23 okt t/m 27 okt']
value= value.to_numpy()
index_end = index_end.to_numpy()
df = df.drop(df.loc[0:int(value)-3].index)
df = df.drop(df.loc[int(index_end):(df.tail(1).index.item())].index)
print(df.to_markdown())


