import json
import pandas as pd
import csv
import tzevaadom
from powerbiclient import Report
data = tzevaadom.alerts_history("en","month")
df = pd.DataFrame(columns=["city","rocket_date","rocket_time","alert_date_time","category","category_desc","matrix_id","rid"])
for i in range(0, len(data)):
    df.loc[i] = [data[i]["data"],data[i]["date"],data[i]["time"],data[i]["alertDate"],data[i]["category"],data[i]["category_desc"],data[i]["matrix_id"],data[i]["rid"]]
df['rocket_date'] = pd.to_datetime(df['rocket_date'], format='%d.%m.%Y')
df['rocket_time'] = pd.to_datetime(df['rocket_time'], format='%H:%M:%S').dt.time
df['alert_date_time'] = pd.to_datetime(df['alert_date_time'], format='%Y-%m-%dT%H:%M:%S')
pd.set_option('display.max_columns', None)  # Display all columns
pd.set_option('display.max_rows', None)
city_date_counts = df.groupby(['city', 'rocket_date', 'rocket_time', 'category_desc']).size().reset_index(name='count')
sorted_city_date_counts = city_date_counts.sort_values(by='count', ascending=False)
sorted_city_date_counts.to_excel("rockets_data.xlsx", sheet_name="data", index=False)






