import requests
import pprint
import pandas as pd
from unidecode import unidecode


excel_data_df = pd.read_excel('ilce_merkez_coord.xlsx') #excel den dataframe e donusturme

for col in ['ILCEADI']:  
    excel_data_df[col] = excel_data_df[col].apply(unidecode)  #turkce karakterler decode edildi

print(excel_data_df)

ilce = excel_data_df.loc[0][5]

url = 'https://api.weatherstack.com/current?access_key=51d7bfab86b02431a6a6cc5d489ea60b&query='+ilce

res = requests.get(url)
data = res.json()

time = data['location']['localtime']


print('Time : ',time)
#print(data)


'''temp = data['main']['temp']
wind_speed = data['wind']['speed']

print('Temperature : ',temp)
print('Wind Speed : ',wind_speed)

'''
