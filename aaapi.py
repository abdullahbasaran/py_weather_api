import requests
import pprint
import pandas as pd
from unidecode import unidecode
import schedule
import time
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np


excel_data_df = pd.read_excel('ilce_merkez_coord.xlsx') #excel den dataframe e donusturme

for col in ['ILCEADI']:  
    excel_data_df[col] = excel_data_df[col].apply(unidecode)  #turkce karakterler decode edildi

print(excel_data_df)

ilce = excel_data_df.loc[0][5]

url = 'https://api.weatherstack.com/historical?access_key=51d7bfab86b02431a6a6cc5d489ea60b&query='+ilce

res = requests.get(url)
data = res.json()

# time = data['location']['localtime'] #istenilen degerin cekilmesi
temp = data['current']['temperature']
sunrise = data['historical'][sdate]['astro']['sunrise']
sunset = data['historical'][sdate]['astro']['sunset']
moonrise = data['historical'][sdate]['astro']['moonrise']
moonset = data['historical'][sdate]['astro']['moonset']
moon_illumination = data['historical'][sdate]['astro']['moon_illumination']

print('Time : ',time)
print('Temprature : ',temp)

def ExcelWriter(self):
    df = pd.DataFrame({'ILCEKOD': [kod_ilce], 'TIME': [sdate], 'TEMPRATURE': [temp], 'SUNRISE': [sunrise], 'SUNSET': [
                          sunset], 'MOONRISE': [moonrise], 'MOONSET': [moonset]})  # excel tablosunda her bir sutun icin degerlerin  girilmesi
    writer = ExcelWriter('New_Excel.xlsx')
    df.to_excel(writer,'Sheet1',index=False)
    writer.save()

'''
schedule.every(10).seconds.do(ExcelWriter) #excele schedule ile yazma
while True:
    schedule.run_pending()
    #time.sleep(10)
'''

#print(data)

