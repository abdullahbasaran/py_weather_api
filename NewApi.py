import requests
import pprint
import pandas as pd
from unidecode import unidecode
import schedule
import time
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from datetime import date, timedelta
import datetime
from openpyxl import load_workbook
import json

# excel den cekilen verinin dataframe e donusturme
excel_data_df = pd.read_excel('ilce_merkez_coord.xlsx')

# turkce karakterler decode edildi
for col in ['ILCEADI']:
    excel_data_df[col] = excel_data_df[col].apply(
        unidecode)  

print(excel_data_df) 

i = 1
# edate = date(2008, 5, 5) 

#1 Marttan bugune kadar tarih generate edilmesi 
dt = date(2020, 3, 1)
end = date.today()
step = datetime.timedelta(days=1)

#
def WriteToExcel(kod_ilce, sdate, temp, sunrise, sunset, moonrise, moonset, hour, wind_speed, temperature, wind_degree, wind_dir, weather_code, weather_descriptions, precip, humidity, visibility, pressure, cloudcover, heatindex, dewpoint, windchill, windgust, feelslike, chanceofrain, chanceofremdry, chanceofwindy, chanceofovercast, chanceofsunshine, chanceoffrost, chanceofhightemp, chanceoffog, chanceofsnow, chanceofthunder):

    hourly_label = ""
    hourInt = int(hour)

    if(hourInt == 0):
        hourly_label = "(00:00-03:00)"
    elif(hourInt == 300):
        hourly_label = "(03:00-06:00) "

    elif(hourInt == 600):
        hourly_label = "(06:00-09:00) "

    elif(hourInt == 900):
        hourly_label = "(09:00-12:00) "

    elif(hourInt == 1200):
        hourly_label = "(12:00-15:00) "

    elif(hourInt == 1500):
        hourly_label = "(15:00-18:00) "

    elif(hourInt == 1800):
        hourly_label = "(18:00-21:00) "

    else:
        hourly_label = "(21:00-00:00) "

    # excel tablosuNA her bir sutun icin degerlerin  girilmesi
    df = pd.DataFrame({'ILCEKOD': [kod_ilce], 'TIME': [sdate], 'TEMPRATURE': [temp], 'SUNRISE': [sunrise], 'SUNSET': [
                      sunset], 'MOONRISE': [moonrise], 'MOONSET': [moonset], 'HOURLY': [hourly_label], 'WINDSPEED': [wind_speed], 'TEMPERATURE': [temperature], 'WIND_DEGREE': [wind_degree], 'WIND_DIRECTION': [wind_dir], 'WEATHER_CODE': [weather_code], 'WEATHER_DESCRIPTIONS': [weather_descriptions], 'PRECIPITATION_LEVEL': [precip], 'HUMIDITY': [humidity], 'VISIBILITY': [visibility], 'PRESSURE': [pressure], 'CLOUD_COVER_LEVEL': [cloudcover], 'HEAT_INDEX': [heatindex], 'DEW_POINT': [dewpoint], 'WIND_CHILL': [windchill], 'WIND_GUST': [windgust], 'FEELS_LIKE': [feelslike], 'CHANCE_OF_RAIN': [chanceofrain], 'CHANCE_OF_DRY': [chanceofremdry], 'CHANCE_OF_WINDY': [chanceofwindy], 'CHANCE_OF_OVERCAST': [chanceofovercast], 'CHANCE_OF_SUNSHINE': [chanceofsunshine], 'CHANCE_OF_FROST': [chanceoffrost], 'CHANCE_OF_HIGH_TEMP': [chanceofhightemp], 'CHANCE_OF_FOG': [chanceoffog], 'CHANCE_OF_SNOW': [chanceofsnow], 'CHANCE_OF_THUNDER': [chanceofthunder]})
    #print(df)
    writer = pd.ExcelWriter('A_New_Excel.xlsx')
    df.to_excel(writer, 'Sheet1', index=False)

    # var olan bir workbook acilmasi
    writer.book = load_workbook('A_New_Excel.xlsx')

    # var olan dosyanin yazilmasi
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    # var olan dosyanin okunmasi

    reader = pd.read_excel(r'A_New_Excel.xlsx')
    # yeni sayfalarin yazdirilmasi
    df.to_excel(writer, index=False, header=False, startrow=len(reader)+1)
    writer.save()
    writer.close()


while dt < end:
    print(str(dt))
    print(str(end))
    sdate = str(dt.strftime('%Y-%m-%d'))
    # http://api.weatherstack.com/historical?access_key=51d7bfab86b02431a6a6cc5d489ea60b&query=Aladag&historical_date=2015-01-21&hourly=1
    ilce = excel_data_df.loc[0][5]
    kod_ilce = excel_data_df.loc[0][3]
    url = 'https://api.weatherstack.com/historical?access_key=51d7bfab86b02431a6a6cc5d489ea60b&hourly=1&historical_date='+sdate+'&query='+ilce
    res = requests.get(url)
    data = res.json()
    # time = data['location']['localtime'] #istenilen degerin cekilmesi
    temp = data['current']['temperature']  # istenilen degerin cekilmesi
    sunrise = data['historical'][sdate]['astro']['sunrise']
    sunset = data['historical'][sdate]['astro']['sunset']
    moonrise = data['historical'][sdate]['astro']['moonrise']
    moonset = data['historical'][sdate]['astro']['moonset']
    moon_illumination = data['historical'][sdate]['astro']['moon_illumination']

    for hourly_object in data['historical'][sdate]['hourly']:
        WriteToExcel(kod_ilce, sdate, temp, sunrise, sunset, moonrise, moonset, hourly_object['time'], hourly_object['wind_speed'], hourly_object['temperature'], hourly_object['wind_degree'], hourly_object['wind_dir'], hourly_object['weather_code'], hourly_object['weather_descriptions'], hourly_object['precip'], hourly_object['humidity'], hourly_object['visibility'], hourly_object['pressure'], hourly_object['cloudcover'], hourly_object[
                     'heatindex'], hourly_object['dewpoint'], hourly_object['windchill'], hourly_object['windgust'], hourly_object['feelslike'], hourly_object['chanceofrain'], hourly_object['chanceofremdry'], hourly_object['chanceofwindy'], hourly_object['chanceofovercast'], hourly_object['chanceofsunshine'], hourly_object['chanceoffrost'], hourly_object['chanceofhightemp'], hourly_object['chanceoffog'], hourly_object['chanceofsnow'], hourly_object['chanceofthunder'])
    dt += step
    i = i+1


# schedule.every(10).seconds.do(WriteToExcel) #excele schedule ile yazma
# while True:
#    schedule.run_pending()
# time.sleep(10)


# print(data)
