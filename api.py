import requests
import pprint

city = input('Enter your city:')

url = 'http://api.openweathermap.org/data/2.5/weather?q={}&appid=1d0bdcc3ac9ff7582a02adab86f59c6b'.format(city)

res = requests.get(url)

data = res.json()

temp = data['main']['temp']
wind_speed = data['wind']['speed']

print('Temperature : ',temp)
print('Wind Speed : ',wind_speed)

