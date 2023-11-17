import requests
import json
import pandas as pd
import os

API_KEY = 'c3ab9a2a4ccfaac2065c44ce3977630b'
OPENUV_API_KEY = 'openuv-g05efkrlov78pjf-io'  
GEOCODING_API_KEY = '4fa407584c56440da4a40124afb8cddb'  

def get_uv_index(city):
    base_url = 'https://api.opencagedata.com/geocode/v1/json'
    
    params = {'q': city, 'key': GEOCODING_API_KEY}
    response = requests.get(base_url, params=params)

    if response.status_code == 200:
        location_data = response.json()
        latitude = location_data['results'][0]['geometry']['lat']
        longitude = location_data['results'][0]['geometry']['lng']

        # Perform the UV index API call
        uv_params = {'lat': latitude, 'lng': longitude}
        uv_response = requests.get('https://api.openuv.io/api/v1/uv', headers={'x-access-token': OPENUV_API_KEY}, params=uv_params)

        if uv_response.status_code == 200:
            uv_data = uv_response.json()
            current_uv_index = uv_data.get('result', {}).get('uv', None)
            return current_uv_index
        else:
            print(f"UV Index API Error: {uv_response.status_code}")
            return None
    else:
        print(f"Geocoding API Error: {response.status_code}")
        return None

def get_5_day_forecast(city_name):
    base_url = 'http://api.openweathermap.org/data/2.5/forecast'
    
    params = {
        'q': city_name,
        'appid': API_KEY,
        'units': 'metric',
    }
    
    response = requests.get(base_url, params=params)
    
    if response.status_code == 200:
        data = response.json()
        return data
    else:
        print(f"Error: {response.status_code}")
        return None

def display_5_day_forecast(data):
    if data:
        city = data['city']['name']
        forecast_list = data['list']

        current_day = None
        day_counter = 0

        save_directory = r"D:\Hist_data\\"
        excel_filename = f"{save_directory}{city}_5_day_forecast.xlsx"

        with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
            for item in forecast_list:
                timestamp = item['dt_txt']
                date = timestamp.split()[0]

                if date != current_day:
                    if current_day is not None:
                        df = pd.DataFrame(forecast_data)
                        df.to_excel(writer, index=False, sheet_name=f'Day {current_day}')

                    current_day = date
                    day_counter += 1
                    print(f"Day {day_counter}:")

                    forecast_data = []

                day_data = {
                    'Date & Time': timestamp,
                    'Temperature (°C)': item['main']['temp'],
                    'Description': item['weather'][0]['description'],
                    'Cloudiness (%)': item['clouds']['all'],
                    'Humidity (%)': item['main']['humidity'],
                    'Wind Speed (m/s)': item['wind']['speed'],
                    'Rain (3h) (mm)': item.get('rain', {}).get('3h', 0),
                    'UV Index': get_uv_index(city)  # Use the function to get UV Index
                }
                forecast_data.append(day_data)

                print(f"Date & Time: {timestamp}")
                print(f"Temperature: {day_data['Temperature (°C)']}°C")
                print(f"Description: {day_data['Description']}")
                print(f"Cloudiness: {day_data['Cloudiness (%)']}%")
                print(f"Humidity: {day_data['Humidity (%)']}%")
                print(f"Wind Speed: {day_data['Wind Speed (m/s)']} m/s")
                print(f"Rain (3h): {day_data['Rain (3h) (mm)']} mm")
                print(f"UV Index: {day_data['UV Index']}\n")

            df = pd.DataFrame(forecast_data)
            df.to_excel(writer, index=False, sheet_name=f'Day {current_day}')

        save_to_excel = input("Do you want to save the 5-Day Weather Forecast to Excel? (yes/no): ").lower()

        if save_to_excel == 'yes':
            print(f"5-Day Weather Forecast saved to {excel_filename}")
        else:
            os.remove(excel_filename)
            print("Forecast data not saved.")
    else:
        print("No data found :(")

if __name__ == '__main__':
    print('Ulapp Forecasting(BETA)')
    city_name = input("Enter city name: ")
    forecast_data = get_5_day_forecast(city_name)
    display_5_day_forecast(forecast_data)
