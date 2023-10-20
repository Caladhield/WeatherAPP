import pandas as pd
import pytz
import os
import requests
from datetime import datetime, timedelta

longitude = 18.02151508449004
latitude = 59.30996552541549

smhi_url = f'https://opendata-download-metfcst.smhi.se/api/category/pmp3g/version/2/geotype/point/lon/{longitude:.6f}/lat/{latitude:.6f}/data.json'

openweathermap_api_key = 'fa49fd29427dc75eb3b4febb8fa2c1b1'
openweathermap_url = f'https://api.openweathermap.org/data/3.0/onecall?lat=59.30&lon=18.02&exclude=minutely&appid={openweathermap_api_key}'

def get_smhi_data():
    response = requests.get(smhi_url)
    if response.status_code == 200:
        data = response.json()

        weather_data = []

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        for time_series in data['timeSeries']:
            timestamp = time_series['validTime'].replace('T', ' ').replace('Z', '')
            time = datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
            local_time = (time - datetime.now()).total_seconds() / 3600

            if 0 <= local_time <= 24:
                date = timestamp.split()[0]
                hour = int(timestamp.split()[1][:2])

                temperature = 0
                precipitation = 0
                provider = 'SMHI'

                for items in time_series['parameters']:
                    if items['name'] == 't':
                        temperature = items['values'][0]
                    elif items['name'] == 'pcat':
                        precipitation = items['values'][0]

                precipitation = precipitation > 0

                weather_data.append({
                    'Created': current_time,
                    'Longitude': longitude,
                    'Latitude': latitude,
                    'Datum': date,
                    'Hour': hour,
                    'Temperature (°C)': temperature,
                    'Precipitation': precipitation,
                    'Provider': provider,
                })

        update_weather_data(weather_data)

    else:
        print(f'SMHI Data Error: {response.status_code}')

def get_openweathermap_data():
    response = requests.get(openweathermap_url)
    if response.status_code == 200:
        data = response.json()
        weather_data = []

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        now_sweden = datetime.now(pytz.timezone('Europe/Stockholm'))
        next_24_hours = now_sweden + timedelta(hours=24)

        for item in data['hourly']:
            time = item['dt']

            # Convert the OWM data timestamp to the Swedish time zone
            timestamp_utc = datetime.utcfromtimestamp(time)
            timestamp_sweden = timestamp_utc.astimezone(pytz.timezone('Europe/Stockholm'))

            if now_sweden <= timestamp_sweden < next_24_hours:
                temp_kelvin = item['temp']
                temp_celsius = temp_kelvin - 273.15
                weather = item['weather'][0]['main']

                date = timestamp_sweden.strftime('%Y-%m-%d')
                hour = timestamp_sweden.strftime('%H')

                precipitation = weather == 'Rain' or weather == 'Snow'

                weather_data.append({
                    'Created': current_time,
                    'Longitude': longitude,
                    'Latitude': latitude,
                    'Datum': date,
                    'Hour': hour,
                    'Temperature (°C)': temp_celsius,
                    'Precipitation': precipitation,
                    'Provider': 'OWM',
                })

        update_weather_data(weather_data)

    else:
        print(f'OpenWeatherMap Data Error: {response.status_code}')

def update_weather_data(new_data):
    file_path = "Combined_Weather_Data.xlsx"
    if os.path.exists(file_path):
        existing_data = pd.read_excel(file_path)
        combined_data = pd.concat([existing_data, pd.DataFrame(new_data)])
        combined_data = combined_data.drop_duplicates(subset=['Created', 'Datum', 'Hour', 'Provider'], keep='last')
    else:
        combined_data = pd.DataFrame(new_data)

    combined_data.to_excel(file_path, index=False, engine='openpyxl')
    print(f'Data saved to {file_path}')

def print_latest_forecast(filename, provider):
    try:
        dataframe = pd.read_excel(filename)

        # Convert 'Datum' to datetime
        dataframe['Datum'] = pd.to_datetime(dataframe['Datum'])

        # Filter by provider and sort by 'Datum' and 'Hour' in ascending order
        latest_data = dataframe[dataframe['Provider'] == provider].sort_values(by=['Datum', 'Hour'])

        # Create a set to store unique (date, hour) pairs
        unique_hours = set()
        formatted_date = None

        for _, row in latest_data.iterrows():
            date = row['Datum']
            hour = row['Hour']
            pair = (date, hour)

            if pair not in unique_hours:
                unique_hours.add(pair)
                temperature = row['Temperature (°C)']
                precipitation = row['Precipitation']
                if precipitation:
                    precipitation_text = "Nederbörd"
                else:
                    precipitation_text = "Ingen ederbörd"
                if date != formatted_date:
                    # Print a new date if the data in 'Datum' column changes
                    formatted_date = date
                    print(f"Prognos från {provider} {formatted_date.strftime('%Y-%m-%d')}:")
                # Print the forecast for each hour in the next 24 hours
                print(f"{hour:02d}:00 {int(temperature)} grader {precipitation_text}.")

                # Sets the max to 24 unique data points
                if len(unique_hours) >= 24:
                    break

    except FileNotFoundError:
        print(f'File {filename} does not exist')

while True:
    print("Menu:\n")
    print("1. Get SMHI Data")
    print("2. Get OpenWeatherMap Data")
    print("3. Print SMHI Prognosis")
    print("4. Print OpenWeatherMap Prognosis")
    print("5. Exit")

    choice = input("Select an option: ")
    if choice == '1':
        get_smhi_data()
    elif choice == '2':
        get_openweathermap_data()
    elif choice == '3':
        print_latest_forecast("Combined_Weather_Data.xlsx", "SMHI")
    elif choice == '4':
        print_latest_forecast("Combined_Weather_Data.xlsx", "OWM")
    elif choice == '5':
        print("Exiting the program.")
        break
    else:
        print("Please choose a valid option (1, 2, 3, 4, or 5).")
