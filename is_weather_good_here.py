from dotenv import load_dotenv
import os
import requests
from pprint import pprint
from get_coordinates import get_coordinates
import time

load_dotenv()

def get_weather(lat, lng):
    api_key = os.getenv("OPEN_WEATHER_API_KEY")
    if not api_key:
        raise ValueError("API key not found in environment variables.")
    
    base_url = "http://api.openweathermap.org/data/2.5/weather"
    params = {
        "lat": lat, 
        "lon": lng,
        "appid": api_key,
        "units": "metric" 
    }

    response = requests.get(base_url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        response.raise_for_status()

if __name__ == "__main__":
    location = input("Enter Location: ")
    
    try:
        coordinates = get_coordinates(location)
        weather_response = get_weather(coordinates['lat'], coordinates['lng'])

        print(f"Fetching weather data for {coordinates['complete_address']}... \n")
        time.sleep(3)

        pprint(weather_response)
    except Exception as e:
        print(f"Error: {e}")