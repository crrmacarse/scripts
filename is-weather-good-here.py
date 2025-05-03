from dotenv import load_dotenv
import os
import googlemaps
import argparse
import requests
from pprint import pprint

load_dotenv()

def get_coordinates(location):
    api_key = os.getenv("GOOGLE_MAPS_API_KEY")
    if not api_key:
        raise ValueError("API key not found in environment variables.")
    
    gmaps = googlemaps.Client(key=api_key)
    geocode_result = gmaps.geocode(location)
    
    if geocode_result:
        formatted_address = geocode_result[0]['formatted_address']
        coordinates = geocode_result[0]['geometry']['location']

        return {
            "complete_address": formatted_address,
            "lat": coordinates['lat'],
            "lng": coordinates['lng']
        }
    else:
        raise ValueError("No results found for the given location.")
    
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
    # parse passed params
    parser = argparse.ArgumentParser(description="Is weather good here?")
    parser.add_argument("--location", required=True, help="Location to get weather data for")

    args = parser.parse_args()

    # load passed params to variables
    location = args.location
    
    try:
        coordinates = get_coordinates(location)
        weather_response = get_weather(coordinates['lat'], coordinates['lng'])

        print(f"Weather forecast for {coordinates['complete_address']}")
        pprint(weather_response)
    except Exception as e:
        print(f"Error fetching coordinates: {e}")