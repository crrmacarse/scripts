from dotenv import load_dotenv
import os
import googlemaps
import argparse

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
        return {"complete address": formatted_address, "latitude": coordinates['lat'], "longitude": coordinates['lng']}
    else:
        raise ValueError("No results found for the given location.")

if __name__ == "__main__":
    # parse passed params
    parser = argparse.ArgumentParser(description="Get Coordinats")
    parser.add_argument("--location", required=True, help="Location to get coordinates for")

    args = parser.parse_args()

    # load passed params to variables
    location = args.location
    
    try:
        coordinates = get_coordinates(location)
        print(f"Coordinates for {location}: {coordinates}")
    except Exception as e:
        print(f"Error fetching coordinates: {e}")