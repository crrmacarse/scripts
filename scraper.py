import requests
from pprint import pprint
from bs4 import BeautifulSoup
from utils.validators import is_valid_url  # Import the reusable validator
import argparse

def scrape(passed_url):
     url = is_valid_url(passed_url)

     response = requests.get(url)
     soup = BeautifulSoup(response.text, 'lxml')

     return soup.prettify()

if __name__ == "__main__":
     # parse passed params
     parser = argparse.ArgumentParser(description="Web Scraper")
     parser.add_argument("--url", required=True, help="Website URL to scrape")

     args = parser.parse_args()

     # load passed params to variables
     url = args.url

     try:
          scraped = scrape(url)

          pprint(scraped)
     except Exception as e:
          print(f"Error: {e}")