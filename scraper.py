import requests
from pprint import pprint
from bs4 import BeautifulSoup
from utils.validators import is_valid_url  # Import the reusable validator

def scrape(passed_url):
     url = is_valid_url(passed_url)

     response = requests.get(url)
     soup = BeautifulSoup(response.text, 'lxml')

     return soup.prettify()

if __name__ == "__main__":
     # parse passed params
     url = input("Input URL: ")

     try:
          scraped = scrape(url)

          pprint(scraped)
     except Exception as e:
          print(f"Error: {e}")