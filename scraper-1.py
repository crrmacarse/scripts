import requests
from pprint import pprint
from bs4 import BeautifulSoup

url = 'https://cpu.edu.ph/'
data = requests.get(url)

soup = BeautifulSoup(data.text, 'html.parser')

data = []
divItem = soup.find('div', { 'class': 'grid-items' })
for div in divItem.find_all('div', { 'class': 'title' }):
     data.append(div.text)

pprint(data)
