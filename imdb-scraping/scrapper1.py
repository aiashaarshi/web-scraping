
import requests 
from bs4 import BeautifulSoup
from openpyxl import Workbook
import re

wb = Workbook()

# grab the active worksheet
ws = wb.active
ws.title = 'Top 2022 Movies'
ws.append(['Ranking', 'Name', 'Release Year', 'Rating'])


response = requests.get("https://www.imdb.com/list/ls099693310/")
# print(response.text)

soup = BeautifulSoup(response.text, 'html.parser')

# print(soup.a)

movies = soup.find_all("div", {"class": "lister-item-content"})

for movie in movies:
    # print(movie) 
     
    ranking = movie.h3.find("span", {"class": "lister-item-index unbold text-primary"}).text
    ranking = ranking.replace('.', '') 
    name = movie.h3.a.text
    
    year = movie.h3.find("span", {"class": "lister-item-year text-muted unbold"}).text.strip()
    # year = year.strip('()I ')
    year = re.sub('[^0-9]','', year)
    if year == '':
        year = 'N/A'

    # year = year.replace('(', '').replace(')', '') 
    
    ratingDiv = movie.find('div', {"class": "ipl-rating-star small"})
    rating = 'N/A'
    if ratingDiv is not None:
        rating = ratingDiv.find('span', {'class': 'ipl-rating-star__rating'}).text
    ws.append([ranking, name, year, rating])
    print([ranking, name, year, rating])


wb.save("sample.xlsx")
