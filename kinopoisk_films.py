import os
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment
from time import sleep

#Создание файла excel
cwd = os.getcwd()

workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Top'
filename = 'kinopoisk_films.xlsx'
worksheet.append(['Название', 'Рейтинг', 'Год', 'Страна', 'Жанр', 'Режиссер', 'Бюджет', 'Длительность'])

#выравнивание
alignment = Alignment(horizontal='center', vertical='center')
#применяем ко всему
for row in worksheet.rows:
    for cell in row:
        cell.alignment = alignment

#ширина колонок
for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
    worksheet.column_dimensions[col].width = 20


workbook.save(os.path.join(cwd, filename))

headers = { 'User-Agent':
                'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/69.0'}

film_blocks = []

for page_num in range(1, 31):
    sleep(3)
    url = f'https://www.kinopoisk.ru/lists/movies/country--1/?page=2{page_num}'
    response = requests.get(url, headers=headers)

    soup = BeautifulSoup(response.text, 'html.parser')
    movie_blocks = soup.find_all('div', class_ = 'styles_root__ti07r')

    for i in movie_blocks:
        film_url = 'https://www.kinopoisk.ru' + i.find('a').get('href')
        film_blocks.append(film_url)


for urls in film_blocks:
    response = requests.get(urls, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    name = soup.find('h1', class_ = 'styles_title__65Zwx styles_root__l9kHe styles_root__5sqsd styles_rootInLight__juoEZ').text
    rating = soup.find('span', class_ = 'styles_ratingKpTop__84afd').text
    year = soup.find('a', class_ = 'styles_linkDark__7m929 styles_link__3QfAk').text
    country = soup.find('a', class_ = 'styles_linkDark__7m929 styles_link__3QfAk').text
    genre = soup.find_all('a', class_ = 'styles_linkDark__7m929 styles_link__3QfAk').text
    director = soup.find('a', class_ = 'styles_linkDark__7m929 styles_link__3QfAk').text
    budget = soup.find('a', class_ = 'styles_linkDark__7m929 styles_link__3QfAk').text
    duration = soup.find('div', class_ = 'styles_valueDark__BCk93 styles_value__g6yP4').text
    worksheet.append([name, rating, year, country, genre, director, budget, duration])


