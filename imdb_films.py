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
filename = 'imdb_films.xlsx'
worksheet.append(['Название', 'Рейтинг', 'Бюджет', 'Жанр', 'Длительность', 'Страна', 'Режиссер', 'Год', 'Сборы'])

#выравнивание
alignment = Alignment(horizontal='center', vertical='center')
#применяем ко всему
for row in worksheet.rows:
    for cell in row:
        cell.alignment = alignment

#ширина колонок
for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
    worksheet.column_dimensions[col].width = 20


workbook.save(os.path.join(cwd, filename))


url_base = "https://www.imdb.com/search/title/?count=100&groups=top_1000&sort=user_rating"

for i in range(1, 1001, 100):
    url = url_base + f"&start={i}"
    response = requests.get(url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.content, 'html.parser')
    movies = soup.find_all('div', class_='lister-item-content')
    for movie in movies:
        title = movie.find('a').text.strip()
        year = movie.find('span', class_='lister-item-year').text.strip('()')
        rating = movie.find('div', class_='ratings-imdb-rating').text.strip()

        # Проверка наличия режиссера и извлечение данных
        director_element = movie.find('p', class_='').find_all('a')
        director = director_element[0].text.strip() if director_element else ""

        # Проверка наличия жанра и извлечение данных
        genre_element = movie.find('a', class_='').find_all('span', class_='genre')
        genre = genre_element[0].text.strip() if genre_element else ""

        # Проверка наличия страны происхождения и извлечение данных
        country_element = movie.find('p', class_='').find_all('span', class_='ghost')
        country = country_element[0].text.strip() if country_element else ""

        # Проверка наличия бюджета и извлечение данных
        budget_element = movie.find('p', class_='').find_all('span', class_='ghost')
        budget = budget_element[1].text.strip() if len(budget_element) >= 2 else ""

        # Проверка наличия общего сбора и извлечение данных
        gross_element = movie.find('p', class_='').find_all('span', class_='ghost')
        gross = gross_element[2].text.strip() if len(gross_element) >= 3 else ""

        runtime_element = movie.find('p', class_='').find_all('span', class_='runtime')
        runtime = runtime_element[0].text.strip() if runtime_element else ""

        worksheet.append([title, rating, budget, genre, runtime, country,  director, year,  gross])


workbook.save(filename)


