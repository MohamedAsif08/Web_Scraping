import openpyxl
import requests
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Movies List'
sheet.append(['Rank', 'Name', 'Year', 'Duration', 'Grade', 'Rating', 'Reviewers'])

try:
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0.'}
    response = requests.get("https://m.imdb.com/chart/top/", headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    movies = soup.find('ul',
                       class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-9d2f6de0-0 iMNUXk "
                              "compact-list-view ipc-metadata-list--base")

    for movie in movies:
        # Movies List
        movie_list = movie.find('h3').text.split(". ")
        movie_rank = movie_list[0]
        movie_name = movie_list[1]

        # Other Info
        other_info = movie.find('div', class_="sc-479faa3c-7 jXgjdT cli-title-metadata")
        other_info_list = []
        for i in other_info:
            other_info_list.append(i.text)
        year = other_info_list[0]
        Duration = other_info_list[1]
        if len(other_info_list) > 2:
            grade = other_info_list[2]
        else:
            grade = 'Null'

        # Rating List
        rating_info_list = movie.find('span',
                                      class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb '
                                             'ratingGroup--imdb-rating')
        rating_list = []
        for i in rating_info_list:
            rating_list.append(i.text)
        rating = rating_list[1]
        reviewers = rating_list[2].replace("\xa0(", "").replace(")", "")
        Final_list = [movie_rank, movie_name, year, Duration, grade, rating, reviewers]
        sheet.append(Final_list)

except Exception as Error:
    print(Error)

excel.save('Movies.xlsx')
print('Movie List Created')