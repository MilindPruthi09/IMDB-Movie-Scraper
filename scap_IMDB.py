from bs4 import BeautifulSoup
import openpyxl
import requests, openpyxl
import re

excel=openpyxl.Workbook()

sheet=excel.active
sheet.title='Top Rated Movies'

sheet.append(['Movie Rank', 'Movie Name', 'Year Of Release' , 'Duration', 'Rating'])

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
}

try:
    source=requests.get('https://www.imdb.com/chart/top/', headers=headers)
    source.raise_for_status()

    soup=BeautifulSoup(source.text, 'html.parser')
    
    movies=soup.find('ul', class_='ipc-metadata-list ipc-metadata-list--dividers-between sc-a1e81754-0 dHaCOW compact-list-view ipc-metadata-list--base').find_all('li')

    # Iterating in movies variable to extract the data through each iteration
    for movie in movies:

        movie_name=movie.find('div', class_='ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-b189961a-9 bnSrml cli-title').text

        # Splitting movie id and movie name
        movie_id=movie_name.split(".")[0]
        movie_name=movie_name.split(".")[1]

        # getting the other movie details like year of release, duration.
        movie_details=movie.find('div', class_='sc-b189961a-7 btCcOY cli-title-metadata').text
        movie_year=movie_details[:4]
        movie_details_rest=movie_details[4:]

        # Using regex extracting the duration of movie.
        pattern = r"(.*?m)(.*?)"
        matches = re.findall(pattern, movie_details_rest)[0][0]

        movie_time=matches
        movie_rating=movie.find('span', 'ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text
        movie_rating=movie_rating[:3]

        # Appending the data from each iteration to sheet.
        sheet.append([movie_id, movie_name, movie_year, movie_time, movie_rating])

except Exception as e:
    print(e)

# Exporting the scrapped movie data to excel.
excel.save('IMDB Movie Ratings.xlsx')