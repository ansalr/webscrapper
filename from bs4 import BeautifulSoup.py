from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Movie list"
sheet.append(['Rank','Name','Date','rating'])

try:
    response = requests.get("https://www.imdb.com/chart/top/")
    soup = BeautifulSoup(response.text,'html.parser')
    print(soup)
    movies = soup.find('tbody',class_="lister-list").find_all("tr")

    for movie in movies:
        movie_name = movie.find('td',class_="titleColumn").a.text
        movie_rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0
        ]
        movie_date = movie.find('td',class_="titleColumn").span.text
        movie_date = movie_date.replace('(',"")
        movie_date = movie_date.replace(')',"")
        movie_rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        #print (movie)
        #print (movie_rank,movie_name,movie_date,movie_rating)
        sheet.append([movie_rank,movie_name,movie_date,movie_rating])
          
except Exception as e:
    print(e)
excel.save('movies.xlsx')