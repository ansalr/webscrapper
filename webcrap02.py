from bs4 import BeautifulSoup
import openpyxl,requests

excel = openpyxl.Workbook()
sheet = excel.active
sheet.append(["rank","name","year","rating"])

try:
    web_req = requests.get("https://www.imdb.com/chart/top/")
    web_data = BeautifulSoup(web_req.text,"html.parser")
    top_120 = web_data.find("tbody",class_="lister-list").find_all("tr")
    for data in top_120 :
       movie_name = data.find("td",class_="titleColumn").a.text
       movie_rating = data.find("td",class_="ratingColumn imdbRating").strong.text
       movie_year = data.find("td",class_="titleColumn").span.text
       movie_year= movie_year.replace("(","")
       movie_year = movie_year.replace(")","")
       movie_rank = data.find("td",class_="titleColumn").get_text(strip=True).split(".")[0]
       print(movie_rank,movie_name,movie_year,movie_rating)
       sheet.append([movie_rank,movie_name,movie_year,movie_rating])
       
except Exception as e:
    print(e)

excel.save("movies.xlsx")