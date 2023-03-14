from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title ='MOVIE LIST' 
sheet.append(['TRENDING','MOVIE NAME','RATING'])


try:

    responce = requests.get("https://www.imdb.com/chart/moviemeter/?ref_=nv_mv_mpm")
    text_data = BeautifulSoup(responce.text,'html.parser')
    #print (text_data)
    movies = text_data.find('tbody',class_="lister-list").find_all("tr")
    for movie in movies:
        #print (movie)
        movie_name = movie.find('td',class_="titleColumn").a.text
        rank = movie.find('div',class_="velocity")
        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        print(rank,movie_name,rating)
        sheet.append([rank,movie_name,rating])
        


except Exception as e:
    print(e)

excel.save('IMDB trending.xlsx')



