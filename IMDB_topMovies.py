from bs4 import BeautifulSoup
import requests
import openpyxl

WbObject = openpyxl.Workbook()
sheetObj = WbObject.active
sheetObj.title= 'IMDB Top Rated Movie'
# print(WbObject.sheetnames)
sheetObj.append(["Rank","Movie Name","Year of Release","IMDB Rating"])

try:
    req = requests.get('https://www.imdb.com/chart/top/')
    contents= req.content

    soup = BeautifulSoup(contents,'html.parser')
    # print(soup.prettify)
    tbody = soup.find('tbody', class_="lister-list").find_all('tr')
    # print(len(tbody))
    for movie in tbody:
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        name = movie.find('td',class_='titleColumn').a.text 
        year = movie.find('td',class_='titleColumn').span.text.strip('()')
        rating = movie.find('td', class_='ratingColumn imdbRating').get_text(strip=True)
        sheetObj.append([rank,name,year,rating])
        print(rank, name, year, rating)
        
except Exception as e:
    print(e)

WbObject.save("IMDB_Movie_Rating.xlsx")