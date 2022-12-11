from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'   # add the sheet name
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])    # add the first row (added 4 columns)


try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()                                  # to handle webpage error

    soup = BeautifulSoup(source.text, 'html.parser')

    movies = soup.find('tbody', class_="lister-list").find_all('tr')        

    # first of all, we will use find by searching its attribute and class then we will use find_all to get the list of elements
    # then by using for loop we can fetch data accordingly

    for movie in movies:
        name = movie.find('td',class_= "titleColumn").a.text
        rank = movie.find('td',class_="titleColumn").get_text(strip= True).split('.')[0]       #get_text(strip=True)  is used to remove extra line and spaces
        year = movie.find('td',class_="titleColumn").span.text.strip('()')                     # strip('()')    is used to remove special character braces '()'
        rating =  movie.find('td',class_= "ratingColumn imdbRating").strong.text
        print(rank,name,year, rating)
        sheet.append([rank,name,year, rating])

except Exception as e:
    print(e)
excel.save('IMDB Movie Ratings.xlsx')

