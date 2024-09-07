import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

def ScrapeIMDBTop250():
    try:
        imdbTop250URL = 'https://www.imdb.com/chart/top/'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(imdbTop250URL, headers=headers)
        
        if response.status_code != 200:
            raise Exception(f"Failed to retrieve page. Status code: {response.status_code}")
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # print(soup.prettify())
        
        movie_data_list = []
        for movieRow in soup.select('tbody.lister-list tr'):
            movieTitle = movieRow.find('td', class_='titleColumn').find('a').get_text(strip=True)
            releaseYear = movieRow.find('span', class_='secondaryInfo').get_text(strip=True)
            imdbRating = movieRow.find('td', class_='ratingColumn imdbRating').find('strong').get_text(strip=True)
            movie_data_list.append({'Title': movieTitle, 'Year': releaseYear, 'Rating': imdbRating})

        WorkbookObject = Workbook()
        active_sheet = WorkbookObject.active
        active_sheet.append(['Title', 'Year', 'Rating'])

        for movieData in movie_data_list:
            active_sheet.append([movieData['Title'], movieData['Year'], movieData['Rating']])
        
        WorkbookObject.save('imdb_top_250.xlsx')
        print("Data successfully written to 'imdb_top_250.xlsx'")
    
    except requests.exceptions.RequestException as networkError:
        print(f"Network error: {networkError}")
    
    except Exception as generalError:
        print(f"An error occurred: {generalError}")

ScrapeIMDBTop250()
