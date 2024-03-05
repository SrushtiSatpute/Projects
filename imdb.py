from bs4 import BeautifulSoup
import requests
import openpyxl

# Create a new Excel workbook
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top movies of 2023'

# Add headers to the sheet
sheet.append(['Movie Rank', 'Movie Name', 'Movie Genre', 'IMDB Rating', 'Year Of Release', 'Duration', 'Metascore', 'Votes'])

try:
    source = requests.get('https://www.imdb.com/list/ls576754431/')
    source.raise_for_status() # Corrected to call the function

    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find_all('div', class_='lister-item-content')

    for movie in movies:
        # Extract relevant data
        name = movie.find('h3', class_='lister-item-header').a.text
        rank = movie.find('h3', class_='lister-item-header').find('span', class_='lister-item-index unbold text-primary').text.split('.')[0]
        year_text = movie.find('h3', class_='lister-item-header').find('span', class_='lister-item-year text-muted unbold').text
        year = year_text.split('(')[-1].split(')')[0] # More robust extraction of the year
        # Corrected the class names for the rating
        rating = movie.find('div', class_='ipl-rating-star small').find('span', class_='ipl-rating-star__rating').text 
        genre = movie.find('p', class_='text-muted text-small').find('span', class_='genre').text.strip() 
        # Added strip() to remove leading/trailing whitespace
        duration_elem = movie.find('p', class_='text-muted text-small').find('span', class_='runtime')
        duration = duration_elem.text if duration_elem else 'N/A'
        metascore_elem = movie.find('span', class_='metascore')
        metascore = metascore_elem.text.strip() if metascore_elem else 'N/A'
        votes_elem = movie.find('span', attrs={'name': 'nv'})
        votes = votes_elem['data-value'] if votes_elem else 'N/A'

        # Append data to the sheet
        sheet.append([rank, name, genre, rating, year, duration, metascore, votes])

    # Save the Excel file

    excel.save('top_movies_2023.xlsx')
    print("Data successfully saved to 'top_movies_2023.xlsx'")

except requests.RequestException as e:
    print(f"Error fetching data: {e}")