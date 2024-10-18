import numpy
import pandas as pd
import requests
import numpy as np
import matplotlib.pyplot as plt

# importing user data
harp_df = pd.read_csv('harperalmen-films.csv')
nate_df = pd.read_csv('smokingn8-films.csv')
robb_df = pd.read_csv('freshrob87-films.csv')
json_df = pd.read_csv('abildjason-films.csv')
nana_df = pd.read_csv('leana666-films.csv')
rach_df = pd.read_csv('almenjoy98-films.csv')
shua_df = pd.read_csv('shua_too-films.csv')

# adding user to each user's dataframe
harp_df['User'] = 'harperalmen'
nate_df['User'] = 'smokingn8'
robb_df['User'] = 'freshrob87'
json_df['User'] = 'abildjason'
nana_df['User'] = 'leana666'
rach_df['User'] = 'almenjoy98'
shua_df['User'] = 'shua_too'

#combining everyone into one dataframe
fren_df = pd.concat([harp_df, nate_df,robb_df,json_df,nana_df,rach_df,shua_df], ignore_index=True)

# Headers
# Film_title,Release_year,Director,Cast,Average_rating,Owner_rating,Genres
# List of unique movies from overall list
# set probably not right move because I need the corresponding years (remakes with same titles)
all_movies_df = fren_df[['Film_title', 'Release_year']]
movies_df = pd.DataFrame.drop_duplicates(all_movies_df, ignore_index=True)
print(movies_df)

#Get IMDb movie IDs from omdb, create DF with them
def get_imdb_id(Film_title, year=None):
    api_key = 'cc4972ca'  # Just the API key here
    url = f"http://www.omdbapi.com/?apikey={api_key}&t={Film_title}"
    if year:
        url += f"&y={year}"
    response = requests.get(url)
    data = response.json()
    return data.get('imdbID')

movie_ids = get_imdb_id(movies_df, year='Release_year')
print(movie_ids)

#movies_df['MovieID'] = movies_df.apply(lambda row: get_imdb_id(row['Film_title'], row['Year']), axis=1)
#print(movies_df)
