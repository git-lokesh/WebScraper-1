import pandas as pd
import requests
import gzip
import io

def download_imdb_datasets():
    """
    Download and process IMDb's official datasets
    """
    base_url = "https://datasets.imdbws.com/"
    
    files_to_download = {
        "title.basics.tsv.gz": "Basic movie information",
        "title.ratings.tsv.gz": "Movie ratings and votes"
    }
    
    datasets = {}
    
    for filename, description in files_to_download.items():
        print(f"Downloading {filename} - {description}")
        
        try:
            response = requests.get(base_url + filename)
            response.raise_for_status()
            
            with gzip.open(io.BytesIO(response.content), 'rt', encoding='utf-8') as f:
                df = pd.read_csv(f, sep='\t', na_values=['\\N'])
            
            datasets[filename.replace('.tsv.gz', '')] = df
            print(f"✅ Downloaded {len(df)} records")
            
        except Exception as e:
            print(f"Error downloading {filename}: {str(e)}")
    
    return datasets

def create_top_250_from_datasets(datasets):
    """
    Create Top 250 list from the official datasets
    """
    if 'title.basics' not in datasets or 'title.ratings' not in datasets:
        print("Required datasets not available")
        return pd.DataFrame()
    
    basics = datasets['title.basics']
    ratings = datasets['title.ratings']
    
    movies = basics[basics['titleType'] == 'movie'].copy()
    
    movie_ratings = movies.merge(ratings, on='tconst', how='inner')
    
    popular_movies = movie_ratings[movie_ratings['numVotes'] >= 25000]
    
    top_movies = popular_movies.sort_values('averageRating', ascending=False).head(250)
    
    result = pd.DataFrame({
        'Rank': range(1, len(top_movies) + 1),
        'MovieTitle': top_movies['primaryTitle'].values,
        'Year': top_movies['startYear'].values,
        'Rating': top_movies['averageRating'].values,
        'Votes': top_movies['numVotes'].values
    })
    
    return result

datasets = download_imdb_datasets()
top_250 = create_top_250_from_datasets(datasets)
top_250.to_csv("extracted_data.csv", index=False)
