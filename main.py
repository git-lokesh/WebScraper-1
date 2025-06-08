import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

genres = ["action", "drama", "romance", "sci-fi", "comedy", "animation"]
base_url = "https://www.imdb.com/chart/top/"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
}

def scrape_genre(genre):
    print(f"Scraping genre: {genre}")
    response = requests.get(base_url.format(genre), headers=headers)
    soup = BeautifulSoup(response.content, "html.parser")
    containers = soup.find_all("div", class_="lister-item mode-advanced")
    print(f"Found {len(containers)} movie entries for {genre}.")

    rows = []
    for container in containers:
        title_tag = container.h3.a
        rating_tag = container.find("div", class_="ratings-bar")
        vote_tags = container.find_all("span", attrs={"name": "nv"})

        title = title_tag.text.strip() if title_tag else None
        rating = container.strong.text.strip() if container.strong else None
        votes = vote_tags[0]['data-value'] if vote_tags else None
        gross = vote_tags[1].text if len(vote_tags) > 1 else None

        rows.append({
            "MovieTitle": title,
            "Genre": genre.capitalize(),
            "Rating": rating,
            "Votercount": votes,
            "Gross": gross if gross else "N/A"
        })

    df = pd.DataFrame(rows)
    print(f"Completed genre: {genre}, rows collected: {len(df)}\n")
    return df

print("Starting IMDb scraping by genre...")
all_data = []

for genre in genres:
    genre_df = scrape_genre(genre)
    all_data.append(genre_df)
    time.sleep(1)  

final_df = pd.concat(all_data, ignore_index=True)
print("Scraping completed. Saving to CSV...")
final_df.to_csv("IMDb_Genre_Top50.csv", index=False)
print("Data saved to IMDb_Genre_Top50.csv")
