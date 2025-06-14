# Install: pip install selenium webdriver-manager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

def scrape_imdb_with_selenium():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    
    try:
        driver.get("https://www.imdb.com/chart/top/")
        time.sleep(3) 
        movies_data = []
        for i in range(1, 251):
            try:
                title = driver.find_element(By.XPATH, f"//h3[contains(@class, 'title')][{i}]").text
                year = driver.find_element(By.XPATH, f"//span[contains(@class, 'year')][{i}]").text
                rating = driver.find_element(By.XPATH, f"//span[contains(@class, 'rating')][{i}]").text
                votes = driver.find_element(By.XPATH, f"//span[contains(@class, 'votes')][{i}]").text
                
                movies_data.append({
                    "Rank": i,
                    "MovieTitle": title,
                    "Year": year,
                    "Rating": rating,
                    "Votes": votes
                })
            except:
                continue
        
        return pd.DataFrame(movies_data)
    
    finally:
        driver.quit()

df = scrape_imdb_with_selenium()
df.to_csv("Selenium_Scraped_Data.csv", index=False)
