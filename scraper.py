#Flipkart Product Data Scraper using Python

import requests
from bs4 import BeautifulSoup
import csv

# URL
#url = "https://www.flipkart.com/search?q=mobiles"
url = "https://www.geeksforgeeks.org/python/python-programming-language-tutorial/"

# Headers (to avoid blocking)
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

# Request
response = requests.get(url, headers=headers)

print("Status Code:", response.status_code)
print("Scraping Flipkart...")
print("Please wait...")
print(response)
# Parse HTML
soup = BeautifulSoup(response.text, "html.parser")

# Find the main content container
content = soup.find('div', class_='article--viewer_content')
if content:
    for para in content.find_all('p'):
        print(para.text.strip())
else:
    print("No article content found.")


# Find product containers
products = soup.find_all("div", {"class": "_1AtVbE"})

# Open CSV file
with open("flipkart_products.csv", "w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(["Product Name", "Price", "Rating", "Link"])

    for item in products:
        try:
            name = item.find("div", class_="_4rR01T").text
        except:
            name = "N/A"

        try:
            price = item.find("div", class_="_30jeq3 _1_WHN1").text
        except:
            price = "N/A"

        try:
            rating = item.find("div", class_="_3LWZlK").text
        except:
            rating = "N/A"

        try:
            link = item.find("a", class_="_1fQZEK")["href"]
            link = "https://www.flipkart.com" + link
        except:
            link = "N/A"

        if name != "N/A":
            writer.writerow([name, price, rating, link])

print("Scraping Completed! Data saved to flipkart_products.csv")