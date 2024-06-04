import requests
from bs4 import BeautifulSoup
import csv

# URL of the website to scrape
url = "https://www.thebmc.co.uk/climbing-wall-finder"

# Send an HTTP GET request to the website
response = requests.get(url)

# Parse the HTML content using Beautiful Soup
soup = BeautifulSoup(response.content, "html.parser")

# Find the relevant HTML elements that contain climbing center data
center_elements = soup.find_all("div", class_="center-info")

# Initialize a list to store the extracted data
data = []

# Extract climbing center names and postcodes
for center in center_elements:
    name = center.find("h2").text.strip()
    postcode = center.find("span", class_="postcode").text.strip()
    data.append([name, postcode])

# Write the data to a CSV file
csv_file = "climbing_centers.csv"
with open(csv_file, "w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(["Name", "Postcode"])  # Write header
    writer.writerows(data)                # Write data rows

print("Scraping and CSV creation complete.")
