#Importing Libraries
import pandas as pd
import matplotlib.pyplot as plt
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#Parsing Url
url = "https://disease.sh/v3/covid-19/countries"
response = requests.get(url)
fileName = "covidData.html"

if(response.status_code == 200):
    with open(fileName, "w") as file:
        file.write(response.text)   

    soup = BeautifulSoup(response.text, 'html.parser')
    print("Data fetched successfully and saved to", fileName)
else:
    print(f"Error {response.status_code}: Failed to retrieve data from the URL")

print(soup.prettify()[:1000])  # Print first 1000 characters for brevity