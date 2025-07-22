#Importing Libraries
import pandas as pd
import matplotlib.pyplot as plt
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#Parsing Url
url = "https://disease.sh/v3/covid-19/countries"
response = requests.get(url)
fileName = "covidData.html"

if(response.status_code == 200):
    with open(fileName, "w") as file:
        file.write(response.text)   
else:
    print(f"Error {response.status_code}: Failed to retrieve data from the URL")