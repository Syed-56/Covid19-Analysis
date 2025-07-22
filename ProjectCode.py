#Importing Libraries
import pandas as pd
import matplotlib.pyplot as plt
import requests
import json
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#Parsing Url
url = "https://disease.sh/v3/covid-19/countries"
response = requests.get(url)
data = response.json()
with open('Covid-Data.json', 'w') as f:
    json.dump(data, f, indent=4)

#Creating DataFrame
# print(data)   Unformatted JSON data
df = pd.json_normalize(data)
print(df.head())