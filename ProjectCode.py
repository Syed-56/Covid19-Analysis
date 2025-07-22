#Importing Libraries
import pandas as pd
import matplotlib.pyplot as plt
import requests
import json
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

def createChart(df, column, title, filename, color='tomato'):
    plt.figure(figsize=(10,6))
    plt.bar(df['country'], df[column], color=color)
    plt.title(title)
    plt.xlabel("Country")
    plt.ylabel(column.replace("_", " ").title())
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"visuals/{filename}.png")
    plt.close()

#Parsing Url
url = "https://disease.sh/v3/covid-19/countries"
response = requests.get(url)
data = response.json()
with open('Covid-Data.json', 'w') as f:
    json.dump(data, f, indent=4)

#Creating DataFrame and storin useful data only
df = pd.json_normalize(data)
df = df[['country', 'cases', 'deaths', 'recovered', 'active', 'tests', 'population']]
df.dropna(inplace=True)
print(df.head())

#Analyzing healthcare data
df['death_rate'] = (df['deaths'] / df['cases']).round(4) * 100
df['recovery-rate'] = (df['recovered'] / df['cases']).round(4) * 100
df['tests-per-case'] = (df['tests'] / df['cases']).round(2)
print(df[['country', 'death_rate', 'recovery-rate', 'tests-per-case']].head())

#Sorting DataFrame
mostAffected = df.sort_values(by='cases', ascending=False).head(10)
leastAffected = df.sort_values(by='cases', ascending=True).head(10)
highestDeathRate = df.sort_values(by='death_rate', ascending=False).head(10)
lowestDeathRate = df.sort_values(by='death_rate', ascending=True).head(10)
mostActiveCases = df.sort_values(by='active', ascending=False).head(10)

print("Top 3 Most Affected Countries:\n", mostAffected[['country', 'cases']].head(3).to_string(index=False))
print("Top 3 Least Affected Countries:\n", leastAffected[['country', 'cases']].head(3).to_string(index=False))
print("Top 3 Countries with Highest Death Rate:\n", highestDeathRate[['country', 'death_rate']].head(3).to_string(index=False))
print("Top 3 Countries with Lowest Death Rate:\n", lowestDeathRate[['country', 'death_rate']].head(3).to_string(index=False))
print("Top 3 Countries with Most Active Cases:\n", mostActiveCases[['country', 'active']].head(3).to_string(index=False))

df.to_excel("covid_report.xlsx", index=False)  #saving to Excel

#Creating Visuals
createChart(mostAffected, 'cases', 'Top 10 Most Affected Countries by COVID-19 Cases', 'most_affected_cases')
createChart(leastAffected, 'cases', 'Top 10 Least Affected Countries by COVID-19 Cases', 'least_affected_cases', color='lightblue')
createChart(highestDeathRate, 'death_rate', 'Top 10 Countries with Highest Death Rate', 'highest_death_rate', color='red')
createChart(lowestDeathRate, 'death_rate', 'Top 10 Countries with Lowest Death Rate', 'lowest_death_rate', color='green')
createChart(mostActiveCases, 'active', 'Top 10 Countries with Most Active Cases', 'most_active_cases', color='orange')