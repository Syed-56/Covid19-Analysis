import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import requests
import json
from openpyxl.drawing.image import Image
from tabulate import tabulate
from openpyxl import load_workbook


def createChart(df, column, title, filename, color='skyblue'):
    plt.figure(figsize=(10, 8))
    plt.barh(df['country'], df[column], color=color)

    plt.title(title, fontsize=16, fontweight='bold', loc='center',)
    plt.xlabel(column.replace("_", " ").title(), fontsize=13, fontweight='bold', labelpad=15)
    plt.ylabel("Country", fontsize=13, fontweight='bold', labelpad=15)

    plt.gca().xaxis.set_major_formatter(mtick.StrMethodFormatter('{x:,.0f}'))
    plt.xticks(fontsize=11)
    plt.yticks(fontsize=11)

    plt.subplots_adjust(left=0.1, right=0.95, top=0.88, bottom=0.15)
    plt.subplots_adjust(right=1.20)
    plt.savefig(f"visuals/{filename}.png", bbox_inches='tight') 
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
excluded_countries = ['Aruba', 'Saint Helena', 'Tokelau', 'Western Sahara', 'Diamond Princess', 'MS Zaandam', 'Wallis and Futuna','British Virgin Islands','Cabo Verde','Caribbean Netherlands','Cayman Islands','Channel Islands','Cook Islands','Curaçao','Falkland Islands (Malvinas)','Faroe Islands','Gibraltar','Guadeloupe','Isle of Man','Libyan Arab Jamahiriya','Macao','Martinique','Montserrat','New Caledonia','Niue','Réunion','Saint Martin','Saint Pierre Miquelon','Sint Maarten','St. Barth']
df = df[~df['country'].isin(excluded_countries)]
df.dropna(inplace=True)
print(tabulate(df.head(),headers='keys', tablefmt='pretty'))

#Analyzing healthcare data
df['death_rate'] = ((df['deaths'] / df['cases'])*100).round(3)
df['recovery-rate'] = ((df['recovered'] / df['cases'])*100).round(3)
df['tests-per-case'] = ((df['tests'] / df['cases'])*100).round(3)
print(tabulate(df[['country', 'death_rate', 'recovery-rate', 'tests-per-case']].head(), headers='keys', tablefmt='pretty'))

#Sorting DataFrame
mostAffected = df.sort_values(by='cases', ascending=False).head(10)
leastAffected = df.sort_values(by='cases', ascending=True).head(10)
highestDeathRate = df.sort_values(by='death_rate', ascending=False).head(10)
lowestDeathRate = df.sort_values(by='death_rate', ascending=True).head(10)
mostActiveCases = df.sort_values(by='active', ascending=False).head(10)

print("Top 3 Most Affected Countries:")
print(tabulate(mostAffected[['country', 'cases']].head(3), headers='keys', tablefmt='pretty'))

print("\nTop 3 Least Affected Countries:")
print(tabulate(leastAffected[['country', 'cases']].head(3), headers='keys', tablefmt='pretty'))

print("\nTop 3 Countries with Highest Death Rate:")
print(tabulate(highestDeathRate[['country', 'death_rate']].head(3), headers='keys', tablefmt='pretty'))

print("\nTop 3 Countries with Lowest Death Rate:")
print(tabulate(lowestDeathRate[['country', 'death_rate']].head(3), headers='keys', tablefmt='pretty'))

print("\nTop 3 Countries with Most Active Cases:")
print(tabulate(mostActiveCases[['country', 'active']].head(3), headers='keys', tablefmt='pretty'))

# Save DataFrame to Excel
df.to_excel("covid_report.xlsx", index=False)

# Auto-adjust column widths
wb = load_workbook("covid_report.xlsx")
ws = wb.active

for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2
    ws.column_dimensions[col_letter].width = adjusted_width

wb.save("covid_report.xlsx")

#Creating Visuals
createChart(mostAffected, 'cases', 'Top 10 Most Affected Countries by COVID-19 Cases', 'most_affected_cases')
createChart(leastAffected, 'cases', 'Top 10 Least Affected Countries by COVID-19 Cases', 'least_affected_cases', color='lightblue')
createChart(highestDeathRate, 'death_rate', 'Top 10 Countries with Highest Death Rate', 'highest_death_rate', color='red')
createChart(lowestDeathRate, 'death_rate', 'Top 10 Countries with Lowest Death Rate', 'lowest_death_rate', color='green')
createChart(mostActiveCases, 'active', 'Top 10 Countries with Most Active Cases', 'most_active_cases', color='orange')