import datetime
import pandas
import collections
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def make_wine_dict():
    wine_data = pandas.read_excel("wine.xlsx", sheet_name="Лист1",
                                  na_values=['N/A', 'NA'],
                                  keep_default_na=False)
    wine_data = wine_data.to_dict(orient='record')
    product_data = collections.defaultdict(list)
    for wine in wine_data:
        category = wine["Категория"]
        product_data[category].append(wine)
    return product_data


def spelling_year(age):
    year = "лет"
    if (age // 10) % 10 != 1:
        if age % 10 == 1:
            year = "год"
        elif age % 10 in (2, 3, 4):
            year = "года"
    return year


establishment_date = datetime.datetime(year=1922, month=1, day=1)
current_date = datetime.datetime.now()
winery_age = current_date - establishment_date
winery_age = winery_age.days//365
spelling_years = ""
splelling_years = str(spelling_year(winery_age))
wine_data = dict(make_wine_dict())

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)
template = env.get_template('template.html')
rendered_page = template.render(
    wines=wine_data,
    establishment_date=winery_age,
    spelled_year=splelling_years
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
