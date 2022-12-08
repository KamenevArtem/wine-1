import datetime
import pandas
import collections
import argparse
from pprint import pprint
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def collect_input_data(file_path="wine.xlsx"):
    input_data = pandas.read_excel(file_path, sheet_name="Лист1",
                                  na_values=['N/A', 'NA'],
                                  keep_default_na=False)
    input_data = input_data.to_dict(orient='record')
    wine_info = collections.defaultdict(list)
    for wine in input_data:
        category = wine["Категория"]
        wine_info[category].append(wine)
    return wine_info


def get_correct_spelling(age):
    year = "лет"
    if age % 10 == 1:
        year = "год"
    if age % 10 in (2, 3, 4):
        year = "года"
    if (age//10) % 10 == 1:
        year = "лет"
    return year


def parse_arg():
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--directory', nargs='?',
                        help='Path to input data file',
                        default="wine.xlsx")
    arg = parser.parse_args()
    return arg


def main():
    arg = parse_arg()
    file_path = arg.directory
    establishment_year = 1922
    current_date = datetime.datetime.now()
    winery_age = current_date.year - establishment_year
    splelling_years = get_correct_spelling(winery_age)
    wine_info = collect_input_data(file_path)
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    
    template = env.get_template('template.html')
    rendered_page = template.render(
        wines=wine_info,
        establishment_date=winery_age,
        spelled_year=splelling_years
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
    
    
if __name__ == "__main__":
    main()