from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

from pprint import pprint

import datetime

import openpyxl
import xlrd
import pandas as pd

import collections as cl

now = datetime.datetime.now()
word = ''
year_est = 1920

if (now.year - year_est)%10 == 1:
    word = 'год'
elif (now.year - year_est)%10 == 2 or (now.year - year_est)%10 == 3 or (now.year - year_est)%10 == 4:
    word = 'года'
elif (now.year - year_est)%10 == 0 or (now.year - year_est)%10 == 5 or (now.year - year_est)%10 == 6 or (now.year - year_est)%10 == 7 or (now.year - year_est)%10 == 8 or (now.year - year_est)%10 == 9:
    word = 'лет'

data = pd.read_excel('wine3.xlsx', na_values='nan', keep_default_na=False)

wines = data.to_dict(orient='records')

new_dict = cl.defaultdict(list)

for drinks in wines:
    new_dict[drinks.get('Категория')].append(drinks)

pprint (new_dict)

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)
template = env.get_template('template.html')

rendered_page = template.render(
    age = now.year - year_est,
    correct_word = word,
    wines = wines,
    collection = sorted(new_dict.keys())
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
