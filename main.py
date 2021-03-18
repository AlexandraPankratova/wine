import argparse
import collections as cl
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    parser = argparse.ArgumentParser(
        description='Программа развертывает сайт магазина авторского вина')
    parser.add_argument('File', help='Название файла с исходными данными')
    file_name = parser.parse_args()

    current_date = datetime.datetime.now()
    company_age = current_date.year - 1920

    if company_age % 10 == 1:
        right_term_for_year = 'год'
    elif company_age % 10 in (2, 3, 4):
        right_term_for_year = 'года'
    else:
        right_term_for_year = 'лет'

    shop_stock = pd.read_excel(
        file_name.File,
        na_values='nan',
        keep_default_na=False,
    )

    wines = shop_stock.to_dict(orient='records')

    categorised_shop_stock = cl.defaultdict(list)

    for drinks in wines:
        categorised_shop_stock[drinks.get('Категория')].append(drinks)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml']),
    )
    template = env.get_template('template.html')

    rendered_page = template.render(
        age=company_age,
        year=right_term_for_year,
        categorised_drinks=sorted(categorised_shop_stock.items()),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __name__ == '__main__':
    main()
