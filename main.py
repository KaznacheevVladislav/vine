from http.server import HTTPServer, SimpleHTTPRequestHandler
from pprint import PrettyPrinter
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime

import pandas as pd
import collections


def convert_excel_to_dict(io, sheet_name):

    excel_data_df = pd.read_excel(
        io=io,
        sheet_name=sheet_name,
        na_values='',
        keep_default_na=False
    )
    excel_data_df = excel_data_df.fillna('')

    vines = collections.defaultdict(list)

    if 'Категория' in excel_data_df.columns.tolist():
        for category in sorted(list(set(excel_data_df['Категория']))):
            vines[category] = excel_data_df[excel_data_df['Категория']
                                            == category].to_dict('records')

    return vines


vines = convert_excel_to_dict('wine3.xlsx', 'Лист1')

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

rendered_page = template.render(
    vines=vines,
    count_years=str(datetime.datetime.now().year - 1920)
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
