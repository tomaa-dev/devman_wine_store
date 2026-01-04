from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
from collections import defaultdict
import argparse
import os


def get_year_word_form(delta):
    plural_form = 'лет'
    single_form = 'год'
    few_form = 'года'
    few_set = {2, 3, 4}
    if 11 <= delta % 100  <= 14:
        return plural_form
    if delta % 10 == 1:
        return single_form
    elif delta % 10 in few_set:
        return few_form
    else:
        return plural_form


def main():
    default_dir = os.path.join(os.path.dirname(__file__), 'wine3.xlsx')
    parser = argparse.ArgumentParser(description="""Программа читает данные о винах из Excel-файла, подставляет их в Jinja2-шаблон и генерирует статическую страницу index.html, 
                                    а затем запускает простой HTTP-сервер для локального просмотра.""")
    parser.add_argument('--directory', type=str, default=default_dir, help='Позволяет указать другой путь к файлу')
    args = parser.parse_args()
    directory = args.directory

    now = datetime.datetime.now()
    event1 = now.year
    event2 = 1920
    delta = event1-event2

    excel_data_df = pandas.read_excel(
        directory,
        usecols=['Категория', 'Название','Сорт', 'Цена', 'Картинка', 'Акция'],
        na_values=['N/A', 'NA'],
        keep_default_na=False
    )

    wines = excel_data_df.to_dict(orient='records')
    wine_categories = defaultdict(list)
    for wine in wines:
        wine_categories[wine['Категория']].append(wine)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    rendered_page = template.render(
        wine_categories=wine_categories,
        age_of_the_winery=delta,
        year=get_year_word_form(delta)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()