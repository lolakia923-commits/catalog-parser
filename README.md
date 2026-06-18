# Catalog Parser — сайт → Excel

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
[![Kwork](https://img.shields.io/badge/Kwork-Заказать-00a046?style=flat-square)](https://kwork.ru/user/nikitasokolov44)

Демо-парсер на Python: собирает товары с каталога и выгружает в **Excel** и **CSV**.

## Возможности

- Название, цена, наличие, рейтинг, категория, ссылка
- Оформленный Excel (заголовки, ширина колонок, гиперссылки)
- Настраиваемое число страниц

## Запуск

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python parser.py --pages 2 --out-dir .
```

## Пример результата

См. `example_output.xlsx` — 149 товаров.

## Стек

Python 3 · requests · BeautifulSoup4 · openpyxl

## Автор

[Nikita Sokolov](https://github.com/lolakia923-commits) · [Kwork](https://kwork.ru/user/nikitasokolov44)
