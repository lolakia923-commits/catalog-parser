# Catalog Parser — сайт → Excel

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
