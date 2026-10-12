# Book Catalog Web Scraper (Python)

Scrapes every book from [books.toscrape.com](https://books.toscrape.com) — a website built for scraping practice — and exports the data to Excel and CSV with a ready-to-read summary.

The same approach works for e-commerce catalogs, price monitoring, real-estate listings or directories.

## What it does

- **Crawls** all catalog pages automatically by following the "next" button (50 pages, 1,000 books)
- **Extracts** title, price, star rating, availability and product URL for each book
- **Cleans** the data: prices become numbers, star ratings become 1–5
- **Handles errors**: retries failed requests and waits between pages to avoid overloading the server
- **Exports** a formatted Excel file and a CSV

## Output

`libros.xlsx` contains:

| Sheet | Content |
| --- | --- |
| Resumen | Total books, average/min/max price, average rating, books in stock |
| Por calificación | Number of books and average price per star rating |
| Top 10 más caros | The 10 most expensive books |
| 5 estrellas más baratos | The cheapest 5-star books (best value) |
| Todos los libros | The full dataset |

## How to run

```bash
pip3 install -r requirements.txt
python3 scraper_libros.py               # all pages
python3 scraper_libros.py --paginas 3   # only the first 3 pages (quick test)
```

Options: `--paginas N` (max pages), `--url` (start page), `--salida` (output file name).

## Tech stack

Python · requests · BeautifulSoup · pandas · openpyxl

## Author

Santiago — Python developer specializing in automation, web scraping and data processing.
