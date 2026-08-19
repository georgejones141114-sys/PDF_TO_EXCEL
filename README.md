# RSE Report Converter

A tiny web app: upload a Rwanda Stock Exchange market report PDF, get back a
styled `.xlsx` workbook with the equities, bonds, weekly report, and market
stats tables already parsed out into clean sheets.

It's a Flask wrapper around the same extraction logic as the original
`PDF_to_Excel.py` script, generalized to work directly from an uploaded PDF
(no pre-existing template workbook required) and to locate each section by
its text markers rather than hardcoded page numbers.

## What it produces

One workbook, seven sheets:

| Sheet | Contents |
|---|---|
| `STOCK` | Security, closing price, volume, value |
| `MARKET STATS` | RSI, ALSI, equity turnover, bond turnover, market cap |
| `EXCHANGE RATE` | Buying / selling rates per currency |
| `EQUITIES MARKET` | Full equities trading table |
| `BONDS_GOV+CORP` | Government + corporate bonds trading table |
| `BONDS TRADES` | Bonds that traded today (see note below) |
| `WEEKLY REPORT` | This week's daily volume / value / deals |

**Note on `BONDS TRADES`:** the PDF only reports a single total traded
volume per bond for the day, not a deal-by-deal breakdown. This sheet lists
one row per bond that traded (best effort). If you need the deal-by-deal
reconstruction that appears in some hand-built workbooks, that information
isn't present in the source PDF and can't be derived from it automatically.

## Running locally (macOS)

Requires Python 3.10+. Check what you have with `python3 --version`; if
it's missing or too old, install it with `brew install python3`.

```zsh
git clone <this-repo-url>
cd rse-pdf-to-excel
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
PORT=5001 python3 app.py
```

Then open http://localhost:5001, drop in a PDF, and download the workbook.

When you're done, leave the virtual environment with `deactivate`.

## Running the tests

```zsh
source .venv/bin/activate
pip install -r requirements-dev.txt
pytest tests/ -v
```

## Deploying

GitHub Pages only serves static files, so it can't run this Flask app
directly. Push this repo to GitHub, then deploy the app itself to any
Python host, for example:

**Render**
1. New → Web Service → connect this repo.
2. Build command: `pip install -r requirements.txt`
3. Start command: `gunicorn app:app`

**Railway**
1. New Project → Deploy from GitHub repo.
2. Railway auto-detects `requirements.txt`; set the start command to
   `gunicorn app:app` if it isn't picked up automatically.

**Fly.io / any Docker host**
The app also runs fine behind any WSGI server — `gunicorn app:app` is the
production entry point (see `Procfile`).

Once deployed, you can link to the hosted app from this repo's GitHub Pages
site (or just from the README) if you want a "front door" on GitHub Pages
itself.

## How the parsing works

`converter.py` does the heavy lifting:

1. Extracts text from every page with `pdfplumber`. The cover page's
   "Market overview" / "Closing bell" section is genuinely two columns of
   prose side by side — extracted naively, pdfplumber interleaves their
   lines. The converter detects that page by its section headings and
   extracts the left/right halves separately so sentences aren't split
   across columns.
2. Finds each data table (equities, government bonds, corporate bonds,
   weekly report) by matching row patterns (ISIN codes, date-prefixed
   rows, etc.) rather than by page number, so a report that gains or loses
   a page still parses.
3. Pulls scalar figures (RSI/ALSI, turnover, market cap, exchange rates)
   out of the overview and indices sections with targeted regexes.
4. Writes everything into a styled `openpyxl` workbook (grey borders,
   bold+filled header row, frozen header, auto-sized columns).

If a future report changes its wording or table layout significantly, the
converter will likely need small regex updates in `converter.py` — the
functions are split one-per-section to make that easy to find.

## Limitations

- Built and tested against the standard RSE daily/weekly report template.
  Reports with a substantially different layout may not parse correctly —
  the app will show an error rather than silently producing a wrong
  workbook when it can't find the expected tables at all, but it can't
  catch every possible layout drift.
- `BONDS TRADES` is a best-effort aggregate, not a deal-by-deal
  reconstruction (see note above).
- Max upload size is 20 MB.
