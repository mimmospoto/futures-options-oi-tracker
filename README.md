# Futures Options Open-Interest Tracker

[![tests](https://github.com/mimmospoto/futures-options-oi-tracker/actions/workflows/tests.yml/badge.svg)](https://github.com/mimmospoto/futures-options-oi-tracker/actions/workflows/tests.yml)

Track how **open interest** in futures options builds up strike by strike, day after day, and get it as an Excel report with charts.

Open interest (OI) is the number of option contracts still open at each strike. Where it piles up, and how it moves over time, shows where traders are positioned. Large call or put OI can act as support or resistance, and it matters around expiry. Barchart shows only today's snapshot. This tool saves a snapshot each day, so you can see the changes over time.

It was first built to follow **coffee (KC)** options, but it works with any futures root listed on [barchart.com](https://www.barchart.com), such as `CL` (crude oil), `ZC` (corn) or `GC` (gold).

## What you get

Each run:

1. Fetches the nearest *N* futures contracts for a root, for example `KCZ26`, `KCH27` and so on.
2. Downloads the full option chain for each contract: strike, call or put, and open interest.
3. Appends it to a CSV history, `data/<ROOT>_history.csv`. That file is the tool's database. Running it again on the same day replaces that day's data instead of duplicating it.
4. Rebuilds an Excel report, `data/<ROOT>_open_interest.xlsx`, containing:
   - **Summary sheet:** one row per contract with the underlying price, total call and put OI, the **put/call ratio**, the strikes with the highest call and put OI, and the change in OI since the previous run.
   - **One sheet per contract:** a *date × strike* table of call OI and one of put OI, using the strikes around the money. Each table has a bar chart comparing the last 5 days.

Example output from a run on 2026-09-25, coffee:

| Contract | Underlying | Call OI | Put OI | Put/Call | Max call OI strike | Max put OI strike |
|---|---|---|---|---|---|---|
| KCZ26 | 278.60 | 45,270 | 76,865 | 1.70 | 300 | 250 |
| KCH27 | 271.25 | 16,671 | 29,155 | 1.75 | 300 | 200 |
| KCK27 | 268.90 | 549 | 5,905 | 10.76 | 320 | 125 |
| KCN27 | 267.60 | 139 | 344 | 2.47 | 270 | 227.5 |

## Installation

Requires Python 3.9+ and Google Chrome. Selenium downloads a matching chromedriver automatically.

```bash
git clone https://github.com/mimmospoto/futures-options-oi-tracker.git
cd futures-options-oi-tracker
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python -m oi_tracker KC                 # coffee, 4 nearest contracts, 10 strikes each side
python -m oi_tracker CL -c 6 -s 15      # crude oil, 6 contracts, 15 strikes each side
python -m oi_tracker ZC -o ~/oi-data    # write output to a different folder
```

| Option | Default | Meaning |
|---|---|---|
| `-c, --contracts` | 4 | Number of nearest futures contracts to track |
| `-s, --strikes` | 10 | Strikes shown on each side of the current price |
| `-o, --out` | `data` | Folder for the CSV history and the Excel report |
| `--show-browser` | off | Run Chrome visibly, which helps with debugging |

Run it once per trading day, ideally after the settlement data is published. A cron job does this for you:

```cron
# weekdays at 18:30
30 18 * * 1-5 cd /path/to/futures-options-oi-tracker && .venv/bin/python -m oi_tracker KC
```

## How it works

```
oi_tracker/
├── barchart.py  # headless Chrome session + Barchart JSON API calls
├── data.py      # raw chain -> tidy snapshot, strike window, CSV history
├── report.py    # Excel summary, tables and bar charts (openpyxl)
└── cli.py       # command-line interface
```

Barchart blocks plain HTTP clients. To get past that, the tool opens a headless Chrome session with Selenium and lets the site set its session cookies. It then calls Barchart's own JSON endpoint from inside the page with `fetch()`. This is far more robust than scraping the rendered HTML table, which is how the first version worked; you can find it in the git history.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

The tests use synthetic option chains, so they need neither a network connection nor a browser.

## Disclaimer

This project is for personal and educational use. The data comes from barchart.com, and you are responsible for following its [terms of use](https://www.barchart.com/terms). It is not investment advice.

## License

MIT
