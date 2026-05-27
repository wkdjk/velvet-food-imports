# Velvet Food Imports

Korea MFDS deer velvet import tracker. Ingests weekly MFDS xlsx files, stores records in Google Sheets, and publishes an HTML report to GitHub Pages.

## Architecture

```
MFDS xlsx upload → scripts/ingest_mfds.py → Google Sheets → scripts/generate.py → docs/index.html → GitHub Pages
```

## Requirements

- Python 3.10+
- Google service account with editor access to the master Sheets file

## Setup

1. Copy `.env.example` to `.env` and fill in real values.
2. Install dependencies: `pip install -r requirements.txt`
3. Run historical ingest once: `PYTHONPATH=. python scripts/ingest_historical.py`
4. For each new MFDS file: place it in `data/mfds/` then run `PYTHONPATH=. python scripts/ingest_mfds.py`

## PYTHONPATH

All scripts must be run with `PYTHONPATH=.` set — either in the command or in your shell session.
Without it, `src/` imports will fail silently.

```bash
export PYTHONPATH=.
python scripts/ingest_historical.py
```

## Data files

- `data/historical/` — Commander's historical xlsx (not committed; see `.gitignore`)
- `data/mfds/` — MFDS weekly xlsx files (not committed; renamed to `*_done.xlsx` after processing)

## Environment variables

See `.env.example` for all required variables. All secrets are loaded via `os.environ` — no literal credentials appear in any code or file.

## GitHub Actions

- `.github/workflows/ingest_mfds.yml` — triggered on push of new xlsx to `data/mfds/**`
- `.github/workflows/daily_build.yml` — cron KST 07:00 daily; regenerates report from existing Sheets data

## Folder structure

```
src/
  sheets/       — Google Sheets API wrapper
  parsers/      — xlsx parsers (historical + MFDS)
  aggregator/   — monthly, by_importer, pivot, by_country
  publisher/    — HTML report generator
scripts/        — CLI entry points
templates/      — Jinja2 HTML templates
static/         — CSS and assets
data/
  historical/   — Commander xlsx (git-ignored)
  mfds/         — MFDS weekly files (git-ignored)
docs/           — Generated report (committed to trigger Pages)
tests/          — pytest unit and integration tests
```
