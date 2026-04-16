# DOI Title Matcher

`DOI Title Matcher` is a small Python script that looks up paper titles through Crossref and OpenAlex, then writes matched DOIs into a new Excel file.

## Features

- Reads titles from an Excel workbook
- Tries Crossref first and falls back to OpenAlex when needed
- Uses a similarity threshold to reduce incorrect matches
- Writes a clean two-column Excel output with auto-sized columns

## Requirements

- Python 3.10+
- Network access to the Crossref and OpenAlex APIs

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python fill_doi_from_titles.py \
  --input input.xlsx \
  --output output.xlsx \
  --title-column title \
  --output-title-column title \
  --email your_email@example.com \
  --workers 4 \
  --min-score 0.90
```

## Arguments

- `--input`: input `.xlsx` file
- `--output`: output `.xlsx` file
- `--title-column`: source column containing article titles; defaults to the first column
- `--output-title-column`: title column name used in the output file; defaults to `title`
- `--email`: contact email used for polite API access
- `--workers`: concurrent lookup workers; defaults to `4`
- `--min-score`: minimum title similarity threshold; defaults to `0.90`
- `--delay`: delay after each lookup request in seconds; defaults to `0.2`
- `--limit`: optional row limit for quick testing

## Output

The generated workbook contains two columns:

- `title`
- `doi`

## Notes

- This repository intentionally excludes local data files, generated result files, and virtual environments.
- If your source column is not named `title`, set `--title-column` explicitly.
- Tune `--min-score` and `--delay` based on your data quality and API usage needs.
