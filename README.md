# CRM Import Converter

> **This build has an English user interface.** The same service with a Russian
> interface lives in a separate repository:
> [LeadTableeCSV](https://github.com/alexeyd-sudo/LeadTableeCSV).
> The code is the same; only the UI strings, messages and report labels differ.

A web service that turns an arbitrary lead spreadsheet (Excel/CSV) into a CSV
ready for CRM import.

```
py -m venv venv && venv\Scripts\activate
pip install -r requirements.txt
py app.py            ->  http://localhost:5001
```

The port comes from the `PORT` environment variable and defaults to **5001**
here (the Russian build defaults to **5000**), so both can run side by side:

```
set PORT=5050 && py app.py
```

## Documentation

| Document | Audience |
|---|---|
| **[docs/SUPPORT.md](docs/SUPPORT.md)** | support and end users: how to run it, what every option and report means, common problems |
| **[docs/DEVELOPER.md](docs/DEVELOPER.md)** | developers: architecture, data flow, HTTP API, extension points, production setup |
| **[docs/DATA-ANALYSIS.md](docs/DATA-ANALYSIS.md)** | the source-data analysis and the weak spots it exposed |
| **[docs/CHANGELOG.md](docs/CHANGELOG.md)** | what changed |

## How it works

1. **Upload a file** — `.xlsx`, `.xlsm` or `.csv`, up to 25 MB.
   The first row must contain the column headers.

2. **Scanning.** The service classifies the content of every column
   (e-mail / phone / website / text), finds the phone columns and pulls up to
   5 numbers out of each cell, normalising them to `+…` form. A cell holding
   several numbers is split into sub-columns `— Phone 1`, `— Phone 2`, … which
   can be mapped to different CRM fields.

3. **Field mapping.** The target fields are pre-selected automatically (from
   the header wording and from the actual content) — check them. Targets:

   * **Company Name / Lead Name** — one value written to 2 result columns
   * **Mobile / Home / Other Phone**
   * **Corporate / Other Website**
   * **Work / Home / Other E-mail**
   * **Country OUTREACH / Outreach comment / Comment** — optional
   * **Source** — not shown; every row gets `import`

4. **Result** — two files and two reports.

## Two key ideas

### Duplicates: two files, not a decision made for you

The service never decides on your behalf. It produces **both** the full file
(one row per source row) **and** the clean one (duplicates merged, data from
every copy collected into a single row), **plus** a CSV report saying which
rows were merged into which and on what grounds — with the original
spreadsheet row numbers.

The merge signals are individually toggleable: company name (including the
parenthesised handle — `Alpha (alpha_mx)` = `alpha_mx`), e-mail,
phone (last 9 digits, so a number with and without a country code counts as
one) and website.

On the real database: **436 rows → 355**, 57 duplicate groups, and rows with
no contact details at all drop from 35 to 14 — merging recovers data from the
sparser copies.

### Values in the wrong column are fixed automatically

The type is decided **by the content, not by the column header**:

| Signal | Type |
|---|---|
| contains `@` | e-mail |
| digits, `+`, brackets, dashes only | phone |
| `http(s)://` scheme, `www.` prefix, or a domain with a real TLD | website |

An e-mail typed into the phone column ends up in an E-mail field, a link ends
up in Website, a number from the website column ends up in Phone. If the value
is already present there, it is dropped as a duplicate. Every decision is
written to `field_fixes_report.csv` with the source row number, so the
automation can be audited.

## Output format

`;` as the delimiter, CRLF, quotes only where needed. The `;` matches
`reference_import.csv`, the file that imported correctly; a comma-separated file
made the whole row land in a single column.

The encoding is **chosen on step 3**, default **UTF-8 with BOM**. The BOM is
the marker Excel and most importers use to detect the encoding. Without it the
receiver reads the two UTF-8 bytes of `é` as two characters of its own ANSI
codepage, and after a pass through ASCII you get `México → M??xico` — two
question marks per letter, the signature of a double re-encoding.

If the CRM mangles letters under every encoding, there is a
**“Transliterate accented characters”** option (`México → Mexico`): the result
becomes pure ASCII, which no re-encoding can corrupt. See
[docs/SUPPORT.md](docs/SUPPORT.md).

## Tests

```
py tests\test_converter.py       # 40 tests, no dependencies
```

## Production

```
waitress-serve --host=0.0.0.0 --port=5001 --threads=4 app:app
```

Use exactly **1 worker** — session state lives in process memory.
Details and environment variables in [docs/DEVELOPER.md](docs/DEVELOPER.md).

## Limits

* Max upload 25 MB (`MAX_CONTENT_LENGTH` in `app.py`), 200 000 rows and
  300 columns (`MAX_ROWS`/`MAX_COLS` in `converter.py`).
* Sessions live 6 hours, then they are deleted together with the uploaded file
  (`SESSION_TTL_SECONDS`).
* At most 5 phone numbers are taken from one cell (`MAX_PHONES_PER_CELL`);
  anything above that is reported as a warning rather than dropped silently.
* **There is no authentication** — this is an internal-network tool. Put it
  behind a reverse proxy with auth before exposing it.

## Localisation

The interface, messages and reports are in English. The column-header keywords
used to recognise the *uploaded* file (`converter.py`: `PHONE_HEADER_KEYWORDS`,
`HEADER_TARGET_HINTS`, …) deliberately still include Spanish and Russian terms —
an English UI does not mean the spreadsheets themselves are in English.
