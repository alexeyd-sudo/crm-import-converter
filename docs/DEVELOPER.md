# Developer guide

Technical walkthrough: architecture, data flow, HTTP API, extension points.
The user-facing side is in [SUPPORT.md](SUPPORT.md); the source-data analysis
and the reasoning behind the design is in [DATA-ANALYSIS.md](DATA-ANALYSIS.md).

---

## 1. Stack and layout

Python 3.9+, Flask, openpyxl. No database, no queues, no frontend build step.

```
Converter/
├── app.py                    Flask: routes, sessions, files
├── converter.py              all parsing/cleaning/merging logic (Flask-free)
├── requirements.txt
├── templates/index.html      single-page UI
├── static/app.js             ~330 lines of vanilla JS, no dependencies
├── static/style.css
├── tests/test_converter.py   40 tests, runs without pytest
├── uploads/<session_id>/     the source file + 4 generated CSVs
└── docs/
    ├── SUPPORT.md
    ├── DEVELOPER.md
    ├── DATA-ANALYSIS.md
    └── CHANGELOG.md
```

`converter.py` **does not import Flask** on purpose: it can be used as a
library and tested on its own.

---

## 2. Data flow

```
      file
        │  read_table()  → xlsx: openpyxl (+ hyperlinks), csv: sniffer + 4 encodings
        ▼
  headers[], rows[][], hyperlinks[][]
        │  analyze_columns()   classifies every cell of a sample
        ▼
  col_stats[]  {counts:{email,phone,url,text}, dominant, share, header_hint}
        │  detect_phone_columns()  header OR ≥50% of cells look like a phone
        ▼
  phone_cols[]
        │  scan_phones()   per cell: phone numbers + "leftovers" (e-mails/links)
        ▼
  phone_data{col: {per_row, leftovers, max_count, stats}}
        │  build_source_columns()  the flat list shown on the left of the UI
        ▼
  source_columns[]  ── suggest_mapping() ──► pre-selected dropdowns
        │
        │  ▼ the operator edits mapping + options
        │
        │  build_records()   atomisation + routing of wrong-type values
        ▼
  records[] (dict keyed by OUTPUT_HEADERS + _row),  fixes[]
        │  dedupe_records()   union-find over name/e-mail/phone/website
        ▼
  merged[], groups[]
        │  records_to_rows() + write_csv()
        ▼
  4 files: result_raw.csv, result_merged.csv,
           report_duplicates.csv, report_fixes.csv
```

---

## 3. Key concepts

### 3.1 Atoms (`extract_atoms`)

A cell is **not** assumed to hold what its column header promises. It is
decomposed into typed atoms:

```python
{'emails': [...],       # regex based, with the "gmail,com" → "gmail.com" fix
 'urls': [...],         # confident: scheme, www. prefix, or a TLD in KNOWN_TLDS
 'urls_loose': [...],   # get_url() fallback - the whole cell as a single URL
 'phones': [...],       # normalised, '+'-prefixed
 'phone_status': [...], # ok | fixed | cc_added | invalid, one per phone
 'overflow': 0,         # numbers that did not fit MAX_PHONES_PER_CELL
 'text': '...'}         # the cleaned cell as a whole
```

Order matters: e-mails are cut out first (otherwise `info2024@x.com` yields a
"phone" of `2024`), then `wa.me/...` becomes a phone, then the remaining URLs,
then the digits.

**Type rules** (`classify_value`):

| Type | Signal |
|---|---|
| e-mail | contains `@` and matches `EMAIL_RE` |
| website | `http(s)://` scheme, a `www.` prefix, or a domain whose TLD is in `KNOWN_TLDS` |
| phone | ≥7 digits and almost nothing besides digits, `+`, brackets, dashes, spaces |
| text | everything else |

`KNOWN_TLDS` is deliberately **not the full IANA list**: in this data the
Instagram handles look like domains (`alpha.cw`, `distribuidora.beta`), and a
full list would turn them into "websites". If you widen it, check that
`test_classify_text_not_url` still passes.

### 3.2 Routing of misplaced values (`build_records`)

A row is assembled in two passes:

1. **Direct pass.** For every (column → target field) pair, atoms of the
   *matching* type go into that field's bucket. Atoms of a *foreign* type go
   into `pending`. Fields of type `text` take the whole cell and never give
   anything away — that is what stops a company name like `alpha.cw` from
   drifting into Website.
2. **Placement pass.** For every foreign value, `TYPE_CHAINS[kind]` is walked
   to pick a destination:
   1. a field of that type the operator mapped which is **empty** in this row;
   2. a field of the chain the operator did not map at all;
   3. otherwise the **last** field of the chain (`Other Website` /
      `Other Phone Number`), appending with `, `.

   If the value already exists in any field of that type it is dropped and
   marked `duplicate`. Every decision is recorded in `fixes[]`.

Disabled with `autofix_types: false`.

### 3.3 Rescuing data from a "phone" column

A column split into sub-columns (`"2:0"`, `"2:1"`, …) used to make the original
cell text unreachable. Now `scan_phones()` also stores `leftovers` — the
e-mails and links found in the same cell — and `_atoms_for_key()` hands them
out together with sub-column `:0`. As a result:

* an e-mail column misdetected as a phone column **is not lost** — it can be
  pointed at `Work E-mail` (test
  `test_mis_detected_phone_column_does_not_lose_data`);
* a link sitting next to a number still reaches Website.

### 3.4 Duplicates (`dedupe_records`)

Union-find over type-prefixed keys (so a phone can never collide with a name):

| Key | How it is built | Option |
|---|---|---|
| `n:` | `alnum(name)`; for `Name (handle)` **two** keys, from the name and the handle | `dedupe_by_name` |
| `e:` | every address from Work/Home/Other E-mail, lowercased | `dedupe_by_email` |
| `p:` | the last **9 digits** of every number (with and without country code become one key) | `dedupe_by_phone` |
| `w:` | the domain; for social networks the domain plus the account (`instagram.com/alpha`) | `dedupe_by_website` (off by default) |

The group root is the **lowest index**, so the output preserves
first-occurrence order.

Merging a group (`_merge_group`): the name comes from `pick_best_name()`
(prefers a variant containing a space, then the longer one); phones go through
`dedupe_join_phones()` (which collapses "same number without country code");
websites are keyed by `url_key`; everything else is `dedupe_join` with `, `.

`dedupe_records()` is **never applied to the raw file** — that one is always
written as-is. This is deliberate: the operator compares the two files and
decides.

### 3.5 Output encoding

`é` in UTF-8 is 2 bytes (`C3 A9`). A receiver that does not know the encoding
reads them as 2 characters of its own ANSI codepage (`Ã©` / `Г©`), and a later
pass through ASCII turns each into `?` — hence `México → M??xico` (**two**
question marks per letter). A single `?` means the ASCII reduction happened
straight away, without the intermediate step.

That is why the default is **UTF-8 with BOM** (`CSV_ENCODING = 'utf-8-sig'`):
the BOM is the only in-band encoding signal Excel and most importers
understand. `utf-8`, `cp1252` and `cp1251` are also available
(`CSV_ENCODINGS`, option `csv_encoding`).

`write_csv()` **never raises** on an unrepresentable character.
`fit_to_encoding()` degrades in steps: character as-is →
`strip_accents(character)` → `?`, and returns
`{'encoding', 'replaced', 'affected_rows'}` so the UI can report the loss.

`strip_accents()` (option `transliterate`) removes diacritics **from Latin
script only**: the NFKD decomposition is applied when the base character is
ASCII. Cyrillic is untouched — otherwise `й` would become `и`, which is
corruption rather than transliteration. Characters with no decomposition
(`ß`, `®`, `—`, `№`) come from `MANUAL_TRANSLIT`. The result is pure ASCII,
which re-encoding cannot damage.

The reports are always written as `utf-8-sig` regardless of the option: they
are read in Excel, not imported.

### 3.6 Reports

`_row` on a record is the row number **as in Excel** (`row_idx + 2`, because
row 1 is the header). The same numbers go into both reports.

---

## 4. HTTP API

### `POST /api/upload`

`multipart/form-data`, field `file`.

```jsonc
{
  "session_id": "1572d983…",           // 32 hex chars
  "row_count": 436,
  "source_columns": [                   // left-hand side of the UI
    {"key": "2:0", "label": "Phone … — Phone 1", "kind": "phone",
     "samples": ["+522281110011"], "stats": {...}, "analysis": {...}}
  ],
  "target_fields": [...],               // = converter.TARGET_FIELDS
  "phone_summary": [{"header": "...", "max_count": 5, "stats": {...}}],
  "suggested_mapping": {"0": "company_lead", "1": "work_email", ...},
  "warnings": ["Phone column \"…\" contains 11 links …"],
  "defaults": {...},                    // = converter.DEFAULT_OPTIONS
  "encodings": [{"id": "utf-8-sig", "label": "…", "hint": "…"}, ...]
}
```

Column `key`: `"3"` is plain column #3; `"2:1"` is the second phone extracted
from column #2.

Errors: `400` (no file / wrong type / empty / table too large),
`413` (> 25 MB, also JSON).

### `POST /api/convert`

```jsonc
{
  "session_id": "1572d983…",
  "mapping": {"0": "company_lead", "2:0": "mobile_phone"},
  "options": {
    "autofix_types": true,
    "default_country_code": "52",
    "skip_rows_without_name": false,
    "sanitize_formulas": false,
    "csv_encoding": "utf-8-sig",       // utf-8-sig | utf-8 | cp1252 | cp1251
    "transliterate": false,            // é -> e, ñ -> n
    "dedupe_by_name": true, "dedupe_by_email": true,
    "dedupe_by_phone": true, "dedupe_by_website": false
  }
}
```

Response:

```jsonc
{
  "headers": [...],
  "encoding": {"id": "utf-8-sig", "label": "UTF-8 with BOM (recommended)",
               "transliterate": false, "replaced": 0, "affected_rows": 0},
  "raw":     {"row_count": 436, "filled": {...}, "preview": [[...]], "download_url": "/download/…/raw"},
  "merged":  {"row_count": 355, "filled": {...}, "preview": [[...]], "download_url": "/download/…/merged"},
  "duplicates": {"group_count": 57, "collapsed_rows": 81,
                 "groups": [{"name": "...", "size": 3, "rows": [3,186,353],
                             "names": [...], "matched_on": ["name","phone"]}],
                 "download_url": "/download/…/duplicates"},
  "fixes": {"moved": 2, "dropped_duplicates": 10,
            "items": [{"row": 63, "column": "…", "kind": "url", "value": "…",
                       "from": "mobile_phone", "to": "other_website", "action": "moved"}],
            "download_url": "/download/…/fixes"},

  // legacy fields duplicating `raw`, so older scripts keep working
  "row_count": 436, "preview": [...], "filled": {...}, "download_url": "/download/…/raw"
}
```

`groups` and `items` are capped at 50 entries in JSON; the full lists are in
the CSV reports.

Errors: `400` (malformed `session_id` / empty mapping), `404` (session expired).

### `GET /download/<session_id>/<kind>`

`kind` ∈ `raw | merged | duplicates | fixes`. `GET /download/<session_id>`
without `kind` means `raw` (backwards compatibility).

`session_id` is validated against `^[0-9a-f]{32}$`, otherwise `404`.

### `GET /healthz`

`{"status": "ok", "sessions": N}`.

---

## 5. Sessions and files

The store is a plain in-process `dict` behind a `threading.Lock`. A session
holds `headers`, `rows`, `hyperlinks`, `phone_data`, `col_stats`, `labels`,
`filename`, `created`.

`cleanup_sessions()` runs on every `/api/upload`:

* drops sessions older than `SESSION_TTL_SECONDS` (default 6 h) together with
  their directory;
* if there are more than `MAX_SESSIONS` (50), removes the oldest;
* deletes orphan directories in `uploads/` left over from a previous process.

Both settings are read from the environment.

> **One worker only.** State lives in process memory. With several workers an
> `/api/convert` request can hit a process that never saw the `/api/upload`.
> If you need to scale, move `SESSIONS` into Redis or serialise the session to
> `uploads/<sid>/session.pickle`.

---

## 6. Adapting to a different CRM

| What to change | Where |
|---|---|
| The set of CRM fields | `TARGET_FIELDS` + `OUTPUT_HEADERS` in `converter.py`. The `type` (`text`/`phone`/`url`/`email`) drives both parsing and relocation |
| Where surplus values spill | `TYPE_CHAINS` (+ `PHONE_OUTPUTS` / `EMAIL_OUTPUTS` / `URL_OUTPUTS`) |
| CSV delimiter / line endings | `CSV_DELIMITER`, `CSV_LINETERMINATOR` |
| Default encoding and the available list | `CSV_ENCODING`, `CSV_ENCODINGS` |
| Transliteration table | `MANUAL_TRANSLIT` (characters with no NFKD decomposition) |
| The `Source` column value | `SOURCE_VALUE`, or `options['source_value']` |
| Max phone numbers per cell | `MAX_PHONES_PER_CELL` |
| Header-based mapping suggestions | `HEADER_TARGET_HINTS` |
| Phone-column keywords | `PHONE_HEADER_KEYWORDS`, `NON_PHONE_HEADER_SUBSTRINGS`, `NON_PHONE_HEADER_WORDS` |
| Domain zones | `KNOWN_TLDS`, `SOCIAL_HOSTS`, `PHONE_LINK_HOSTS` |
| Tracking params that get stripped | `TRACKING_PARAMS` |
| Size limits | `MAX_CONTENT_LENGTH` (app.py), `MAX_ROWS` / `MAX_COLS` (converter.py) |

The header keyword lists intentionally contain Spanish and Russian terms:
the interface is English, but the uploaded spreadsheets are not necessarily.

Adding a new target field:

1. an entry in `TARGET_FIELDS` (`id`, `label`, `type`, `outputs`);
2. the names from `outputs` go into `OUTPUT_HEADERS`;
3. if the type is strict, add the `id` to the right `TYPE_CHAINS` and to
   `PHONE_OUTPUTS`/`EMAIL_OUTPUTS`/`URL_OUTPUTS`;
4. optionally a rule in `HEADER_TARGET_HINTS` and a branch in
   `suggest_mapping()`.

No frontend change is needed: the field list comes from the server.

---

## 7. Tests

```
py tests\test_converter.py     # no dependencies, prints PASS/FAIL
py -m pytest tests             # if pytest is installed
```

`test_real_file_regression` walks up from `tests/` looking for the real
`leads.xlsx` and silently skips if it is not there.

Covered: value classification, phone-column false positives, recovery of
misplaced fields, phone normalisation and the default country code, every
duplicate-merging mode, report row numbering, the CSV dialect, output
encodings and transliteration, formula protection, mapping suggestions.

**Run the tests whenever you touch `converter.py`** — nearly every one of them
pins down a specific bug that was found in real data.

---

## 8. Running

### Development

```
py -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
py app.py                       # http://127.0.0.1:5001 (FLASK_DEBUG=1 for the debugger)
```

### Production

Copy-paste configuration is in [`deploy/`](../deploy/): systemd unit, env file,
nginx vhost, tmpfiles rule.

```
gunicorn --workers 1 --threads 2 --bind 127.0.0.1:5101 app:app
```

Exactly **one worker** (see §5). This is not advice, it is a correctness
requirement: measured with `gunicorn -w 2`, 5 of 20 `/api/convert` calls returned
`404 Session not found` because they landed on the process that never saw the
upload. `--threads 2` rather than 4 because parsing is CPU-bound under the GIL,
so a bigger pool only lets one client starve everyone else — with 4 threads busy,
`/healthz` was measured blocked for 16 s.

Never run `app.py` directly on a server. It binds to `127.0.0.1` and leaves the
debugger off unless `FLASK_DEBUG=1`; setting that on a public host hands out
remote code execution through the Werkzeug console.

Environment variables — see the table in the README. Summary: `PORT`,
`UPLOAD_DIR`, `SESSION_TTL_SECONDS`, `MAX_SESSIONS`, `MAX_CONTENT_LENGTH`,
`MAX_XLSX_UNCOMPRESSED`, `MAX_ROWS`, `MAX_COLS`, `HOST`, `FLASK_DEBUG`.

Put a hard `MemoryMax` on the unit. The service holds every live session's parsed
table in RAM, so on a shared box a ceiling decides whether a careless upload gets
*this* unit killed or lets the kernel OOM-killer pick a neighbouring service.

---

## 9. Security

Done:

* `session_id` is strictly `^[0-9a-f]{32}$`; path traversal is closed;
* extension allow-list, upload cap, row/column limits — and, the one that
  matters, a cap on the **uncompressed** size of the `.xlsx` archive, checked
  from the ZIP directory before any XML is parsed. `MAX_ROWS` alone cannot stop
  a memory blow-up: openpyxl builds its object model before a row count exists,
  so the guard used to fire only after the RAM was already spent. A 70 KB file
  expanding to 67 MB of XML was accepted;
* every client-supplied mapping key is matched against `SOURCE_KEY_RE` and every
  option is coerced to its expected type. Malformed JSON returns `400`, not the
  `500` it used to — `int('abc')`, `mapping` as a list and a dict in
  `source_value` were all reachable crashes;
* session TTL is enforced on every request, not only at the next upload;
* uploads live in `UPLOAD_DIR` (`0700`, files `0600`), configurable so that lead
  data need not sit inside the code directory;
* result files are written to a temporary name and `os.replace()`d into place, so
  a download can never serve a half-written or mixed-conversion file;
* downloads carry `Cache-Control: private, no-store`;
* both CSV reports are **always** escaped against formula injection — they exist
  to be opened in Excel, so that cannot be an option. The main CSV keeps the
  opt-in switch, and phone-shaped values (`+` followed only by digits and
  separators) are exempt, which is what made that option unusable before;
* all HTML insertion points on the frontend go through `escapeHtml`.

Deliberately not done — **must be provided by the reverse proxy**
(`deploy/nginx.conf` does all three):

* **no authentication** — anyone who can reach the port can upload files and
  download someone else's results if they know the `session_id`. It handles
  customer contact details, so basic auth / SSO in front is mandatory, not
  optional;
* no CSRF protection. The API is cookie-less, so there is nothing to steal, but
  `multipart/form-data` is a CORS-simple content type: a third-party page can
  make a visitor's browser POST to `/api/upload` and burn server CPU. Rate
  limiting in the proxy is what contains this;
* no rate limiting;
* `openpyxl` parses arbitrary `.xlsx`. It does not execute formulas or VBA, and
  XXE was verified closed, so the realistic risk is resource consumption — which
  is what `MAX_XLSX_UNCOMPRESSED` plus `MemoryMax` are for.
