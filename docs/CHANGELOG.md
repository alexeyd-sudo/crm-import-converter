# Changelog

## v2.1 — output encoding

Symptom: after import, names such as `Alpha Estética México (alphaestética)`
turned into `Medi M??xico Distribuidora (medim??xico)`.

**The cause was not in the service** — the output was valid UTF-8
(`M\xc3\xa9xico`). The file was written **without a BOM**, the receiver did not
recognise the encoding, read the 2 bytes of `é` as 2 characters of its own ANSI
codepage and then reduced the result to ASCII — hence **two** question marks
per letter. Tellingly, the reference `reference_import.csv`, the file that "imported
fine", contains no non-ASCII bytes at all: it already read
`Mediest?tica M?xico`.

### Added

* **Output encoding is selectable** in the UI: `UTF-8 with BOM` (the new
  default), `UTF-8 without BOM`, `Windows-1252`, `Windows-1251`. Each comes
  with a hint about when to use it.
* **“Transliterate accented characters” option** — `México → Mexico`,
  `Estética → Estetica`, `ā → a`, `® → (R)`. The result is pure ASCII, which
  re-encoding cannot corrupt. Cyrillic is left alone (NFKD would turn `й` into
  `и`, which is corruption rather than transliteration).
* **Loss report**: how many characters did not fit the chosen encoding and how
  many rows were affected, shown with the results.

### Changed

* **Default encoding `utf-8` → `utf-8-sig`** (UTF-8 with BOM). The BOM is the
  only in-band encoding signal Excel and most importers understand. The old
  behaviour is still available as "UTF-8 without BOM".
* The reports (`duplicates_report.csv`, `field_fixes_report.csv`) are always
  written as UTF-8 with BOM — they are opened in Excel, not imported.
* `write_csv()` no longer **raises** on an unrepresentable character: it
  degrades through "character → without accent → `?`" and returns statistics.
  Previously `Alphā Médica` in Windows-1252 would have thrown
  `UnicodeEncodeError` and returned a 500.
* The download `Content-Type` carries the real `charset`.

---

## v2.0 — duplicate handling and misplaced values

Driven by an analysis of the real source table — see
[DATA-ANALYSIS.md](DATA-ANALYSIS.md).

### Added

* **Two output files instead of one**: `crm_import_merged.csv` (duplicates
  merged) and `crm_import_with_duplicates.csv` (one row per source row).
  Duplicates were previously not handled at all — 436 rows in, 436 out.
* **Duplicate merging** (`dedupe_records`) — union-find over company name
  (including the parenthesised handle), e-mail, phone (last 9 digits) and
  website. Each signal is individually toggleable. On the real file:
  436 → 355 rows, 57 groups; rows with no contacts at all went from 35 to 14.
* **`duplicates_report.csv`** — which source rows ended up in which group and
  on what grounds.
* **Content-based type detection** (`classify_value`, `extract_atoms`):
  `@` → e-mail, digits only → phone, a domain with a real TLD → website.
* **Relocation of values found in the wrong column** — an e-mail from the phone
  column goes to E-mail, a link goes to Website, a number from the website
  column goes to Phone. Duplicates are dropped during relocation.
* **`field_fixes_report.csv`** — what was moved, from where, to where and why.
* **Automatic mapping suggestions** (`suggest_mapping`) from the header and the
  column content; highlighted with a blue border in the UI.
* **Column type badges** with a per-type breakdown in the tooltip.
* **Pre-conversion warnings**: links/e-mails in the phone column, numbers
  without a country code, phone-cap overflow, "mixed" columns.
* **Default country code** — numbers with no `+` and fewer than 11 digits get
  the given code (`6611110009` → `+52 6611110009`). Previously 108 numbers got
  a meaningless `+6611110009`.
* Options: do not relocate values between fields, skip rows without a company
  name, protect against Excel formula injection.
* `GET /healthz`, `GET /download/<sid>/<kind>`.
* **40 self-tests**, `py tests\test_converter.py` (runs without pytest).
* Documentation: SUPPORT.md, DEVELOPER.md, DATA-ANALYSIS.md.

### Fixed

* **False positives in phone-column detection.** Keywords were matched as
  substrings, so `Contact email`, `Warehouse` and `Hotel name` counted as phone
  columns. Short keywords (`wa`, `tel`, `cel`) now match whole words, `contact`
  was removed, a stop-list was added (`mail`, `website`, `id`, `zip`, `date`,
  `price`, …), and the content threshold went from 0.3 to 0.5, computed from a
  whole-cell verdict.
* **Data loss when that detection was wrong.** A column split into
  "Phone 1/2/3" made the original cell text unreachable. The e-mails and links
  from a phone cell are now kept separately and available for relocation.
* **Digits inside e-mails became phone numbers** (`info2024@x.com` → `2024`) —
  addresses are cut out before the number search.
* **`wa.me/...` landed in both the phone and the website field** — phone links
  are excluded from website recognition.
* **Phone cap per cell 3 → 5**; overflow is counted and surfaced as a warning
  rather than dropped silently. The mapping suggestion spills surplus
  sub-columns into Other Phone Number instead of leaving them unmapped.
* **The same number appeared in both Mobile and Home** — repetitions across
  fields of the same type are cleaned up.
* **Sessions and `uploads/` grew without bound** — TTL 6 h, 50-session cap,
  orphan-directory sweep.
* **`/download/<id>`** — strict `^[0-9a-f]{32}$` check instead of stripping
  slashes.
* **Error 413** (file > 25 MB) returned HTML, so the frontend crashed on
  `resp.json()`. Now JSON.
* `MAX_ROWS` / `MAX_COLS` limits; `TableTooLargeError` is no longer dead code.
* Broken session directories are removed when the file cannot be read.

### Compatibility

* `GET /download/<session_id>` without a suffix still serves the file with
  duplicates.
* `/api/convert` still returns `row_count`, `preview`, `filled` and
  `download_url` at the top level.
* `converter.build_output_rows()` is kept as a wrapper around `build_records()`.

---

## v1.0 — first version

Web interface: Excel/CSV upload, phone parsing, manual column-to-field mapping,
and a single CSV export in the `reference_import.csv` dialect (`;`, CRLF) — replacing
a one-shot script that wrote `,` plus quotes around every field, which made the
import fall apart with the whole row in one column.
