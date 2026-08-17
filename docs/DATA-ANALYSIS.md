# Source data analysis and the weak spots it exposed

This document records **what is wrong with the source data**
(`leads.xlsx`) and **which of those problems the service could not
handle** before it was reworked. Every figure comes from running the pipeline
over the real file. Row numbers are as shown in Excel (row 1 is the header).

---

## 1. The source table

| | |
|---|---|
| File | `leads.xlsx` |
| Data rows | **436** (rows 2–437) |
| Columns | 5 |

| # | Header | What is actually inside (by cell content) |
|---|--------|--------------------------------------------|
| A | `Company name` | 389 text, 11 "domain-looking" (`alpha.cw`, `distribuidora.beta` — Instagram handles, not websites) |
| B | `Email` | 261 e-mails, no noise |
| C | `Phone or WhatsApp number` | 314 phones, **15 links**, 100 empty |
| D | `Website` | 168 websites, **1 phone**, 1 text |
| E | `Instagram or LinkedIn` | 89 links |

---

## 2. Weak spots (before the rework)

### 2.1 Duplicates were not handled at all — **the main problem**

The old service emitted exactly as many rows as it received: 436 → 436.
Meanwhile the file contains:

* **36 groups** of rows with an identical company name (43 redundant rows),
  including a whole repeated block, rows 353–388, duplicating rows 3–38.
* More groups that are only visible through the **handle in parentheses**:
  `distribuidora_alpha_mx` (row 6) and
  `Distribuidora Alpha (distribuidora_alpha_mx)` (row 394) are one company.
* **17 e-mails** and **39 phone numbers** appear in more than one row, i.e. the
  same company is duplicated under different names
  (`Alpha Store` / `Alpha Store Colombia`,
  `BetaPro` / `Beta Pro México (betapro.mx)`).

Across all three signals: **57 groups, 81 redundant rows**, 436 → **355**.

| Group size | Groups |
|---|---|
| 2 rows | 38 |
| 3 rows | 15 |
| 4 rows | 3 |
| 5 rows | 1 |

| What matched | Groups |
|---|---|
| name + phone | 19 |
| name only | 16 |
| e-mail + name + phone | 8 |
| e-mail + phone | 4 |
| e-mail only | 3 |
| phone only | 5 |
| e-mail + name | 2 |

**Duplicates are not merely redundant — they carry data.** The repeated block
353–388 is partly impoverished: some rows lost their e-mail and phone. Merging
recovers the contacts: rows with **no contact details at all** go from
**35 to 14**, and in 5 groups the merge filled a field that was empty on the
first row.

> **Fix.** The service now always produces **two files**: the full one (with
> duplicates, one row per source row) and the clean one (merged), plus a CSV
> report saying which rows were merged into which and on what grounds. The
> merge signals (name / e-mail / phone / website) are individual checkboxes.

---

### 2.2 Misplaced values were silently dropped

Values regularly sit in the wrong column:

| Row | Column | What is there | Old behaviour |
|---|---|---|---|
| 63 | Phone | `+56911110008 / https://www.instagram.com/alpha.cl/` | phone kept, **link discarded** |
| 69 | Phone | `https://www.linkedin.com/company/alpha-group / +507 200 0000_x0003_` | phone kept, link discarded |
| 70–78 | Phone | **URL only**, no phone at all (9 rows) | cell → **empty**, data lost |
| 85 | Website | `+52 55 1111 0006` | `get_url()` could not parse it → **empty** |

In total: **15 links in the phone column** and **1 phone in the website column**.

Worse, the old logic was **irreversible**: a column detected as a phone column
was shown in the UI only as "Phone 1/2/3". The original cell text became
unreachable. Had the detection been wrong (see 2.3), the entire column would
have been lost with no way to remap it by hand.

> **Fix.** Every cell is now decomposed into "atoms" by content:
> **`@` → e-mail**, **digits/`+`/brackets only → phone**, **a dotted domain
> with a real TLD → website**. A value of the wrong type is moved to a suitable
> field (e-mail → Work/Home/Other E-mail, link → Corporate/Other Website,
> number → Mobile/Home/Other Phone); if such a value is already there it is
> dropped as a duplicate. All of it is written to `field_fixes_report.csv` with
> the row number. On the real file: **2 moves + 10 dropped duplicates**.

---

### 2.3 Phone-column detection produced false positives

The old code:

```python
PHONE_HEADER_KEYWORDS = ('phone', 'tel', 'whatsapp', 'wa', 'cel', ..., 'contact', 'number')
def looks_like_phone_header(header):
    return any(k in header.lower() for k in PHONE_HEADER_KEYWORDS)   # substring!
```

Substring matching meant these counted as phone columns:

* `Contact email` → contains `contact`
* `Warehouse` → contains `wa`
* `Hotel name` → contains `tel`

And the content check (`threshold 0.3`, "does any phone-like fragment exist")
fired on indices, dates, prices and IDs.

Combined with 2.2 this meant losing a whole column: `Contact email` would have
been split into "Phone 1/2/3" (empty) and the addresses themselves would have
become unreachable.

> **Fix.** Keywords of length ≤3 (`wa`, `tel`, `cel`) are matched **as whole
> words** rather than substrings; `contact` was removed. A stop-list was added
> (`mail`, `website`, `zip`, `id`, `date`, `price`, …), and the short stop-words
> are also matched as words so that `Provider phone` (contains `id`) is not
> vetoed. The content threshold went from 0.3 to **0.5** and now uses a verdict
> for the **whole cell** (`classify_value`) instead of "are there digits".
> Most importantly, even a wrong detection **loses nothing**: the e-mails and
> links from a phone cell are kept separately and remain available for moving.

---

### 2.4 Phones: `+` was prepended blindly

`normalize_phone()` added `+` to any number with 8+ digits:

```
3311110004  →  +3311110004      # no such country code exists
6611110009  →  +6611110009      # a Mexican number that needs +52
```

On the real file **108 numbers** were "fixed" this way — roughly **a third** of
all phone numbers ended up in an invalid international format.

> **Fix.** A **default country code** setting was added. Numbers with no `+`
> and fewer than 11 digits get the given code: `6611110009` → `+52 6611110009`.
> The column statistics now show `✓ok` / `+added` / `country code` /
> `⚠ invalid` separately, and the mapping step warns that
> "108 numbers had no + and no country code".

---

### 2.5 The 3-phones-per-cell cap truncated data silently

`MAX_PHONES_PER_CELL = 3`. The file contains a cell with **5 numbers** — 2 of
them simply disappeared with no message at all.

On top of that, the mapping suggestion only spread phones across three fields
(Mobile / Home / Other), so the 4th and 5th sub-columns stayed unmapped and
were lost too.

> **Fix.** The cap is now 5; anything above is counted (`stats.overflow`) and
> surfaced as a warning. The mapping suggestion "spills" surplus columns into
> the last field of the chain (Other Phone Number) — the values are joined with
> a comma rather than dropped.

---

### 2.6 One number landed in two fields

If the source had the same number in both "Mobile" and "Home", it was written
twice. Likewise for e-mails and websites.

> **Fix.** `_strip_cross_field_duplicates()` — after a row is assembled,
> repetitions across fields of the same type are removed, keeping the first.

---

### 2.7 Output encoding: `México` arrived as `M??xico`

The service emitted correct UTF-8, but **without a BOM**. The receiving system
did not recognise the encoding, read the two UTF-8 bytes of `é` (`C3 A9`) as
two characters of its own ANSI codepage, and a later pass through ASCII turned
each into `?` — hence **two** question marks per letter.

Telling detail: the reference `reference_import.csv`, the file that "imported
fine", contains **no non-ASCII bytes at all** — it already read
`Mediest?tica M?xico`. The receiving side has always been ASCII-only.

> **Fix.** The encoding is now selectable in the UI (`UTF-8 with BOM` — the new
> default, `UTF-8 without BOM`, `Windows-1252`, `Windows-1251`), plus a
> **“transliterate accented characters”** option (`México → Mexico`) that
> produces pure ASCII which no re-encoding can corrupt. `write_csv()` no longer
> raises on an unrepresentable character: it degrades through
> "character → without accent → `?`" and reports how much was lost.

---

### 2.8 Operational problems

| Problem | Before | After |
|---|---|---|
| In-memory sessions | grew without bound, never cleaned | TTL 6 h + 50-session cap (`SESSION_TTL_SECONDS`, `MAX_SESSIONS`) |
| `uploads/` directory | never cleaned, lead data sat forever | removed with the session + orphan-directory sweep |
| `/download/<id>` | `id` went into the path after `.replace('/','')` | strict `^[0-9a-f]{32}$` check |
| Error 413 (file > 25 MB) | Flask HTML page; the frontend crashed on `resp.json()` | JSON `{'error': ...}` |
| Huge file | read whole, no limit | `MAX_ROWS=200 000`, `MAX_COLS=300`, `TableTooLargeError` |
| Formulas in CSV | `=`, `+`, `@` at the start of a cell — formula injection when opened in Excel | optional protection (off by default, so `+` phone numbers keep working) |
| Tests | none | 40 tests, `py tests\test_converter.py` |

---

## 3. Known limitations that remain

* **Merging by phone can join different companies.** There are 5 such groups in
  the real file, and 4 of them are clearly the same company under different
  names. But rows 294 (`Gamma Insumos SRL.`) and 321
  (`Delta Medical RD`) share `+1 809-555-0100`, which may be a
  shared call centre. So: **always check `duplicates_report.csv`**, and untick
  "merge by phone" when in doubt.
* **Merging by website is off by default** — several companies can share an
  aggregator site.
* Comments and text fields are joined with `, ` when merging; for long comments
  the result can get bulky.
* The `KNOWN_TLDS` list is deliberately incomplete: otherwise Instagram handles
  such as `alpha.cw` would be treated as websites. If your file has sites in
  exotic zones, add the zone to the list (see `docs/DEVELOPER.md`).
