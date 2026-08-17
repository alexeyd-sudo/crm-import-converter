# Support guide

For people who use the service and answer questions about it.
Technical details are in [DEVELOPER.md](DEVELOPER.md).

---

## 1. What the service does

Turns an arbitrary lead spreadsheet (Excel or CSV) into a CSV for CRM import.
Along the way it:

* works out what is really inside each column (e-mail / phone / website / text);
* pulls several phone numbers out of one cell and normalises them to `+…`;
* **moves values that ended up in the wrong column** (an e-mail in the phone
  column and so on);
* **finds duplicate companies** and hands you two files — with and without them;
* writes two CSV reports so both automations can be checked.

---

## 2. Workflow

### Step 1. File

Drop in an `.xlsx`, `.xlsm` or `.csv` (up to 25 MB). **The first row of the
file must be the header row.**

### Step 2. Scanning

Runs automatically: the file is read, the columns are classified and the phone
numbers are parsed.

### Step 3. Field mapping

Source columns on the left, CRM fields on the right.

* The fields on the right are **already filled in** (highlighted with a blue
  border) from the header wording and the column content. **This is a guess —
  check it.**
* Each column carries a coloured type badge (`e-mail` / `phone` /
  `website/link` / `text`). Hover it to see how many cells of each type were
  found.
* If a cell held several phone numbers, the column is split into
  `— Phone 1`, `— Phone 2`, … which can go to different CRM fields.
* Yellow warnings appear above. Read them: they mention links found in the
  phone column, numbers without a country code and so on.
* The **`Source`** field is not listed — every row automatically gets `import`.
* Several columns may target the same field; the values are joined with `, `.

**Processing options** (collapsible block at the bottom):

| Option | Default | When to change it |
|---|---|---|
| Fix values in the wrong column | **on** | Turn off only if you need the data exactly as-is, with no moves |
| Skip rows without a company name | off | Turn on if the file has a tail of empty rows |
| Excel formula protection | off | Turn on only if the file will be opened in Excel. **Not for CRM import** — it puts a `'` in front of phone numbers |
| Default country code | empty | Set it if the file has numbers with no `+` and no country code and you know the country (`52` Mexico, `57` Colombia, `56` Chile, `54` Argentina, `598` Uruguay, `507` Panama) |
| CSV encoding | **UTF-8 with BOM** | Change only if the CRM shows `M??xico` / `MÃ©xico` — see “Common problems” |
| Transliterate accented characters | off | Turn on if the CRM breaks letters under any encoding: `México → Mexico`. Produces pure ASCII, which cannot be corrupted |
| Merge by company name | **on** | — |
| Merge by e-mail | **on** | — |
| Merge by phone | **on** | Turn off if companies may share a number (call centre, agency) |
| Merge by website | off | Turn on if every company has its own domain |

### Step 4. Result

Two files:

| File | Contents | When to take it |
|---|---|---|
| **Clean CSV** (`crm_import_merged.csv`) | duplicates merged, data from every copy collected into one row | the normal case — **take this one** |
| **CSV with duplicates** (`crm_import_with_duplicates.csv`) | one row per source row | when you need to check by hand, or when the merge does not suit you |

Two reports (also CSV, they open in Excel):

| Report | Columns | Purpose |
|---|---|---|
| `duplicates_report.csv` | group, kept name, group size, **source row number**, source name, what matched | verify that the merged companies really are the same |
| `field_fixes_report.csv` | **source row number**, column, detected type, value, where it was mapped, where it was moved | verify that the values were moved correctly |

The row numbers in the reports **match the row numbers in Excel** (row 1 is the
header).

The same data is shown on screen in collapsible blocks (first 50 entries), plus
a preview of the first 5 rows of each file.

---

## 3. Output format

* Delimiter — **semicolon `;`**
* Encoding — chosen on step 3, default **UTF-8 with BOM**
* Line endings — `CRLF`
* Quotes — only where they are needed
* 14 fixed columns:
  `Company Name; Lead Name; Mobile Phone; Home Phone; Other Phone Number;
  Corporate Website; Other Website; Work E-mail; Home E-mail; Other E-mail;
  Country OUTREACH; Outreach comment; Comment; Source`

---

## 4. Common problems

### “The whole row landed in one CRM column”

The CRM expects a different delimiter. The file uses `;`, matching the known
good `reference_import.csv`. If your system wants a comma, change `CSV_DELIMITER`
in `converter.py` (see DEVELOPER.md).

### “México arrives as `M??xico`, Estética as `Est??tica`”

**Two** question marks per letter is the signature of a double re-encoding.
Our file is UTF-8, where `é` takes 2 bytes (`C3 A9`). The receiver did not
realise the file was UTF-8, read those 2 bytes as 2 separate characters of its
own codepage, and then reduced the result to ASCII — each character became `?`.

What to do, from simplest to most reliable:

1. **Encoding = “UTF-8 with BOM”** (the default). The BOM is a marker at the
   start of the file that lets Excel and most importers detect the encoding
   automatically. If you switched to “UTF-8 without BOM”, switch back.
2. **Encoding = “Windows-1252 (Latin)”.** A single-byte encoding that does
   contain `é í ñ ó`. Use it if the importer cannot handle UTF-8 at all.
   Cyrillic will not survive — the service reports how many characters were lost.
3. **“Transliterate accented characters”.** `México → Mexico`,
   `Estética → Estetica`. The file becomes pure ASCII, so re-encoding cannot
   damage it. This is the most reliable option, and it is essentially how the
   old known-good `reference_import.csv` looked — except that it lost the letters
   to `?` instead of transliterating them.

A single `?` per letter (`M?xico`) is the same diagnosis with the intermediate
step skipped. The same fixes apply.

### “I see `Ã©` or `Г©` instead of `é`”

Half of the same illness: the receiver read UTF-8 as a single-byte encoding but
did not reduce it to ASCII. Same fix as above.

### “I open the result in Excel and see gibberish”

With “UTF-8 without BOM”, Excel does not guess the encoding on a double click.
Either switch to “UTF-8 with BOM”, or open via
*Data → From Text/CSV* → encoding UTF-8.

**The reports** (`duplicates_report.csv`, `field_fixes_report.csv`) are always
written as UTF-8 with BOM regardless of the setting — they are meant to be read
in Excel.

### “The service merged two different companies”

Most likely the phone signal fired. Look at `duplicates_report.csv` — the
*Matched on* column says what matched. Untick “merge by phone” and build the
file again (no need to re-upload if the session is still alive).

### “The service did not merge obvious duplicates”

The names differ and neither row has a phone or an e-mail, so there is nothing
to link them by. Try enabling the website signal, or take the file with
duplicates and merge by hand.

### “The column was not split into phones” / “it was split for no reason”

A column counts as a phone column when its header contains `phone`, `tel`,
`whatsapp`, `cel`, `móvil`, `телефон`… **or** when ≥50 % of its non-empty cells
look like a number. Headers containing `mail`, `website`, `url`, `id`, `zip`,
`date`, `price` are never treated as phone columns.

If the detection got it wrong, nothing is lost: the dropdown lets you send the
column to any field, and misplaced values are still found.

### “Session not found or expired”

A session lives 6 hours, or until the service restarts. Upload the file again.

### “File is larger than 25 MB”

The limit is `MAX_CONTENT_LENGTH` in `app.py`. Alternatively split the file.

### “No data found in the file”

The first row must be the header row, not a report title or a blank line.
Delete anything above the table and upload again.

---

## 5. Checking that the service is alive

```
curl http://localhost:5001/healthz
{"status":"ok","sessions":3}
```

Running the self-tests (a few seconds, harmless):

```
py tests\test_converter.py
```

Expect `40/40 passed`. If anything fails, attach the output to the ticket.

---

## 6. What to attach when escalating

1. The source file (or a fragment containing the problem row).
2. The **row number** in the source.
3. What you got and what you expected.
4. Both reports (`duplicates_report.csv`, `field_fixes_report.csv`).
5. A screenshot of step 3 (mapping + options).
6. The relevant part of the server log.
