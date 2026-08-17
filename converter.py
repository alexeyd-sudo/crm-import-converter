"""
Core parsing / cleaning / mapping logic for the CRM import converter.
No Flask here - this module is UI-agnostic so it can be tested on its own.

High level pipeline
-------------------
    read_table()          file  -> headers, rows, hyperlinks
    analyze_columns()     rows  -> per-column content statistics ("what is
                                   actually inside this column")
    detect_phone_columns()/scan_phones()
                          rows  -> phone columns exploded into sub-columns
    build_source_columns()      -> the list shown on the left of the mapping UI
    suggest_mapping()           -> pre-selected targets for that list
    build_records()       mapping -> output records + "misplaced value" report
    dedupe_records()      records -> merged records + duplicate report
    write_csv()

Two ideas drive most of the code:

1. *Atoms.* Every source cell is decomposed into typed atoms (emails, urls,
   phones, plain text) instead of being trusted to hold what its column
   header promises. An e-mail dropped into a phone column is still found,
   and is routed to an e-mail field instead of being thrown away.

2. *Duplicates are a view, not a decision.* The converter always produces
   BOTH the raw row-per-source-row table and the merged/deduplicated one,
   plus a report describing exactly which rows were merged, so a human can
   audit the merge instead of trusting it blindly.
"""
import csv
import io
import re
import unicodedata
from urllib.parse import urlparse, parse_qs, unquote, urlunparse

import openpyxl

# ---------------------------------------------------------------------------
# Target CRM fields shown on the "right side" of the mapping screen.
# ---------------------------------------------------------------------------

TARGET_FIELDS = [
    {'id': 'company_lead', 'label': 'Company Name / Lead Name', 'type': 'text',
     'outputs': ['Company Name', 'Lead Name'],
     'hint': 'The value is written to 2 result columns at once: Company Name and Lead Name'},
    {'id': 'mobile_phone', 'label': 'Mobile Phone', 'type': 'phone', 'outputs': ['Mobile Phone'], 'hint': ''},
    {'id': 'home_phone', 'label': 'Home Phone', 'type': 'phone', 'outputs': ['Home Phone'], 'hint': ''},
    {'id': 'other_phone', 'label': 'Other Phone Number', 'type': 'phone', 'outputs': ['Other Phone Number'], 'hint': ''},
    {'id': 'corporate_website', 'label': 'Corporate Website', 'type': 'url', 'outputs': ['Corporate Website'], 'hint': ''},
    {'id': 'other_website', 'label': 'Other Website', 'type': 'url', 'outputs': ['Other Website'], 'hint': ''},
    {'id': 'work_email', 'label': 'Work E-mail', 'type': 'email', 'outputs': ['Work E-mail'], 'hint': ''},
    {'id': 'home_email', 'label': 'Home E-mail', 'type': 'email', 'outputs': ['Home E-mail'], 'hint': ''},
    {'id': 'other_email', 'label': 'Other E-mail', 'type': 'email', 'outputs': ['Other E-mail'], 'hint': ''},
    {'id': 'country_outreach', 'label': 'Country OUTREACH', 'type': 'text', 'outputs': ['Country OUTREACH'], 'hint': ''},
    {'id': 'outreach_comment', 'label': 'Outreach comment', 'type': 'text', 'outputs': ['Outreach comment'], 'hint': ''},
    {'id': 'comment', 'label': 'Comment', 'type': 'text', 'outputs': ['Comment'], 'hint': ''},
]
TARGET_BY_ID = {t['id']: t for t in TARGET_FIELDS}

OUTPUT_HEADERS = [
    'Company Name', 'Lead Name',
    'Mobile Phone', 'Home Phone', 'Other Phone Number',
    'Corporate Website', 'Other Website',
    'Work E-mail', 'Home E-mail', 'Other E-mail',
    'Country OUTREACH', 'Outreach comment', 'Comment',
    'Source',
]

# Where a value of a given type is allowed to land, in priority order.
# Used both by the auto-mapping suggestions and by the "misplaced value"
# relocation logic.
TYPE_CHAINS = {
    'email': ['work_email', 'home_email', 'other_email'],
    'phone': ['mobile_phone', 'home_phone', 'other_phone'],
    'url': ['corporate_website', 'other_website'],
}

PHONE_OUTPUTS = ['Mobile Phone', 'Home Phone', 'Other Phone Number']
EMAIL_OUTPUTS = ['Work E-mail', 'Home E-mail', 'Other E-mail']
URL_OUTPUTS = ['Corporate Website', 'Other Website']

SOURCE_VALUE = 'import'

# Max phone numbers pulled out of a single cell. Anything above this is
# counted in stats['overflow'] and surfaced in the UI as a warning rather
# than being dropped silently.
MAX_PHONES_PER_CELL = 5

# CSV dialect that matches the known-good reference_import.csv reference file:
# ';' delimiter, CRLF line endings, quote only when actually needed.
# (The legacy script used ',' + quote-all, which is why imports looked
# "crooked" - most locales/Bitrix read ';' as the field separator.)
CSV_DELIMITER = ';'
CSV_LINETERMINATOR = '\r\n'

# Default is UTF-8 *with* BOM. The BOM is the only in-band signal that tells
# Excel and most CRM importers "this is UTF-8". Without it they fall back to
# the system ANSI codepage, read the two UTF-8 bytes of "é" as two separate
# characters, and a later ASCII conversion turns each into '?' - which is
# where "México" -> "M??xico" (two question marks per letter) comes from.
CSV_ENCODING = 'utf-8-sig'

CSV_ENCODINGS = [
    {'id': 'utf-8-sig', 'label': 'UTF-8 with BOM (recommended)',
     'hint': 'Excel and most CRM importers detect the encoding automatically.'},
    {'id': 'utf-8', 'label': 'UTF-8 without BOM',
     'hint': 'Use if your importer trips over the BOM in the first header column.'},
    {'id': 'cp1252', 'label': 'Windows-1252 (Latin)',
     'hint': 'Use if the target system insists on a single-byte encoding. '
             'Spanish é, í, ñ survive; Cyrillic does not.'},
    {'id': 'cp1251', 'label': 'Windows-1251 (Cyrillic)',
     'hint': 'Only if the data contains Russian text and the importer expects ANSI.'},
]
CSV_ENCODING_IDS = {e['id'] for e in CSV_ENCODINGS}


# ---------------------------------------------------------------------------
# generic text cleaning
# ---------------------------------------------------------------------------

def clean_text(val):
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    val = unicodedata.normalize('NFKC', str(val)).strip()
    val = val.replace('\n', ' ').strip()
    val = re.sub(r'\s{2,}', ' ', val)
    return val


# ---------------------------------------------------------------------------
# email
# ---------------------------------------------------------------------------

EMAIL_RE = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
# "user@gmail,com" -> "user@gmail.com" (comma typed instead of dot)
EMAIL_TYPO_RE = re.compile(r'(@[\w\-]+),([a-zA-Z]{2,4})\b')


def fix_email_typos(s):
    return EMAIL_TYPO_RE.sub(r'\1.\2', s)


def extract_emails(val):
    if not val:
        return []
    s = unicodedata.normalize('NFKC', str(val)).strip()
    s = fix_email_typos(s)
    s = s.replace('mailto:', ' ')
    found = EMAIL_RE.findall(s)
    seen = set()
    out = []
    for e in found:
        key = e.lower()
        if key not in seen:
            seen.add(key)
            out.append(e)
    return out


# ---------------------------------------------------------------------------
# phone
# ---------------------------------------------------------------------------

URL_RE = re.compile(r'https?://\S+')
WAME_RE = re.compile(r'wa\.me/\S+', re.IGNORECASE)
CTRL_ARTIFACT_RE = re.compile(r'(?:_x00[0-9A-Fa-f]{2}_)+')

# Word-ish keywords. Matched with a boundary check (see looks_like_phone_header)
# so that "wa" does not fire on "Warehouse" and "tel" does not fire on "Hotel".
PHONE_HEADER_KEYWORDS = (
    'phone', 'phones', 'tel', 'telefono', 'teléfono', 'telefone', 'whatsapp',
    'wa', 'wsp', 'cel', 'celular', 'movil', 'móvil', 'mobile', 'number',
    'номер', 'телефон', 'моб', 'тел',
)
# Headers that must never be treated as phone columns even if their content
# happens to contain long digit runs. Long, unambiguous fragments are matched
# as substrings; short ones only as whole words, so that "Provider phone"
# (contains "id") is not vetoed.
NON_PHONE_HEADER_SUBSTRINGS = (
    'mail', 'correo', 'website', 'instagram', 'linkedin', 'facebook',
    'postal', 'fecha', 'дата', 'price', 'precio', 'цена', 'сайт', 'почт', 'ссыл',
)
NON_PHONE_HEADER_WORDS = {
    'web', 'site', 'url', 'link', 'id', 'zip', 'code', 'codigo', 'código', 'date',
}


def _split_parens(segment):
    extras = []

    def repl(m):
        content = m.group(1)
        if len(re.sub(r'\D', '', content)) >= 7:
            extras.append(content)
            return ' '
        return ' ' + content + ' '

    main = re.sub(r'\(([^)]*)\)', repl, segment)
    return main, extras


def _clean_phone_candidate(c):
    c = c.replace('(', '').replace(')', '')
    c = re.sub(r'^[^\d+]+', '', c)
    c = re.sub(r'[^\d\s+\-]+$', '', c)
    c = re.sub(r'\s{2,}', ' ', c).strip()
    if len(re.sub(r'\D', '', c)) < 5:
        return ''
    return c


def extract_phone_candidates(val):
    """Pull out every phone-like substring from a cell, cleaned but not yet
    '+'-normalized. Order preserved, de-duplicated by digit signature."""
    if not val:
        return []
    s = unicodedata.normalize('NFKC', str(val))
    s = CTRL_ARTIFACT_RE.sub('', s)

    phones = []
    for m in WAME_RE.findall(s):
        num = m.split('/', 1)[1] if '/' in m else ''
        num = re.sub(r'[^\d+]', '', num)
        if num:
            if not num.startswith('+'):
                num = '+' + num
            phones.append(num)

    s = WAME_RE.sub(' ', s)
    s = URL_RE.sub(' ', s)
    # an e-mail's local part can contain digits ("info2024@x.com") - drop
    # e-mails before looking for numbers
    s = EMAIL_RE.sub(' ', fix_email_typos(s))

    for part in re.split(r'[/&,;]', s):
        part = part.strip()
        if not part:
            continue
        main, extras = _split_parens(part)
        for cand in [main] + extras:
            cleaned = _clean_phone_candidate(cand)
            if cleaned:
                phones.append(cleaned)

    seen = set()
    out = []
    for p in phones:
        sig = re.sub(r'\D', '', p)
        if sig and sig not in seen:
            seen.add(sig)
            out.append(p)
    return out


def phone_signature(p):
    """Digits only - used for duplicate comparison."""
    return re.sub(r'\D', '', p or '')


def normalize_phone(raw, default_cc=None):
    """Ensure a phone string starts with '+'.

    Returns (value, status):
      'ok'       - already had a leading '+' (or 00 international prefix)
      'fixed'    - we added '+' to a number that already looked international
      'cc_added' - we prepended the operator-supplied default country code
                   because the number looked like a national number
      'invalid'  - too few digits to trust; kept as-is and flagged

    Note: 'fixed' is a *guess*. A bare 10-digit national number with no
    country code is ambiguous, which is why `default_cc` exists.
    """
    raw = (raw or '').strip()
    digits = re.sub(r'\D', '', raw)
    if not digits:
        return raw, 'invalid'
    if raw.startswith('+'):
        return raw, 'ok'
    if raw.startswith('00'):
        return '+' + raw[2:].strip(), 'ok'

    cc = re.sub(r'\D', '', default_cc or '')
    if cc and len(digits) <= 10 and not digits.startswith(cc):
        return '+' + cc + ' ' + raw, 'cc_added'
    if len(digits) >= 8:
        return '+' + raw, 'fixed'
    return raw, 'invalid'


def _header_words(header):
    return set(re.split(r'[^0-9a-zа-яё]+', (header or '').lower()))


def looks_like_non_phone_header(header):
    h = (header or '').lower()
    if any(k in h for k in NON_PHONE_HEADER_SUBSTRINGS):
        return True
    return bool(_header_words(h) & NON_PHONE_HEADER_WORDS)


def looks_like_phone_header(header):
    h = (header or '').lower()
    if looks_like_non_phone_header(h):
        return False
    words = _header_words(h)
    if words & set(PHONE_HEADER_KEYWORDS):
        return True
    # allow "phone"/"telefono" as substrings of longer words, but never the
    # short ambiguous ones ('wa', 'tel', 'cel', ...)
    return any(k in h for k in ('phone', 'telefon', 'telefono', 'telefone', 'whatsapp', 'телефон'))


# ---------------------------------------------------------------------------
# urls (website)
# ---------------------------------------------------------------------------

TRACKING_PARAMS = {
    'utm_source', 'utm_medium', 'utm_content', 'utm_campaign',
    'fbclid', 'igshid', 'app_absent', 'text', 'brid', 'h', 'igsh',
    'utm_term', 'utm_id',
}

# Conservative TLD list: only used to decide whether a bare, scheme-less token
# such as "distribuidora-alpha.com.mx" is a website. Deliberately does NOT contain
# exotic ccTLDs, because Instagram handles in this dataset look like
# "alpha.cw" / "distribuidora.beta" and must stay plain text.
KNOWN_TLDS = {
    'com', 'net', 'org', 'info', 'biz', 'io', 'co', 'ai', 'app', 'dev', 'me',
    'shop', 'store', 'online', 'site', 'clinic', 'health', 'care', 'agency',
    'us', 'uk', 'eu', 'es', 'pt', 'fr', 'de', 'it', 'nl', 'pl', 'ru', 'ua',
    'mx', 'cl', 'ar', 'br', 'uy', 'py', 'bo', 'pe', 'ec', 've', 'pa', 'cr',
    'gt', 'hn', 'ni', 'sv', 'do', 'pr', 'cu', 'ca', 'au', 'nz', 'in', 'cn',
    'jp', 'kr', 'tr', 'il', 'ae', 'za',
}

URL_FULL_RE = re.compile(r'(?:https?://|www\.)[^\s,;<>"\']+', re.IGNORECASE)
BARE_DOMAIN_RE = re.compile(
    r'\b(?:[a-z0-9](?:[a-z0-9\-]{0,61}[a-z0-9])?\.)+([a-z]{2,10})\b(?:/[^\s,;<>"\']*)?',
    re.IGNORECASE,
)

# Hosts where the path identifies the account, so the path must be part of
# any duplicate key (otherwise every Instagram lead would merge into one).
SOCIAL_HOSTS = {
    'instagram.com', 'linkedin.com', 'facebook.com', 'fb.com', 'wa.me',
    't.me', 'twitter.com', 'x.com', 'tiktok.com', 'youtube.com',
}

# Links that are really a phone number in disguise. extract_phone_candidates
# turns them into numbers, so extract_urls must not also report them as
# websites (otherwise every "wa.me/5219..." lands in Corporate Website).
PHONE_LINK_HOSTS = {'wa.me', 'api.whatsapp.com', 'web.whatsapp.com'}


def extract_real_url(url):
    if not url:
        return ''
    url = url.strip()
    if url.startswith('mailto:'):
        return ''

    if 'l.facebook.com/l.php' in url:
        try:
            parsed = urlparse(url)
            params = parse_qs(parsed.query)
            if 'u' in params:
                return extract_real_url(params['u'][0])
        except Exception:
            pass

    if 'l.instagram.com' in url:
        try:
            parsed = urlparse(url)
            params = parse_qs(parsed.query)
            if 'u' in params:
                return extract_real_url(unquote(params['u'][0]))
        except Exception:
            pass

    try:
        parsed = urlparse(url)
        if parsed.query:
            params = parse_qs(parsed.query, keep_blank_values=True)
            cleaned = {k: v for k, v in params.items() if k not in TRACKING_PARAMS}
            new_query = '&'.join(f'{k}={v[0]}' for k, v in cleaned.items()) if cleaned else ''
            url = urlunparse(parsed._replace(query=new_query))
    except Exception:
        pass

    return url


def get_url(val, hyp_target=None):
    """Legacy single-value helper: treat the whole cell as one URL.
    Kept because it is the most permissive interpretation and is still used
    as a fallback for url-typed targets."""
    if hyp_target and hyp_target.startswith('mailto:'):
        hyp_target = None

    val_str = unicodedata.normalize('NFKC', str(val)).strip() if val is not None else ''

    url_to_use = None
    if val_str.startswith('http://') or val_str.startswith('https://'):
        url_to_use = val_str
    elif val_str and '.' in val_str and ' ' not in val_str and '@' not in val_str and len(val_str) > 4:
        url_to_use = hyp_target if hyp_target else ('https://' + val_str)
    elif hyp_target:
        url_to_use = hyp_target

    if not url_to_use:
        return ''

    return extract_real_url(url_to_use)


def _tidy_url(u):
    u = u.strip().rstrip('.,;:)')
    if u.lower().startswith('www.'):
        u = 'https://' + u
    elif not re.match(r'^https?://', u, re.IGNORECASE):
        u = 'https://' + u
    return extract_real_url(u)


def extract_urls(val, hyp_target=None):
    """Find every website-ish URL in a cell.

    Returns (confident, loose):
      confident - has a scheme / www. prefix / a recognised TLD. Safe to move
                  into a website field even if it came from another column.
      loose     - the whole-cell fallback (`get_url`) when nothing confident
                  was found. Used only when the target IS a website field.
    """
    s = unicodedata.normalize('NFKC', str(val)).strip() if val is not None else ''
    s = CTRL_ARTIFACT_RE.sub('', s)
    s = EMAIL_RE.sub(' ', fix_email_typos(s))

    found = []
    rest = s
    for m in URL_FULL_RE.findall(s):
        found.append(_tidy_url(m))
    rest = URL_FULL_RE.sub(' ', rest)

    for m in BARE_DOMAIN_RE.finditer(rest):
        if m.group(1).lower() in KNOWN_TLDS:
            found.append(_tidy_url(m.group(0)))

    if hyp_target and not hyp_target.startswith('mailto:'):
        found.append(extract_real_url(hyp_target))

    seen = set()
    confident = []
    for u in found:
        if not u:
            continue
        key = url_key(u)
        if not key or key.split('/')[0] in PHONE_LINK_HOSTS:
            continue  # it is a phone link - extract_phone_candidates owns it
        if key not in seen:
            seen.add(key)
            confident.append(u)

    loose = []
    if not confident:
        u = get_url(val, hyp_target)
        if u and url_key(u).split('/')[0] not in PHONE_LINK_HOSTS:
            loose.append(u)
    return confident, loose


def url_key(u):
    """Normalized key used for de-duplication and duplicate matching."""
    if not u:
        return ''
    try:
        p = urlparse(u if re.match(r'^https?://', u, re.I) else 'https://' + u)
    except Exception:
        return u.strip().lower()
    host = (p.netloc or '').lower().split(':')[0]
    if host.startswith('www.'):
        host = host[4:]
    path = (p.path or '').rstrip('/').lower()
    if host in SOCIAL_HOSTS:
        return host + path
    return host


def website_dup_key(u):
    """Key used when merging duplicates by website. Social links keep their
    path (the account), plain websites collapse to the bare domain."""
    key = url_key(u)
    if not key:
        return ''
    host = key.split('/')[0]
    if host in SOCIAL_HOSTS:
        # instagram.com/foo, linkedin.com/company/foo -> keep 2 segments max
        parts = key.split('/')
        return '/'.join(parts[:3]) if len(parts) > 2 else key
    return host


# ---------------------------------------------------------------------------
# value classification
# ---------------------------------------------------------------------------

def classify_value(val):
    """Best-effort type of a whole cell: 'email' | 'phone' | 'url' | 'text'
    | '' (empty). Used for column statistics and mapping suggestions."""
    s = clean_text(val)
    if not s:
        return ''
    if EMAIL_RE.search(fix_email_typos(s)):
        return 'email'
    confident, _loose = extract_urls(s)
    if confident:
        return 'url'
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 7 and len(re.sub(r'[\d\s+()\-.,/xXeEtT]', '', s)) <= 2:
        return 'phone'
    if extract_phone_candidates(s):
        return 'phone'
    return 'text'


def extract_atoms(val, hyp_target=None, default_cc=None):
    """Decompose one raw cell into typed atoms.

    Returns dict:
      emails      list[str]
      urls        list[str]   confident websites
      urls_loose  list[str]   whole-cell fallback (see extract_urls)
      phones      list[str]   normalized, '+'-prefixed
      phone_status list[str]  status per phone, aligned with `phones`
      overflow    int         phones found beyond MAX_PHONES_PER_CELL
      text        str         cleaned whole cell
    """
    text = clean_text(val)
    emails = extract_emails(text)
    urls, urls_loose = extract_urls(text, hyp_target)
    cands = extract_phone_candidates(text)
    overflow = max(0, len(cands) - MAX_PHONES_PER_CELL)
    phones, statuses = [], []
    for c in cands[:MAX_PHONES_PER_CELL]:
        v, st = normalize_phone(c, default_cc)
        phones.append(v)
        statuses.append(st)
    return {
        'emails': emails,
        'urls': urls,
        'urls_loose': urls_loose,
        'phones': phones,
        'phone_status': statuses,
        'overflow': overflow,
        'text': text,
    }


# ---------------------------------------------------------------------------
# dedupe helpers used when several source columns map to the same target
# ---------------------------------------------------------------------------

def dedupe_join(values):
    seen = set()
    out = []
    for v in values:
        v = (v or '').strip()
        key = v.lower()
        if v and key not in seen:
            seen.add(key)
            out.append(v)
    return ', '.join(out)


def dedupe_join_urls(values):
    seen = set()
    out = []
    for v in values:
        v = (v or '').strip()
        if not v:
            continue
        key = url_key(v)
        if key not in seen:
            seen.add(key)
            out.append(v)
    return ', '.join(out)


def _merge_phone_list(values):
    """Return list of surviving phone strings, collapsing 'same number with /
    without country code' pairs."""
    kept = []  # [digits, original]
    for p in values:
        p = (p or '').strip()
        if not p:
            continue
        digits = phone_signature(p)
        if not digits:
            continue
        dup_idx = None
        for idx, (kd, _kp) in enumerate(kept):
            if kd == digits:
                dup_idx = idx
                break
            if kd.endswith(digits) and len(kd) - len(digits) <= 3:
                dup_idx = idx
                break
            if digits.endswith(kd) and len(digits) - len(kd) <= 3:
                kept[idx] = [digits, p]  # prefer the more complete version
                dup_idx = idx
                break
        if dup_idx is None:
            kept.append([digits, p])
    return [p for _, p in kept]


def dedupe_join_phones(values):
    return ', '.join(_merge_phone_list(values))


# ---------------------------------------------------------------------------
# reading the uploaded table
# ---------------------------------------------------------------------------

class TableTooLargeError(Exception):
    pass


MAX_ROWS = 200_000
MAX_COLS = 300


def read_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    max_col = min(ws.max_column or 0, MAX_COLS)
    if (ws.max_row or 0) > MAX_ROWS:
        raise TableTooLargeError(f'Too many rows: {ws.max_row} (limit {MAX_ROWS}).')
    headers = []
    for c in range(1, max_col + 1):
        h = clean_text(ws.cell(1, c).value)
        headers.append(h or f'Column {c}')

    rows = []
    hyperlinks = []
    for r in range(2, (ws.max_row or 1) + 1):
        raw_vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        if all(v is None or str(v).strip() == '' for v in raw_vals):
            continue
        rows.append([clean_text(v) for v in raw_vals])
        hyperlinks.append([
            (ws.cell(r, c).hyperlink.target if ws.cell(r, c).hyperlink else None)
            for c in range(1, max_col + 1)
        ])
    return headers, rows, hyperlinks


def read_csv_bytes(raw_bytes):
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1251', 'latin-1'):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw_bytes.decode('utf-8', errors='replace')

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=';,\t')
        delim = dialect.delimiter
    except Exception:
        delim = ';' if sample.count(';') >= sample.count(',') else ','

    reader = csv.reader(io.StringIO(text), delimiter=delim)
    data = [row for row in reader if any(cell.strip() for cell in row)]
    if not data:
        return [], [], []
    if len(data) > MAX_ROWS:
        raise TableTooLargeError(f'Too many rows: {len(data)} (limit {MAX_ROWS}).')
    headers = [clean_text(h) or f'Column {i + 1}' for i, h in enumerate(data[0][:MAX_COLS])]
    width = len(headers)
    rows = []
    for r in data[1:]:
        r = list(r) + [''] * (width - len(r))
        rows.append([clean_text(v) for v in r[:width]])
    hyperlinks = [[None] * width for _ in rows]
    return headers, rows, hyperlinks


def read_table(path, filename):
    lower = filename.lower()
    if lower.endswith('.xlsx') or lower.endswith('.xlsm'):
        return read_xlsx(path)
    if lower.endswith('.csv') or lower.endswith('.txt'):
        with open(path, 'rb') as f:
            raw = f.read()
        return read_csv_bytes(raw)
    # fall back: try xlsx first, then csv
    try:
        return read_xlsx(path)
    except TableTooLargeError:
        raise
    except Exception:
        with open(path, 'rb') as f:
            raw = f.read()
        return read_csv_bytes(raw)


# ---------------------------------------------------------------------------
# column content analysis
# ---------------------------------------------------------------------------

HEADER_TARGET_HINTS = [
    (('company', 'empresa', 'compania', 'compañia', 'compañía', 'razon', 'razón',
      'name', 'nombre', 'lead', 'клиент', 'компан', 'название', 'имя'), 'company_lead'),
    (('mail', 'correo', 'почт'), 'email'),
    (('instagram', 'linkedin', 'facebook', 'social', 'red', 'соц'), 'social'),
    (('website', 'web', 'site', 'url', 'сайт', 'ссыл'), 'url'),
    (('phone', 'tel', 'whatsapp', 'wsp', 'cel', 'movil', 'móvil', 'mobile',
      'телефон', 'номер'), 'phone'),
    (('country', 'pais', 'país', 'страна'), 'country_outreach'),
    (('outreach',), 'outreach_comment'),
    (('comment', 'coment', 'note', 'nota', 'коммент', 'замет'), 'comment'),
]


def header_hint(header):
    h = (header or '').lower()
    for keywords, hint in HEADER_TARGET_HINTS:
        if any(k in h for k in keywords):
            return hint
    return ''


def analyze_columns(headers, rows, sample_size=400):
    """Per-column content statistics: how many non-empty cells look like an
    e-mail / a phone / a website / plain text, plus the dominant kind."""
    sample = rows[:sample_size]
    out = []
    for c, header in enumerate(headers):
        counts = {'email': 0, 'phone': 0, 'url': 0, 'text': 0}
        nonempty = 0
        for row in sample:
            cell = row[c] if c < len(row) else ''
            if not cell:
                continue
            nonempty += 1
            k = classify_value(cell)
            if k:
                counts[k] += 1
        dominant, share = '', 0.0
        if nonempty:
            dominant = max(counts, key=lambda k: counts[k])
            share = counts[dominant] / nonempty
            if counts[dominant] == 0:
                dominant, share = '', 0.0
        out.append({
            'index': c,
            'header': header,
            'nonempty': nonempty,
            'sampled': len(sample),
            'counts': counts,
            'dominant': dominant,
            'share': round(share, 3),
            'header_hint': header_hint(header),
        })
    return out


# ---------------------------------------------------------------------------
# phone-column detection + splitting ("scanning" step)
# ---------------------------------------------------------------------------

def detect_phone_columns(headers, rows, sample_size=400, col_stats=None):
    """A column is split into phone sub-columns when its header says "phone"
    OR its content is dominated by phone-shaped values.

    Content detection uses classify_value (whole-cell verdict) rather than
    "does any phone-like substring exist", so an e-mail/URL column is no
    longer mistaken for a phone column.
    """
    stats = col_stats if col_stats is not None else analyze_columns(headers, rows, sample_size)
    phone_cols = []
    for c, st in enumerate(stats):
        if looks_like_non_phone_header(headers[c]):
            continue
        header_hit = looks_like_phone_header(headers[c])
        nonempty = st['nonempty']
        content_hit = nonempty >= 1 and (st['counts']['phone'] / nonempty) >= 0.5
        if header_hit or content_hit:
            phone_cols.append(c)
    return phone_cols


def scan_phones(headers, rows, phone_col_indices, default_cc=None):
    """For every detected phone column, extract up to MAX_PHONES_PER_CELL
    normalized numbers per row, plus the non-phone leftovers (e-mails and
    websites accidentally typed into the phone column).

    Returns:
      phone_data: {col_idx: {'per_row': [[num, ...], ...],
                             'leftovers': [{'emails': [...], 'urls': [...]}, ...],
                             'max_count': int,
                             'stats': {'ok','fixed','cc_added','invalid','overflow',
                                       'foreign_emails','foreign_urls','empty'}}}
    """
    phone_data = {}
    for c in phone_col_indices:
        per_row = []
        leftovers = []
        stats = {'ok': 0, 'fixed': 0, 'cc_added': 0, 'invalid': 0,
                 'overflow': 0, 'foreign_emails': 0, 'foreign_urls': 0, 'empty': 0}
        max_count = 1
        for row in rows:
            cell = row[c] if c < len(row) else ''
            atoms = extract_atoms(cell, None, default_cc)
            for st in atoms['phone_status']:
                stats[st] += 1
            stats['overflow'] += atoms['overflow']
            stats['foreign_emails'] += len(atoms['emails'])
            stats['foreign_urls'] += len(atoms['urls'])
            if cell and not atoms['phones']:
                stats['empty'] += 1
            per_row.append(atoms['phones'])
            leftovers.append({'emails': atoms['emails'], 'urls': atoms['urls']})
            max_count = max(max_count, len(atoms['phones']))
        phone_data[c] = {
            'per_row': per_row,
            'leftovers': leftovers,
            'max_count': min(max_count, MAX_PHONES_PER_CELL),
            'stats': stats,
        }
    return phone_data


def build_source_columns(headers, rows, phone_data, col_stats=None):
    """Build the flat list of mappable "source columns" shown in the UI:
    plain columns as-is, phone columns exploded into sub-columns."""
    stats = col_stats if col_stats is not None else analyze_columns(headers, rows)
    columns = []
    for c, header in enumerate(headers):
        st = stats[c]
        if c in phone_data:
            info = phone_data[c]
            n = info['max_count']
            for sub in range(n):
                key = f'{c}:{sub}'
                label = header if n == 1 else f'{header} — Phone {sub + 1}'
                samples = []
                for pr in info['per_row']:
                    if sub < len(pr) and pr[sub]:
                        samples.append(pr[sub])
                    if len(samples) >= 3:
                        break
                columns.append({
                    'key': key,
                    'label': label,
                    'kind': 'phone',
                    'samples': samples,
                    'stats': info['stats'] if sub == 0 else None,
                    'analysis': st if sub == 0 else None,
                })
        else:
            samples = []
            for row in rows:
                v = row[c] if c < len(row) else ''
                if v:
                    samples.append(v)
                if len(samples) >= 3:
                    break
            columns.append({
                'key': str(c),
                'label': header,
                'kind': st['dominant'] or 'text',
                'samples': samples,
                'stats': None,
                'analysis': st,
            })
    return columns


def suggest_mapping(source_columns):
    """Pre-fill the mapping dropdowns. Header wording wins; content type is
    the tie-breaker. Each concrete target is used at most once, spilling over
    to the next slot in its TYPE_CHAIN (Work -> Home -> Other e-mail, ...)."""
    used = set()
    suggestion = {}

    def take(chain, spill=True):
        """First unused slot of the chain. When the chain is exhausted and
        `spill` is on, reuse the last slot - two source columns mapped to the
        same target are joined with ', ', which is far better than silently
        dropping the 4th phone/e-mail column."""
        for t in chain:
            if t not in used:
                used.add(t)
                return t
        return chain[-1] if (spill and chain) else ''

    for col in source_columns:
        analysis = col.get('analysis') or {}
        hint = analysis.get('header_hint', '')
        kind = col.get('kind') or analysis.get('dominant') or ''
        target = ''
        if hint == 'company_lead':
            target = take(['company_lead'], spill=False)
        elif hint == 'email':
            target = take(TYPE_CHAINS['email'])
        elif hint == 'social':
            target = take(['other_website', 'corporate_website'])
        elif hint == 'url':
            target = take(TYPE_CHAINS['url'])
        elif hint == 'phone':
            target = take(TYPE_CHAINS['phone'])
        elif hint in ('country_outreach', 'outreach_comment', 'comment'):
            target = take([hint], spill=False)
        elif kind in TYPE_CHAINS:
            target = take(TYPE_CHAINS[kind])
        if target:
            suggestion[col['key']] = target
    return suggestion


# ---------------------------------------------------------------------------
# building the final output
# ---------------------------------------------------------------------------

DEFAULT_OPTIONS = {
    'autofix_types': True,       # move misplaced e-mails/phones/urls
    'default_country_code': '',  # e.g. '52' - used by normalize_phone
    'source_value': SOURCE_VALUE,
    'skip_rows_without_name': False,
    'sanitize_formulas': False,  # prefix =,+,-,@ with ' (Excel formula injection)
    'csv_encoding': CSV_ENCODING,
    'transliterate': False,      # é -> e, ñ -> n: guarantees pure ASCII output
    'dedupe_by_name': True,
    'dedupe_by_email': True,
    'dedupe_by_phone': True,
    'dedupe_by_website': False,
}


def normalize_options(opts):
    o = dict(DEFAULT_OPTIONS)
    for k, v in (opts or {}).items():
        if k in o:
            o[k] = v
    o['default_country_code'] = re.sub(r'\D', '', str(o['default_country_code'] or ''))
    if o['csv_encoding'] not in CSV_ENCODING_IDS:
        o['csv_encoding'] = CSV_ENCODING
    return o


def _atoms_for_key(key, row_idx, rows, hyperlinks, phone_data, default_cc):
    """Resolve one source key on one row into atoms.

    For an exploded phone sub-column the phone value is already normalized by
    scan_phones; the non-phone leftovers of that cell are attached to sub 0
    so that they can still be recovered / relocated.
    """
    if ':' in key:
        col_str, sub_str = key.split(':')
        c, sub = int(col_str), int(sub_str)
        info = phone_data.get(c)
        if not info:
            return None
        per_row = info['per_row']
        vals = per_row[row_idx] if row_idx < len(per_row) else []
        value = vals[sub] if sub < len(vals) else ''
        left = info['leftovers'][row_idx] if row_idx < len(info['leftovers']) else {}
        raw_cell = rows[row_idx][c] if row_idx < len(rows) and c < len(rows[row_idx]) else ''
        return {
            'emails': left.get('emails', []) if sub == 0 else [],
            'urls': left.get('urls', []) if sub == 0 else [],
            'urls_loose': [],
            'phones': [value] if value else [],
            'phone_status': [],
            'overflow': 0,
            'text': value or (raw_cell if sub == 0 else ''),
        }
    c = int(key)
    row = rows[row_idx] if row_idx < len(rows) else []
    val = row[c] if c < len(row) else ''
    hyp = None
    if hyperlinks and row_idx < len(hyperlinks):
        hrow = hyperlinks[row_idx]
        hyp = hrow[c] if c < len(hrow) else None
    return extract_atoms(val, hyp, default_cc)


def _dup_key(kind, value):
    if kind == 'phone':
        return phone_signature(value)
    if kind == 'email':
        return (value or '').strip().lower()
    if kind == 'url':
        return url_key(value)
    return (value or '').strip().lower()


def build_records(headers, rows, hyperlinks, phone_data, mapping, options=None,
                  column_labels=None):
    """mapping: {source_key: target_id}.

    Returns (records, fixes):
      records - list of dicts keyed by OUTPUT_HEADERS, plus '_row' (1-based
                source row number, header row excluded) used by the reports.
      fixes   - list of dicts describing every value that was found in the
                "wrong" column, and what happened to it.
    """
    opts = normalize_options(options)
    labels = column_labels or {}
    autofix = bool(opts['autofix_types'])
    default_cc = opts['default_country_code']

    by_target = {}
    for key, target_id in (mapping or {}).items():
        if not target_id or target_id not in TARGET_BY_ID:
            continue
        by_target.setdefault(target_id, []).append(key)
    # stable, UI order
    for t in by_target:
        by_target[t].sort(key=lambda k: (int(k.split(':')[0]), int(k.split(':')[1]) if ':' in k else -1))

    records = []
    fixes = []

    for row_idx in range(len(rows)):
        buckets = {t: [] for t in by_target}
        pending = []  # (kind, value, from_target, source_label)

        for target_id, keys in by_target.items():
            ttype = TARGET_BY_ID[target_id]['type']
            for key in keys:
                atoms = _atoms_for_key(key, row_idx, rows, hyperlinks, phone_data, default_cc)
                if atoms is None:
                    continue
                label = labels.get(key, key)

                if ttype == 'text':
                    if atoms['text']:
                        buckets[target_id].append(atoms['text'])
                    continue

                own = {'email': atoms['emails'],
                       'phone': atoms['phones'],
                       'url': atoms['urls'] + atoms['urls_loose']}[ttype]
                buckets[target_id].extend(own)

                if autofix:
                    for other in ('email', 'phone', 'url'):
                        if other == ttype:
                            continue
                        for v in {'email': atoms['emails'],
                                  'phone': atoms['phones'],
                                  'url': atoms['urls']}[other]:
                            pending.append((other, v, target_id, label))

        # ---- place relocated values -------------------------------------
        for kind, value, from_target, label in pending:
            chain = TYPE_CHAINS[kind]
            existing = set()
            for t in chain:
                for v in buckets.get(t, []):
                    existing.add(_dup_key(kind, v))
            if _dup_key(kind, value) in existing:
                fixes.append({'row': row_idx + 2, 'column': label, 'kind': kind,
                              'value': value, 'from': from_target, 'to': '',
                              'action': 'duplicate'})
                continue
            mapped = [t for t in chain if t in buckets]
            # Prefer a still-empty slot of the right type. When every slot is
            # taken, append to the LAST one ("Other Website" / "Other Phone
            # Number") - that is the natural overflow bucket, and it keeps the
            # primary field clean.
            candidates = ([t for t in mapped if not buckets[t]]
                          + [t for t in chain if t not in buckets]
                          + mapped[::-1]
                          + chain[::-1])
            dest = candidates[0]
            buckets.setdefault(dest, []).append(value)
            fixes.append({'row': row_idx + 2, 'column': label, 'kind': kind,
                          'value': value, 'from': from_target, 'to': dest,
                          'action': 'moved'})

        # ---- flatten into the output record ------------------------------
        record = {h: '' for h in OUTPUT_HEADERS}
        record['Source'] = opts['source_value']
        for target_id, values in buckets.items():
            if not values:
                continue
            ttype = TARGET_BY_ID[target_id]['type']
            if ttype == 'phone':
                joined = dedupe_join_phones(values)
            elif ttype == 'url':
                joined = dedupe_join_urls(values)
            else:
                joined = dedupe_join(values)
            for out_col in TARGET_BY_ID[target_id]['outputs']:
                record[out_col] = joined

        _strip_cross_field_duplicates(record)
        # +2 so the number matches what the operator sees in Excel
        # (row 1 is the header row).
        record['_row'] = row_idx + 2

        if opts['skip_rows_without_name'] and not record['Company Name']:
            continue
        records.append(record)

    return records, fixes


def _strip_cross_field_duplicates(record):
    """The same phone in Mobile and Home (or the same e-mail in Work and
    Other) is noise. Keep the first occurrence, drop the later ones."""
    for outs, kind, joiner in (
        (PHONE_OUTPUTS, 'phone', dedupe_join_phones),
        (EMAIL_OUTPUTS, 'email', dedupe_join),
        (URL_OUTPUTS, 'url', dedupe_join_urls),
    ):
        seen = set()
        for col in outs:
            if not record.get(col):
                continue
            kept = []
            for v in [x.strip() for x in record[col].split(',') if x.strip()]:
                k = _dup_key(kind, v)
                if k and k in seen:
                    continue
                seen.add(k)
                kept.append(v)
            record[col] = joiner(kept)


def records_to_rows(records, sanitize_formulas=False, transliterate=False):
    rows = []
    for rec in records:
        row = [rec.get(h, '') for h in OUTPUT_HEADERS]
        if transliterate:
            row = [strip_accents(v) for v in row]
        if sanitize_formulas:
            row = [_escape_formula(v) for v in row]
        rows.append(row)
    return rows


def _escape_formula(v):
    if isinstance(v, str) and v[:1] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + v
    return v


# ---------------------------------------------------------------------------
# encoding safety net
# ---------------------------------------------------------------------------

# Characters with no useful NFKD decomposition that still deserve an ASCII
# spelling instead of '?'.
MANUAL_TRANSLIT = {
    'ß': 'ss', 'æ': 'ae', 'Æ': 'AE', 'ø': 'o', 'Ø': 'O', 'œ': 'oe', 'Œ': 'OE',
    'đ': 'd', 'Đ': 'D', 'ð': 'd', 'Ð': 'D', 'þ': 'th', 'Þ': 'TH',
    'ł': 'l', 'Ł': 'L', 'ª': 'a', 'º': 'o',
    '®': '(R)', '©': '(C)', '™': '(TM)', '€': 'EUR', '№': 'No',
    '–': '-', '—': '-', '−': '-', '‘': "'", '’': "'", '‚': "'",
    '“': '"', '”': '"', '„': '"', '…': '...', ' ': ' ',
}


def strip_accents(s):
    """é -> e, ñ -> n, ā -> a, ® -> (R).

    Latin letters only: Cyrillic is left alone (NFKD would turn 'й' into
    'и', which is a corruption, not a transliteration)."""
    if not s:
        return s
    out = []
    for ch in s:
        if ch in MANUAL_TRANSLIT:
            out.append(MANUAL_TRANSLIT[ch])
            continue
        decomposed = unicodedata.normalize('NFKD', ch)
        if len(decomposed) > 1 and decomposed[0].isascii():
            out.append(''.join(c for c in decomposed if not unicodedata.combining(c)))
        else:
            out.append(ch)
    return ''.join(out)


def fit_to_encoding(value, encoding):
    """Make `value` representable in `encoding` without raising.

    Falls back gracefully: exact character -> accent-stripped character ->
    '?'. Returns (value, replaced_count) so the UI can warn about losses."""
    if not value:
        return value, 0
    try:
        value.encode(encoding)
        return value, 0
    except UnicodeEncodeError:
        pass
    out = []
    replaced = 0
    for ch in value:
        try:
            ch.encode(encoding)
            out.append(ch)
            continue
        except UnicodeEncodeError:
            pass
        alt = strip_accents(ch)
        try:
            alt.encode(encoding)
            out.append(alt)
            replaced += 1
            continue
        except UnicodeEncodeError:
            pass
        out.append('?')
        replaced += 1
    return ''.join(out), replaced


# ---------------------------------------------------------------------------
# duplicate detection / merging
# ---------------------------------------------------------------------------

PAREN_RE = re.compile(r'^(.*?)\s*\(([^)]+)\)\s*$')


def alnum(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


def name_keys(name):
    """"Distribuidora Alpha (distribuidora_alpha_mx)" produces both
    'distribuidoranautilus' and 'distribuidoranautilusmx', so the row merges
    with a plain 'distribuidora_alpha_mx' row."""
    name = (name or '').strip()
    if not name:
        return []
    m = PAREN_RE.match(name)
    if m:
        return [k for k in (alnum(m.group(1)), alnum(m.group(2))) if len(k) >= 3]
    k = alnum(name)
    return [k] if len(k) >= 3 else []


def pick_best_name(names):
    """Prefer a human-readable name ("Distribuidora Alpha (handle)") over a
    bare handle; longer wins as a tie-break."""
    names = [n for n in names if n]
    if not names:
        return ''
    return max(names, key=lambda n: (' ' in n, len(n)))


class _UnionFind:
    def __init__(self, n):
        self.parent = list(range(n))

    def find(self, i):
        while self.parent[i] != i:
            self.parent[i] = self.parent[self.parent[i]]
            i = self.parent[i]
        return i

    def union(self, i, j):
        ri, rj = self.find(i), self.find(j)
        if ri != rj:
            self.parent[max(ri, rj)] = min(ri, rj)


def _split_multi(value):
    return [v.strip() for v in (value or '').split(',') if v.strip()]


def record_dup_keys(rec, opts):
    """All duplicate keys of one record, tagged by kind so a phone can never
    collide with a company name."""
    keys = []
    if opts['dedupe_by_name']:
        for k in name_keys(rec.get('Company Name', '')):
            keys.append('n:' + k)
    if opts['dedupe_by_email']:
        for col in EMAIL_OUTPUTS:
            for v in _split_multi(rec.get(col, '')):
                keys.append('e:' + v.lower())
    if opts['dedupe_by_phone']:
        for col in PHONE_OUTPUTS:
            for v in _split_multi(rec.get(col, '')):
                d = phone_signature(v)
                if len(d) >= 8:
                    keys.append('p:' + d[-9:])
    if opts['dedupe_by_website']:
        for col in URL_OUTPUTS:
            for v in _split_multi(rec.get(col, '')):
                k = website_dup_key(v)
                if k:
                    keys.append('w:' + k)
    return keys


def dedupe_records(records, options=None):
    """Merge duplicate records.

    Returns (merged, groups):
      merged - list of records, in first-occurrence order
      groups - list of dicts describing every group that had >1 member:
               {'name', 'size', 'rows': [...], 'names': [...], 'matched_on': [...]}
    """
    opts = normalize_options(options)
    if not records:
        return [], []
    if not any(opts[k] for k in ('dedupe_by_name', 'dedupe_by_email',
                                 'dedupe_by_phone', 'dedupe_by_website')):
        return [dict(r) for r in records], []

    uf = _UnionFind(len(records))
    key_owner = {}
    for i, rec in enumerate(records):
        for k in record_dup_keys(rec, opts):
            if k in key_owner:
                uf.union(i, key_owner[k])
            else:
                key_owner[k] = i

    grouped = {}
    order = []
    for i in range(len(records)):
        root = uf.find(i)
        if root not in grouped:
            grouped[root] = []
            order.append(root)
        grouped[root].append(i)

    merged = []
    groups = []
    for root in order:
        members = [records[i] for i in grouped[root]]
        rec = _merge_group(members, opts)
        merged.append(rec)
        if len(members) > 1:
            groups.append({
                'name': rec['Company Name'],
                'size': len(members),
                'rows': [m.get('_row') for m in members],
                'names': [m.get('Company Name', '') for m in members],
                'matched_on': sorted(_infer_reasons(members, opts)),
            })
    return merged, groups


def _infer_reasons(members, opts):
    reasons = set()
    seen = {}
    for m in members:
        for k in record_dup_keys(m, opts):
            if k in seen and seen[k] is not m:
                reasons.add({'n': 'name', 'e': 'e-mail', 'p': 'phone', 'w': 'website'}[k[0]])
            seen[k] = m
    return reasons


def _merge_group(members, opts):
    rec = {h: '' for h in OUTPUT_HEADERS}
    best = pick_best_name([m.get('Company Name', '') for m in members])
    for h in OUTPUT_HEADERS:
        values = []
        for m in members:
            values.extend(_split_multi(m.get(h, '')) if h != 'Source' else [m.get(h, '')])
        if h in ('Company Name', 'Lead Name'):
            rec[h] = best
        elif h in PHONE_OUTPUTS:
            rec[h] = dedupe_join_phones(values)
        elif h in URL_OUTPUTS:
            rec[h] = dedupe_join_urls(values)
        elif h == 'Source':
            rec[h] = next((v for v in values if v), opts['source_value'])
        else:
            rec[h] = dedupe_join(values)
    _strip_cross_field_duplicates(rec)
    rec['_row'] = members[0].get('_row')
    rec['_merged_from'] = [m.get('_row') for m in members]
    return rec


# ---------------------------------------------------------------------------
# reports
# ---------------------------------------------------------------------------

DUP_REPORT_HEADERS = ['Group', 'Kept Company Name', 'Merged rows',
                      'Source row', 'Source Company Name', 'Matched on']


def duplicates_report_rows(groups):
    rows = []
    for i, g in enumerate(groups, start=1):
        for row_no, name in zip(g['rows'], g['names']):
            rows.append([str(i), g['name'], str(g['size']), str(row_no), name,
                         ', '.join(g['matched_on'])])
    return rows


FIX_REPORT_HEADERS = ['Source row', 'Source column', 'Detected type', 'Value',
                      'Mapped to', 'Moved to', 'Action']

_ACTION_LABEL = {'moved': 'moved', 'duplicate': 'duplicate, dropped'}
_KIND_LABEL = {'email': 'e-mail', 'phone': 'phone', 'url': 'website'}


def fixes_report_rows(fixes):
    rows = []
    for f in fixes:
        rows.append([
            str(f['row']), f['column'], _KIND_LABEL.get(f['kind'], f['kind']),
            f['value'],
            TARGET_BY_ID[f['from']]['label'] if f['from'] in TARGET_BY_ID else f['from'],
            TARGET_BY_ID[f['to']]['label'] if f['to'] in TARGET_BY_ID else '',
            _ACTION_LABEL.get(f['action'], f['action']),
        ])
    return rows


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------

def write_csv(path, rows, headers=None, encoding=None):
    """Write the CSV and never raise on an unrepresentable character.

    Returns {'encoding', 'replaced', 'affected_rows'} so the caller can tell
    the operator how much was lost to the chosen codepage."""
    enc = encoding or CSV_ENCODING
    replaced = 0
    affected_rows = 0
    safe_rows = []
    for row in rows:
        row_replaced = 0
        safe = []
        for v in row:
            v, n = fit_to_encoding(v, enc)
            safe.append(v)
            row_replaced += n
        if row_replaced:
            affected_rows += 1
            replaced += row_replaced
        safe_rows.append(safe)

    with open(path, 'w', newline='', encoding=enc) as f:
        writer = csv.writer(
            f, delimiter=CSV_DELIMITER, lineterminator=CSV_LINETERMINATOR,
            quoting=csv.QUOTE_MINIMAL,
        )
        writer.writerow(headers if headers is not None else OUTPUT_HEADERS)
        writer.writerows(safe_rows)
    return {'encoding': enc, 'replaced': replaced, 'affected_rows': affected_rows}


# ---------------------------------------------------------------------------
# backwards-compatible façade
# ---------------------------------------------------------------------------

def build_output_rows(headers, rows, hyperlinks, phone_data, mapping, options=None):
    """Old entry point: mapping -> list of OUTPUT_HEADERS rows, no dedupe."""
    records, _fixes = build_records(headers, rows, hyperlinks, phone_data, mapping, options)
    return records_to_rows(records)
