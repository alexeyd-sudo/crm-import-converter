"""
Flask front-end for the CRM import converter.

All real work lives in converter.py; this module only handles uploads,
session state and file downloads.

Endpoints
---------
GET  /                          the single-page UI
POST /api/upload                file -> session + column analysis + suggested mapping
POST /api/convert               mapping + options -> 4 downloadable CSVs
GET  /download/<sid>/<kind>     kind = raw | merged | duplicates | fixes
GET  /download/<sid>            legacy alias for kind = raw
GET  /healthz                   liveness probe
"""
import os
import re
import shutil
import time
import uuid
import threading

from flask import Flask, request, jsonify, render_template, send_file, abort
from werkzeug.exceptions import HTTPException

import converter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Uploaded lead tables are personal data and must not live inside the code
# directory: on a server /opt is root-owned and read-only for the service user.
UPLOAD_DIR = os.environ.get('UPLOAD_DIR', os.path.join(BASE_DIR, 'uploads'))
os.makedirs(UPLOAD_DIR, mode=0o700, exist_ok=True)

ALLOWED_EXT = {'.xlsx', '.xlsm', '.csv', '.txt'}
# 8 MB of compressed upload is already ~200 MB of sheet XML; converter.py caps
# the uncompressed size separately, which is the limit that actually bounds
# CPU and RAM. See MAX_XLSX_UNCOMPRESSED there.
MAX_CONTENT_LENGTH = int(os.environ.get('MAX_CONTENT_LENGTH', 8 * 1024 * 1024))
SESSION_TTL_SECONDS = int(os.environ.get('SESSION_TTL_SECONDS', 2 * 3600))
# Each session holds the whole parsed table in RAM until it expires, so this
# is a memory setting, not a comfort setting: cap * MAX_XLSX_UNCOMPRESSED * ~5.
MAX_SESSIONS = int(os.environ.get('MAX_SESSIONS', 8))

# Default port. The Russian and English builds of this service must listen
# on DIFFERENT ports, otherwise the second one to start dies with
# 'address already in use'. Override with the PORT environment variable.
PORT = int(os.environ.get('PORT', 5001))

SESSION_ID_RE = re.compile(r'^[0-9a-f]{32}$')

# Files produced per session. kind -> (filename, download name)
ARTIFACTS = {
    'raw': ('result_raw.csv', 'crm_import_with_duplicates.csv'),
    'merged': ('result_merged.csv', 'crm_import_merged.csv'),
    'duplicates': ('report_duplicates.csv', 'duplicates_report.csv'),
    'fixes': ('report_fixes.csv', 'field_fixes_report.csv'),
}

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# In-memory session store: one process, fine for single-worker deployment.
# {session_id: {'created': ts, 'headers': ..., 'rows': ..., ...}}
SESSIONS = {}
SESSIONS_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# session housekeeping
# ---------------------------------------------------------------------------

def _drop_session(session_id):
    """Caller must NOT hold SESSIONS_LOCK: rmtree is blocking disk I/O and
    would stall every other request while it runs."""
    SESSIONS.pop(session_id, None)
    shutil.rmtree(os.path.join(UPLOAD_DIR, session_id), ignore_errors=True)


def get_live_session(session_id):
    """Fetch a session, enforcing the TTL here rather than only at the next
    upload - otherwise an expired session keeps serving lead data for as long
    as nobody uploads a new file."""
    now = time.time()
    with SESSIONS_LOCK:
        session = SESSIONS.get(session_id)
        expired = bool(session) and now - session['created'] > SESSION_TTL_SECONDS
        if expired:
            SESSIONS.pop(session_id, None)
    if expired:
        shutil.rmtree(os.path.join(UPLOAD_DIR, session_id), ignore_errors=True)
        return None
    return session


def cleanup_sessions():
    """Expire old sessions and cap total count. Without this both RAM and
    uploads/ grow forever - the uploaded tables contain lead data, so they
    should not linger on disk either."""
    now = time.time()
    with SESSIONS_LOCK:
        doomed = [sid for sid, s in SESSIONS.items()
                  if now - s['created'] > SESSION_TTL_SECONDS]
        for sid in doomed:
            SESSIONS.pop(sid, None)
        # `>=` because this runs just BEFORE a new session is added, so `>`
        # left room for MAX_SESSIONS + 1 to be resident.
        while len(SESSIONS) >= MAX_SESSIONS:
            oldest = min(SESSIONS, key=lambda sid: SESSIONS[sid]['created'])
            SESSIONS.pop(oldest, None)
            doomed.append(oldest)
    for sid in doomed:
        shutil.rmtree(os.path.join(UPLOAD_DIR, sid), ignore_errors=True)

    # orphan directories left behind by a previous process
    cutoff = now - SESSION_TTL_SECONDS
    try:
        for name in os.listdir(UPLOAD_DIR):
            path = os.path.join(UPLOAD_DIR, name)
            if not os.path.isdir(path) or not SESSION_ID_RE.match(name):
                continue
            with SESSIONS_LOCK:
                if name in SESSIONS:
                    continue
            if os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
    except OSError:
        pass


def _session_dir(session_id):
    return os.path.join(UPLOAD_DIR, session_id)


# ---------------------------------------------------------------------------
# routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html', target_fields=converter.TARGET_FIELDS)


@app.route('/healthz')
def healthz():
    # Deliberately says nothing about load: this endpoint is reachable without
    # auth so a probe can use it, and the session count is nobody else's business.
    return jsonify({'status': 'ok'})


@app.route('/api/upload', methods=['POST'])
def upload():
    cleanup_sessions()

    file = request.files.get('file')
    if not file or not file.filename:
        return jsonify({'error': 'No file selected'}), 400

    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({'error': f'Unsupported file type: {ext or "unknown"}. '
                                 f'Supported formats: '
                                 f'{", ".join(sorted(ALLOWED_EXT))}.'}), 400

    session_id = uuid.uuid4().hex
    session_dir = _session_dir(session_id)
    os.makedirs(session_dir, mode=0o700, exist_ok=True)
    src_path = os.path.join(session_dir, 'source' + ext)
    file.save(src_path)
    os.chmod(src_path, 0o600)

    try:
        headers, rows, hyperlinks = converter.read_table(src_path, filename)
    except converter.TableTooLargeError as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({'error': str(e)}), 400
    except Exception:
        # The exception text can carry server paths, so it goes to the log and
        # the client gets a generic message.
        app.logger.exception('could not read uploaded file %r', filename)
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({'error': 'Could not read the file. Check that it is a '
                                 'valid spreadsheet and that the first row holds '
                                 'the column headers.'}), 400

    if not headers or not rows:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({'error': 'No data found in the file '
                                 '(check that the first row contains the column headers).'}), 400

    col_stats = converter.analyze_columns(headers, rows)
    phone_cols = converter.detect_phone_columns(headers, rows, col_stats=col_stats)
    phone_data = converter.scan_phones(headers, rows, phone_cols)
    source_columns = converter.build_source_columns(headers, rows, phone_data, col_stats)
    suggested = converter.suggest_mapping(source_columns)

    with SESSIONS_LOCK:
        SESSIONS[session_id] = {
            'created': time.time(),
            'headers': headers,
            'rows': rows,
            'hyperlinks': hyperlinks,
            'phone_data': phone_data,
            'col_stats': col_stats,
            'labels': {c['key']: c['label'] for c in source_columns},
            'filename': filename,
        }

    phone_summary = []
    for c in phone_cols:
        info = phone_data[c]
        phone_summary.append({
            'header': headers[c],
            'max_count': info['max_count'],
            'stats': info['stats'],
        })

    return jsonify({
        'session_id': session_id,
        'row_count': len(rows),
        'source_columns': source_columns,
        'target_fields': converter.TARGET_FIELDS,
        'phone_summary': phone_summary,
        'suggested_mapping': suggested,
        'warnings': _upload_warnings(headers, rows, phone_data, col_stats),
        'defaults': converter.DEFAULT_OPTIONS,
        'encodings': converter.CSV_ENCODINGS,
    })


def _upload_warnings(headers, rows, phone_data, col_stats):
    """Things the operator should look at before pressing "convert"."""
    out = []
    for c, info in phone_data.items():
        st = info['stats']
        if st['overflow']:
            out.append(f'Column "{headers[c]}" has cells holding more than '
                       f'{converter.MAX_PHONES_PER_CELL} phone numbers - '
                       f'{st["overflow"]} did not fit.')
        if st['foreign_emails']:
            out.append(f'Phone column "{headers[c]}" contains '
                       f'{st["foreign_emails"]} e-mail addresses - they will be moved '
                       f'to an E-mail field.')
        if st['foreign_urls']:
            out.append(f'Phone column "{headers[c]}" contains '
                       f'{st["foreign_urls"]} links - they will be moved '
                       f'to a Website field.')
        if st['fixed']:
            out.append(f'{st["fixed"]} numbers had no "+" but were long enough to already '
                       f'contain a country code, so "+" was added.')
        if st['needs_cc']:
            out.append(f'{st["needs_cc"]} numbers in "{headers[c]}" are national numbers '
                       f'with no country code. They are kept exactly as they came, '
                       f'WITHOUT a "+", because guessing would invent a country: '
                       f'"449 478 2400" would become "+449 478 2400", which reads as '
                       f'the United Kingdom. Set the default country code below to fix '
                       f'them - otherwise the CRM gets numbers nobody can dial.')
    for st in col_stats:
        if st['nonempty'] and st['dominant'] and st['share'] < 0.7:
            out.append(f'Column "{st["header"]}" is mixed: '
                       + ', '.join(f'{k}={v}' for k, v in st['counts'].items() if v)
                       + '. Auto-fix will sort the values by type.')
    return out


@app.route('/api/convert', methods=['POST'])
def convert():
    data = request.get_json(force=True, silent=True)
    if not isinstance(data, dict):
        return jsonify({'error': 'Expected a JSON object.'}), 400
    session_id = data.get('session_id') or ''
    mapping = data.get('mapping') or {}
    options = data.get('options') or {}

    if not SESSION_ID_RE.match(session_id):
        return jsonify({'error': 'Invalid session id.'}), 400

    # Everything below arrives from the browser, so nothing about its shape is
    # guaranteed. Each of these used to be an unhandled 500.
    if not isinstance(mapping, dict) or not isinstance(options, dict):
        return jsonify({'error': 'mapping and options must be JSON objects.'}), 400
    bad_keys = [k for k in mapping if not converter.SOURCE_KEY_RE.match(str(k))]
    if bad_keys:
        return jsonify({'error': f'Invalid source column key: {bad_keys[0]!r}'}), 400
    bad_targets = [v for v in mapping.values() if v and v not in converter.TARGET_BY_ID]
    if bad_targets:
        return jsonify({'error': f'Unknown CRM field: {bad_targets[0]!r}'}), 400

    session = get_live_session(session_id)
    if not session:
        return jsonify({'error': 'Session not found or expired. Please upload the file again.'}), 404

    unknown = [k for k in mapping if k not in session['labels']]
    if unknown:
        return jsonify({'error': f'Column {unknown[0]!r} is not part of the uploaded '
                                 f'file. Please upload it again.'}), 400

    if not any(mapping.values()):
        return jsonify({'error': 'No source column was mapped to a CRM field.'}), 400

    opts = converter.normalize_options(options)

    # The default country code changes phone normalization, so the phone
    # columns have to be re-scanned when it is set.
    phone_data = session['phone_data']
    if opts['default_country_code']:
        phone_data = converter.scan_phones(
            session['headers'], session['rows'], list(phone_data.keys()),
            opts['default_country_code'])

    records, fixes = converter.build_records(
        session['headers'], session['rows'], session['hyperlinks'],
        phone_data, mapping, opts, session['labels'],
    )
    merged, groups = converter.dedupe_records(records, opts)

    session_dir = _session_dir(session_id)
    os.makedirs(session_dir, exist_ok=True)
    sanitize = bool(opts['sanitize_formulas'])
    translit = bool(opts['transliterate'])
    enc = opts['csv_encoding']

    raw_rows = converter.records_to_rows(records, sanitize, translit)
    merged_rows = converter.records_to_rows(merged, sanitize, translit)

    enc_raw = converter.write_csv(os.path.join(session_dir, ARTIFACTS['raw'][0]),
                                  raw_rows, encoding=enc)
    enc_merged = converter.write_csv(os.path.join(session_dir, ARTIFACTS['merged'][0]),
                                     merged_rows, encoding=enc)
    # Reports are for humans in Excel - always UTF-8 with BOM.
    converter.write_csv(os.path.join(session_dir, ARTIFACTS['duplicates'][0]),
                        converter.duplicates_report_rows(groups),
                        converter.DUP_REPORT_HEADERS, encoding='utf-8-sig')
    converter.write_csv(os.path.join(session_dir, ARTIFACTS['fixes'][0]),
                        converter.fixes_report_rows(fixes),
                        converter.FIX_REPORT_HEADERS, encoding='utf-8-sig')

    with SESSIONS_LOCK:
        if session_id in SESSIONS:
            SESSIONS[session_id]['encoding'] = enc

    moved = sum(1 for f in fixes if f['action'] == 'moved')
    dropped = sum(1 for f in fixes if f['action'] == 'duplicate')

    return jsonify({
        'headers': converter.OUTPUT_HEADERS,
        'encoding': {
            'id': enc,
            'label': next((e['label'] for e in converter.CSV_ENCODINGS if e['id'] == enc), enc),
            'transliterate': translit,
            'replaced': enc_raw['replaced'] + enc_merged['replaced'],
            'affected_rows': max(enc_raw['affected_rows'], enc_merged['affected_rows']),
        },
        'raw': {
            'row_count': len(records),
            'filled': _filled(records),
            'preview': raw_rows[:5],
            'download_url': f'/download/{session_id}/raw',
        },
        'merged': {
            'row_count': len(merged),
            'filled': _filled(merged),
            'preview': merged_rows[:5],
            'download_url': f'/download/{session_id}/merged',
        },
        'duplicates': {
            'group_count': len(groups),
            'collapsed_rows': len(records) - len(merged),
            'review_count': sum(1 for g in groups if g['matched_on'] == ['phone']),
            'download_url': f'/download/{session_id}/duplicates',
            'groups': groups[:50],
        },
        'fixes': {
            'moved': moved,
            'dropped_duplicates': dropped,
            'download_url': f'/download/{session_id}/fixes',
            'items': fixes[:50],
        },
        # legacy field kept so older bookmarks/scripts keep working
        'row_count': len(records),
        'preview': raw_rows[:5],
        'filled': _filled(records),
        'download_url': f'/download/{session_id}/raw',
    })


def _filled(records):
    counts = {h: 0 for h in converter.OUTPUT_HEADERS}
    for rec in records:
        for h in converter.OUTPUT_HEADERS:
            if rec.get(h):
                counts[h] += 1
    return counts


@app.route('/download/<session_id>')
@app.route('/download/<session_id>/<kind>')
def download(session_id, kind='raw'):
    if not SESSION_ID_RE.match(session_id or ''):
        abort(404)
    if kind not in ARTIFACTS:
        abort(404)
    if not get_live_session(session_id):
        abort(404)
    fname, download_name = ARTIFACTS[kind]
    path = os.path.join(_session_dir(session_id), fname)
    if not os.path.isfile(path):
        abort(404)
    # Reports are always UTF-8+BOM; the two result files follow the operator's
    # choice. Declaring the charset helps anything that fetches over HTTP.
    if kind in ('raw', 'merged'):
        session = SESSIONS.get(session_id) or {}
        charset = session.get('encoding') or converter.CSV_ENCODING
    else:
        charset = 'utf-8-sig'
    charset = 'utf-8' if charset == 'utf-8-sig' else charset
    resp = send_file(path, as_attachment=True, download_name=download_name,
                     mimetype='text/csv')
    # Setting the charset via `mimetype` makes Werkzeug append its own default
    # too, producing "text/csv; charset=cp1251; charset=utf-8" - a contradiction
    # that decodes to mojibake in any client that reads the last parameter.
    resp.headers['Content-Type'] = f'text/csv; charset={charset}'
    # The file holds customer contact details: no shared cache should keep it.
    resp.headers['Cache-Control'] = 'private, no-store'
    return resp


@app.errorhandler(413)
def too_large(_e):
    mb = MAX_CONTENT_LENGTH // (1024 * 1024)
    return jsonify({'error': f'File is larger than {mb} MB.'}), 413


@app.errorhandler(ValueError)
def bad_input(e):
    """converter raises ValueError for input it refuses to trust."""
    return jsonify({'error': str(e)}), 400


@app.errorhandler(Exception)
def unhandled(e):
    """The frontend parses every response as JSON, so an HTML 500 page shows up
    as an unreadable error. Anything unexpected still gets logged in full."""
    if isinstance(e, HTTPException):
        return e
    app.logger.exception('unhandled error on %s', request.path)
    return jsonify({'error': 'Internal error. Please try again.'}), 500


if __name__ == '__main__':
    # Development entry point. debug=True would expose the Werkzeug debugger
    # console - arbitrary code execution - so it has to be asked for explicitly,
    # and the default bind is loopback, not every interface.
    app.run(host=os.environ.get('HOST', '127.0.0.1'), port=PORT,
            debug=os.environ.get('FLASK_DEBUG') == '1')
