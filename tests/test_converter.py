# -*- coding: utf-8 -*-
"""
Self-tests for converter.py. No external test runner required:

    py tests\test_converter.py          # from the NEW/ directory
    py -m pytest tests                  # also works if pytest is installed

Every test is a plain function whose name starts with `test_`.
"""
import os
import sys
import csv
import glob
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import converter as C  # noqa: E402


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def make_csv(rows, delimiter=';'):
    fd, path = tempfile.mkstemp(suffix='.csv')
    os.close(fd)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        csv.writer(f, delimiter=delimiter).writerows(rows)
    return path


def run_pipeline(rows, options=None, mapping=None):
    """rows[0] is the header row. Returns (records, fixes, merged, groups)."""
    path = make_csv(rows)
    try:
        headers, data, hyp = C.read_table(path, 'x.csv')
    finally:
        os.unlink(path)
    stats = C.analyze_columns(headers, data)
    phone_cols = C.detect_phone_columns(headers, data, col_stats=stats)
    cc = C.normalize_options(options)['default_country_code']
    phone_data = C.scan_phones(headers, data, phone_cols, cc)
    columns = C.build_source_columns(headers, data, phone_data, stats)
    mapping = mapping if mapping is not None else C.suggest_mapping(columns)
    labels = {c['key']: c['label'] for c in columns}
    records, fixes = C.build_records(headers, data, hyp, phone_data, mapping, options, labels)
    merged, groups = C.dedupe_records(records, options)
    return records, fixes, merged, groups


# ---------------------------------------------------------------------------
# value classification
# ---------------------------------------------------------------------------

def test_classify_email():
    assert C.classify_value('ventas@alpha.mx') == 'email'
    assert C.classify_value('  VENTAS@Alpha.MX ') == 'email'
    # comma typed instead of a dot is still an e-mail
    assert C.classify_value('user@gmail,com') == 'email'


def test_classify_phone():
    assert C.classify_value('+52 33 1234 5678') == 'phone'
    assert C.classify_value('3311110004') == 'phone'
    assert C.classify_value('(664) 111 0007') == 'phone'


def test_classify_url():
    assert C.classify_value('https://alpha.com.mx') == 'url'
    assert C.classify_value('www.alpha.mx') == 'url'
    assert C.classify_value('alpha.com.mx') == 'url'


def test_classify_text_not_url():
    # Instagram handles look like domains but their "TLD" is not real
    assert C.classify_value('distribuidora.beta') == 'text'
    assert C.classify_value('Beta Chile (Betamed)') == 'text'


def test_wame_is_a_phone_not_a_website():
    atoms = C.extract_atoms('wa.me/59399111000')
    assert atoms['phones'] == ['+59399111000']
    assert atoms['urls'] == []


def test_email_digits_do_not_become_a_phone():
    atoms = C.extract_atoms('info2024001@clinic.mx')
    assert atoms['emails'] == ['info2024001@clinic.mx']
    assert atoms['phones'] == []


# ---------------------------------------------------------------------------
# header detection
# ---------------------------------------------------------------------------

def test_phone_header_detection():
    assert C.looks_like_phone_header('Phone or WhatsApp number')
    assert C.looks_like_phone_header('Teléfono')
    assert C.looks_like_phone_header('Cel')


def test_phone_header_false_positives():
    # 'wa' inside 'Warehouse', 'tel' inside 'Hotel', 'contact email'
    assert not C.looks_like_phone_header('Warehouse')
    assert not C.looks_like_phone_header('Hotel name')
    assert not C.looks_like_phone_header('Contact email')
    assert not C.looks_like_phone_header('Website')
    # 'id' must not veto a real phone column
    assert C.looks_like_phone_header('Provider phone')


def test_email_column_is_not_detected_as_phone():
    rows = [
        ['Company', 'Contact email'],
        ['Alpha', 'a@alpha.mx'],
        ['Beta', 'b@beta.mx'],
    ]
    path = make_csv(rows)
    try:
        headers, data, _ = C.read_table(path, 'x.csv')
    finally:
        os.unlink(path)
    assert C.detect_phone_columns(headers, data) == []


# ---------------------------------------------------------------------------
# misplaced values (swapped / wrong-column data)
# ---------------------------------------------------------------------------

def test_swapped_email_and_phone_are_restored():
    records, fixes, _, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha SA', '+52 33 1234 5678', 'ventas@alpha.mx', 'alpha.com.mx'],
    ])
    r = records[0]
    assert r['Work E-mail'] == 'ventas@alpha.mx'
    assert r['Mobile Phone'] == '+52 33 1234 5678'
    assert r['Corporate Website'] == 'https://alpha.com.mx'
    assert {f['kind'] for f in fixes} == {'email', 'phone'}
    assert all(f['action'] == 'moved' for f in fixes)


def test_url_in_phone_column_goes_to_a_website_field():
    records, fixes, _, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Beta', 'b@beta.cl', '+56911110008 / https://instagram.com/beta.cl', 'https://beta.cl'],
    ])
    r = records[0]
    assert r['Mobile Phone'] == '+56911110008'
    assert r['Corporate Website'] == 'https://beta.cl'
    assert 'instagram.com/beta.cl' in r['Other Website']
    assert [f['action'] for f in fixes] == ['moved']


def test_relocated_value_already_present_is_dropped():
    records, fixes, _, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Gamma', '', 'https://gamma.cl', 'https://gamma.cl'],
    ])
    assert records[0]['Corporate Website'] == 'https://gamma.cl'
    assert records[0]['Other Website'] == ''
    assert [f['action'] for f in fixes] == ['duplicate']


def test_autofix_can_be_switched_off():
    records, fixes, _, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha SA', '+52 33 1234 5678', 'ventas@alpha.mx', 'alpha.com.mx'],
    ], options={'autofix_types': False})
    assert records[0]['Work E-mail'] == ''
    assert fixes == []


def test_mis_detected_phone_column_does_not_lose_data():
    """A column whose header screams "phone" but which actually holds
    e-mails must still be recoverable."""
    records, _, _, _ = run_pipeline([
        ['Company', 'Phone'],
        ['Alpha', 'ventas@alpha.mx'],
        ['Beta', 'b@beta.mx'],
    ])
    assert records[0]['Work E-mail'] == 'ventas@alpha.mx'


# ---------------------------------------------------------------------------
# phones
# ---------------------------------------------------------------------------

def test_normalize_phone_statuses():
    assert C.normalize_phone('+52 33 1234 5678') == ('+52 33 1234 5678', 'ok')
    assert C.normalize_phone('0052 33 1234 5678')[1] == 'ok'
    assert C.normalize_phone('523312345678') == ('+523312345678', 'fixed')
    assert C.normalize_phone('12345')[1] == 'invalid'


def test_default_country_code():
    val, status = C.normalize_phone('3312345678', default_cc='52')
    assert status == 'cc_added'
    assert val.startswith('+52')
    # a number that already carries the country code is left alone
    assert C.normalize_phone('523312345678', default_cc='52')[1] == 'fixed'


def test_several_phones_in_one_cell_are_split():
    atoms = C.extract_atoms('+52 33 1111 0005 / 800 111 0010')
    assert len(atoms['phones']) == 2


def test_same_number_with_and_without_country_code_collapses():
    assert C.dedupe_join_phones(['+523311110002', '3311110002']) == '+523311110002'


# ---------------------------------------------------------------------------
# duplicates
# ---------------------------------------------------------------------------

def test_dedupe_by_name_with_handle_in_parens():
    _, _, merged, groups = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['distribuidora_alpha_mx', 'g@x.com', '+52 3311110001', ''],
        ['Distribuidora Alpha (distribuidora_alpha_mx)', '', '+52 3311110003', ''],
    ])
    assert len(merged) == 1
    assert merged[0]['Company Name'] == 'Distribuidora Alpha (distribuidora_alpha_mx)'
    assert '+52 3311110001' in merged[0]['Mobile Phone']
    assert '+52 3311110003' in merged[0]['Mobile Phone']
    assert groups[0]['size'] == 2


def test_dedupe_by_email():
    _, _, merged, groups = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', 'shared@x.com', '+52 331 111 1111', ''],
        ['Alpha Mexico', 'shared@x.com', '+52 331 222 2222', ''],
    ])
    assert len(merged) == 1
    assert 'e-mail' in groups[0]['matched_on']


def test_dedupe_by_phone_can_be_disabled():
    _, _, merged, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', '', '+52 331 111 1111', ''],
        ['Totally Different Co', '', '+52 331 111 1111', ''],
    ], options={'dedupe_by_phone': False})
    assert len(merged) == 2


def test_raw_output_keeps_every_row():
    records, _, merged, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', 'a@x.com', '', ''],
        ['Alpha', 'a@x.com', '', ''],
        ['Alpha', 'a@x.com', '', ''],
    ])
    assert len(records) == 3      # "with duplicates" file
    assert len(merged) == 1       # "clean" file


def test_dedupe_off_entirely():
    records, _, merged, groups = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', 'a@x.com', '', ''],
        ['Alpha', 'a@x.com', '', ''],
    ], options={'dedupe_by_name': False, 'dedupe_by_email': False,
                'dedupe_by_phone': False, 'dedupe_by_website': False})
    assert len(merged) == len(records) == 2
    assert groups == []


def test_website_dedupe_does_not_merge_all_instagram_leads():
    _, _, merged, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', '', '', 'https://instagram.com/alpha'],
        ['Beta', '', '', 'https://instagram.com/beta'],
    ], options={'dedupe_by_website': True})
    assert len(merged) == 2


def test_merged_row_numbers_point_at_the_spreadsheet():
    records, _, _, groups = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['Alpha', 'a@x.com', '', ''],
        ['Alpha', 'a@x.com', '', ''],
    ])
    # header is row 1, so the two data rows are 2 and 3
    assert [r['_row'] for r in records] == [2, 3]
    assert groups[0]['rows'] == [2, 3]


# ---------------------------------------------------------------------------
# cross-field cleanup + output
# ---------------------------------------------------------------------------

def test_same_phone_in_two_fields_is_kept_once():
    records, _, _, _ = run_pipeline([
        ['Company', 'Mobile', 'Home phone'],
        ['Alpha', '+52 331 111 1111', '+52 331 111 1111'],
    ])
    r = records[0]
    assert r['Mobile Phone'] == '+52 331 111 1111'
    assert r['Home Phone'] == ''


def test_skip_rows_without_name():
    records, _, _, _ = run_pipeline([
        ['Company', 'Email', 'Phone', 'Website'],
        ['', 'a@x.com', '', ''],
        ['Alpha', 'b@x.com', '', ''],
    ], options={'skip_rows_without_name': True})
    assert len(records) == 1


def test_csv_dialect():
    fd, path = tempfile.mkstemp(suffix='.csv')
    os.close(fd)
    try:
        C.write_csv(path, [['Alpha; Inc', 'Alpha', '', '', '', '', '', '', '', '',
                            '', '', '', 'import']])
        with open(path, 'rb') as f:
            raw = f.read()
        assert raw.startswith(b'\xef\xbb\xbf')           # BOM by default
        assert b'\r\n' in raw                            # CRLF
        assert raw.split(b'\r\n')[0].count(b';') == 13   # ';' separated
        assert b'"Alpha; Inc"' in raw                    # quoted only when needed
    finally:
        os.unlink(path)


# ---------------------------------------------------------------------------
# encoding
# ---------------------------------------------------------------------------

def _write_and_read(value, encoding=None, transliterate=False):
    fd, path = tempfile.mkstemp(suffix='.csv')
    os.close(fd)
    rec = {h: '' for h in C.OUTPUT_HEADERS}
    rec['Company Name'] = value
    rows = C.records_to_rows([rec], transliterate=transliterate)
    try:
        report = C.write_csv(path, rows, encoding=encoding)
        with open(path, 'rb') as f:
            return f.read(), report
    finally:
        os.unlink(path)


def test_default_output_has_bom_so_excel_detects_utf8():
    """Without a BOM, Excel/CRM read UTF-8 'é' (C3 A9) as two ANSI characters;
    a later ASCII conversion turns each into '?' - hence 'M??xico'."""
    raw, _ = _write_and_read('Alpha Estética México')
    assert raw.startswith(b'\xef\xbb\xbf')
    assert 'México'.encode('utf-8') in raw
    assert b'M??xico' not in raw


def test_utf8_without_bom_is_still_available():
    raw, report = _write_and_read('Alpha Estética México', encoding='utf-8')
    assert not raw.startswith(b'\xef\xbb\xbf')
    assert report['replaced'] == 0


def _data_line(raw, encoding):
    """Second line of the CSV - the first one is the header row."""
    return raw.decode(encoding).split('\r\n')[1]


def test_cp1252_keeps_spanish_accents():
    raw, report = _write_and_read('Alpha Estética México', encoding='cp1252')
    assert report['replaced'] == 0
    assert _data_line(raw, 'cp1252').startswith('Alpha Estética México')


def test_strip_accents():
    assert C.strip_accents('Alpha Estética México (alphaestética)') == \
        'Alpha Estetica Mexico (alphaestetica)'
    assert C.strip_accents('Alphā Médica') == 'Alpha Medica'
    assert C.strip_accents('Estética Uruguay') == 'Estetica Uruguay'
    assert C.strip_accents('AllA Medical Group®') == 'AllA Medical Group(R)'


def test_strip_accents_leaves_cyrillic_alone():
    # NFKD would turn 'й' into 'и' - that is corruption, not transliteration
    assert C.strip_accents('Йошкар-Ола, Россия') == 'Йошкар-Ола, Россия'


def test_transliterate_gives_pure_ascii():
    raw, report = _write_and_read('Alpha Estética México (alphaestética)',
                                  encoding='utf-8', transliterate=True)
    assert report['replaced'] == 0
    assert all(b < 128 for b in raw)
    assert b'Alpha Estetica Mexico (alphaestetica)' in raw


def test_unrepresentable_character_degrades_instead_of_crashing():
    # 'ā' does not exist in cp1252 - must become 'a', not blow up
    raw, report = _write_and_read('Alphā Médica', encoding='cp1252')
    assert report['replaced'] == 1
    assert _data_line(raw, 'cp1252').startswith('Alpha Médica')


def test_cyrillic_in_cp1252_falls_back_to_question_marks():
    _raw, report = _write_and_read('Компания', encoding='cp1252')
    assert report['replaced'] == 8
    assert report['affected_rows'] == 1


def test_unknown_encoding_falls_back_to_default():
    assert C.normalize_options({'csv_encoding': 'rot13'})['csv_encoding'] == C.CSV_ENCODING
    assert C.normalize_options({'csv_encoding': 'cp1251'})['csv_encoding'] == 'cp1251'


def test_formula_sanitizing():
    recs = [{h: '' for h in C.OUTPUT_HEADERS}]
    recs[0]['Company Name'] = '=cmd|calc'
    recs[0]['Mobile Phone'] = '+52 331 111 1111'
    plain = C.records_to_rows(recs, sanitize_formulas=False)[0]
    safe = C.records_to_rows(recs, sanitize_formulas=True)[0]
    assert plain[0] == '=cmd|calc'
    assert safe[0] == "'=cmd|calc"
    assert safe[2] == "'+52 331 111 1111"


def test_suggest_mapping_spills_extra_phone_columns():
    columns = [
        {'key': '0', 'kind': 'text', 'analysis': {'header_hint': 'company_lead', 'dominant': 'text'}},
        {'key': '1:0', 'kind': 'phone', 'analysis': {'header_hint': 'phone', 'dominant': 'phone'}},
        {'key': '1:1', 'kind': 'phone', 'analysis': {'header_hint': 'phone', 'dominant': 'phone'}},
        {'key': '1:2', 'kind': 'phone', 'analysis': {'header_hint': 'phone', 'dominant': 'phone'}},
        {'key': '1:3', 'kind': 'phone', 'analysis': {'header_hint': 'phone', 'dominant': 'phone'}},
    ]
    m = C.suggest_mapping(columns)
    assert m['1:0'] == 'mobile_phone'
    assert m['1:1'] == 'home_phone'
    assert m['1:2'] == 'other_phone'
    # nothing is dropped: the 4th column joins the last phone field
    assert m['1:3'] == 'other_phone'


# ---------------------------------------------------------------------------
# regression: the real source file, when it is available
# ---------------------------------------------------------------------------

def _find_sample():
    """Locate a real lead spreadsheet to regression-test against.

    Point SAMPLE_XLSX at your own file, or drop any .xlsx in a directory above
    the repository. The file itself is never committed - it holds live lead
    data - so this test simply skips when nothing is found.
    """
    explicit = os.environ.get('SAMPLE_XLSX')
    if explicit and os.path.isfile(explicit):
        return explicit
    d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(5):
        d = os.path.dirname(d)
        found = sorted(glob.glob(os.path.join(d, '*.xlsx')))
        if found:
            return found[0]
    return ''


SAMPLE_XLSX = _find_sample()


def test_real_file_regression():
    if not os.path.isfile(SAMPLE_XLSX):
        print('  (skipped: no sample .xlsx found; set SAMPLE_XLSX to run)')
        return
    headers, rows, hyp = C.read_table(SAMPLE_XLSX, os.path.basename(SAMPLE_XLSX))
    assert len(rows) == 436
    stats = C.analyze_columns(headers, rows)
    phone_cols = C.detect_phone_columns(headers, rows, col_stats=stats)
    assert phone_cols == [2]                       # only the phone column
    assert stats[1]['dominant'] == 'email'
    assert stats[3]['dominant'] == 'url'
    phone_data = C.scan_phones(headers, rows, phone_cols)
    columns = C.build_source_columns(headers, rows, phone_data, stats)
    mapping = C.suggest_mapping(columns)
    assert mapping['0'] == 'company_lead'
    assert mapping['1'] == 'work_email'
    records, fixes = C.build_records(
        headers, rows, hyp, phone_data, mapping, {},
        {c['key']: c['label'] for c in columns})
    assert len(records) == 436
    merged, groups = C.dedupe_records(records, {})
    assert len(merged) < 380 and len(merged) > 300  # ~355 with default keys
    assert len(groups) > 40
    # websites typed into the phone column were rescued, not dropped
    assert any(f['kind'] == 'url' for f in fixes)


# ---------------------------------------------------------------------------

def main():
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith('test_') and callable(f)]
    failed = []
    for name, fn in tests:
        try:
            fn()
            print(f'PASS  {name}')
        except Exception as e:  # noqa: BLE001
            failed.append((name, e))
            print(f'FAIL  {name}: {type(e).__name__}: {e}')
    print(f'\n{len(tests) - len(failed)}/{len(tests)} passed')
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
