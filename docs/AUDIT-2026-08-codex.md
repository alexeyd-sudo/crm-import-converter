# Вердикт

**Как есть деплоить нельзя. Готовность: 3/10.**

Основной поток работает: XLSX читается, поля определяются, CSV формируются корректно, штатные тесты проходят. Но для публичного сервиса на общей машине остаются критичные проблемы: отсутствие авторизации и rate limiting, реальный memory-DoS через сжатый XLSX, несовместимые с несколькими gunicorn workers сессии, неработающий TTL, CSV injection и подтверждённое ошибочное объединение разных компаний по телефону.

Репозиторий не изменялся: `git status` чистый, проверенный commit — `61b180c3066c408b42c05049f68962440c338f21`. Все временные файлы находятся в `/tmp/crm-import-converter-audit.wcnFoY`.

# Что реально проверено запуском

Среда:

```text
Python 3.13.9
Flask 3.1.3
openpyxl 3.1.5
waitress 3.0.2
gunicorn 26.0.0 — установлен отдельно, в requirements его нет
```

`pip-audit --local`:

```text
No known vulnerabilities found
```

Это относится только к фактически установленному набору; `requirements.txt` сейчас не фиксирует эти версии.

Тесты с явно заданным реальным файлом:

```text
SAMPLE_XLSX="/Users/loschev/PycharmProjects/crm-import-audit/Ernesto crm.xlsx" \
python tests/test_converter.py

PASS  test_real_file_regression
...
40/40 passed
```

Сервис успешно запущен через WSGI `app:app`:

```text
INFO:waitress:Serving on http://127.0.0.1:51281
```

Реальный файл:

```text
Размер:        44 872 bytes
Листов:        1, "Structured Contacts"
Размер таблицы: 437 × 5, то есть 436 строк данных
Формул:        0
Гиперссылок:   283
```

HTTP-загрузка:

```text
POST /api/upload
HTTP_STATUS=200
TIME_TOTAL=0.041483
row_count=436

Phone statistics:
ok=254
fixed=108
foreign_urls=11
max_count=5
```

HTTP-конвертация с предложенным mapping:

```text
POST /api/convert
HTTP_STATUS=200
TIME_TOTAL=0.036136

raw.row_count=436
merged.row_count=355
duplicate groups=57
collapsed rows=81
fixes: moved=2, duplicate/dropped=10
encoding=utf-8-sig
replaced=0
```

Фактически скачанные файлы:

| Файл | Размер | Строк данных | Колонок |
|---|---:|---:|---:|
| `crm_import_with_duplicates.csv` | 51 010 B | 436 | 14 |
| `crm_import_merged.csv` | 44 825 B | 355 | 14 |
| `duplicates_report.csv` | 10 601 B | 138 | 6 |
| `field_fixes_report.csv` | 1 467 B | 12 | 7 |

У всех четырёх файлов подтверждены BOM `EF BB BF`, CRLF и одинаковая ширина каждой CSV-строки.

Граничные проверки:

| Сценарий | Фактический результат |
|---|---|
| Пустой CSV | `400 No data found` |
| `.xls` | `400 Unsupported file type: .xls` |
| `/etc/hosts`, названный `renamed.xlsx` | `400 File is not a zip file` |
| Файл >25 MiB | JSON `413 File is larger than 25 MB` |
| Формула без cached value в XLSX | Не выполняется; строка исчезает, ответ `400 No data found` |
| Имя `../../escaped.xlsx` | `200`, записан только `uploads/<uuid>/source.xlsx` |
| `/download/../../etc/passwd` | `404` |
| 301 колонка | `200`, молча оставлены только первые 300 |
| CP1252 CSV | `Clínica México → Clнnica Mйxico`, `José → Josй` |
| Mapping только в Comment | `200`, 436 строк с пустыми Company/Lead Name |
| 100 000 строк, 8.86 MB | upload 1.69 с; convert 8.96 с; пик RSS ≈325 MB; 32 MB на диске |
| 10 последовательных XLSX-upload | число открытых FD осталось `42 → 42` |

# БЛОКЕРЫ — обязательно до деплоя

1. **Публичные эндпоинты полностью открыты**

   Места: [app.py:108](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:108), [app.py:120](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:120), [app.py:223](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:223), [app.py:335](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:335). Документация сама это подтверждает в [DEVELOPER.md:396](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/docs/DEVELOPER.md:396).

   Любой пользователь Интернета сможет загружать файлы, занимать RAM/диск/CPU и обрабатывать персональные данные. Запрос с чужим Origin был принят:

   ```text
   Origin: https://evil.example
   POST /api/upload
   HTTP/1.1 200 OK
   ```

   Минимальное исправление — Basic Auth/SSO и rate limiting в nginx. После появления браузерной авторизации добавить проверку Origin или CSRF-токен:

   ```python
   PUBLIC_ORIGIN = os.environ["PUBLIC_ORIGIN"].rstrip("/")

   @app.before_request
   def check_origin():
       if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
           origin = request.headers.get("Origin")
           if origin and origin.rstrip("/") != PUBLIC_ORIGIN:
               abort(403)
   ```

   Без аутентификации публичный запуск недопустим.

2. **Лимит 25 MB не защищает от XLSX zip bomb и исчерпания памяти**

   Места: [app.py:31–52](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:31), [converter.py:630–656](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:630), [converter.py:659–690](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:659), [converter.py:1432–1453](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1432).

   `openpyxl.load_workbook()` распаковывает и строит объектную модель **до** проверки `MAX_ROWS`. Лимит Flask относится только к сжатому HTTP-телу.

   Реальное воспроизведение:

   ```text
   XLSX на диске:       70 104 bytes
   Распакованный XML:   67 125 475 bytes
   POST /api/upload:    200
   RSS waitress:        15 344 KB → 122 144 KB
   ```

   На CSV в 100 000 строк пик достиг 325 152 KB. При этом один session-dir занял 32 MB, а разрешено 50 сессий.

   Обязательны предварительная проверка ZIP и значительно меньшие лимиты:

   ```python
   from zipfile import ZipFile, BadZipFile

   MAX_XLSX_UNPACKED = 100 * 1024 * 1024
   MAX_XLSX_RATIO = 200
   MAX_ROWS = int(os.environ.get("MAX_ROWS", 50_000))
   MAX_COLS = int(os.environ.get("MAX_COLS", 100))

   def validate_xlsx_archive(path):
       try:
           with ZipFile(path) as zf:
               infos = zf.infolist()
               unpacked = sum(i.file_size for i in infos)
               packed = sum(max(i.compress_size, 1) for i in infos)
               if unpacked > MAX_XLSX_UNPACKED:
                   raise TableTooLargeError("XLSX is too large after decompression.")
               if unpacked / max(packed, 1) > MAX_XLSX_RATIO:
                   raise TableTooLargeError("Suspicious XLSX compression ratio.")
       except BadZipFile as exc:
           raise ValueError("Invalid XLSX archive.") from exc
   ```

   Вызывать это до `load_workbook`. CSV необходимо читать построчно и прекращать после лимита; `write_csv()` — писать строки сразу, без промежуточного `safe_rows`. Если нужен приём 100k+, обработку надо выносить в отдельный worker с очередью и собственным memory limit.

3. **Сессии несовместимы с несколькими gunicorn workers; TTL фактически не работает**

   Места: [app.py:54–57](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:54), [app.py:69–97](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:69), [app.py:120–123](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:120), [app.py:160–170](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:160), [app.py:233–236](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:233).

   Состояние хранится в `SESSIONS` внутри процесса. Фактический запуск `gunicorn -w 2`:

   ```text
   20 × GET /healthz:
       7 responses:  sessions=0
      13 responses:  sessions=1

   20 × POST /api/convert:
      15 × HTTP 200
       5 × HTTP 404 "Session not found"
   ```

   TTL проверяется только при следующем `/api/upload`. При `SESSION_TTL_SECONDS=0` после ожидания:

   ```json
   {
     "after_sleep_health": {"sessions": 1, "status": "ok"},
     "convert_before_next_upload": 200,
     "download_before_next_upload": 200,
     "convert_after_next_upload": 404
   }
   ```

   Лимит сессий также ошибочен на единицу: при `MAX_SESSIONS=1` получены количества `[1, 2, 2]`.

   Минимально допустимый вариант для первого деплоя:

   ```text
   gunicorn --workers 1 --threads 1 --worker-class sync ...
   ```

   И обязательная проверка TTL при каждом обращении:

   ```python
   def get_live_session(session_id):
       now = time.time()
       with SESSIONS_LOCK:
           session = SESSIONS.get(session_id)
           if session and now - session["created"] <= SESSION_TTL_SECONDS:
               return session
           if session:
               _drop_session(session_id)
       return None
   ```

   Использовать её и в `convert`, и в `download`. Добавление сессии и ограничение числа должны выполняться под одним lock:

   ```python
   with SESSIONS_LOCK:
       while len(SESSIONS) >= MAX_SESSIONS:
           oldest = min(SESSIONS, key=lambda sid: SESSIONS[sid]["created"])
           _drop_session(oldest)
       SESSIONS[session_id] = session
   ```

   Для более одного worker состояние нужно вынести в Redis/Postgres/общий файловый metadata-store; одного `threading.Lock` недостаточно.

4. **CSV formula injection включена по умолчанию, а отчёты не защищены вообще**

   Места: [converter.py:931–943](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:931), [converter.py:1139–1154](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1139), [app.py:270–276](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:270), [index.html:78–82](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/templates/index.html:78).

   С входным значением `=1+1`:

   ```json
   {
     "default_sanitize": false,
     "raw_first_cell": "=1+1",
     "merged_first_cell": "=1+1"
   }
   ```

   Даже при `sanitize_formulas=true` основной CSV стал безопасным, но `duplicates_report.csv` остался:

   ```csv
   1;=1+1;2;2;=1+1;e-mail
   ```

   Отчёты прямо предназначены для открытия в Excel, поэтому опциональная защита здесь неприемлема.

   Исправление — всегда экранировать все пользовательские текстовые значения, включая оба отчёта:

   ```python
   def escape_for_excel(value, header):
       value = str(value or "")
       # Разрешать "+" без апострофа только после перехода
       # на проверенный digits-only E.164.
       if header in PHONE_OUTPUTS and re.fullmatch(r"\+\d+", value):
           return value
       if value[:1] in ("=", "+", "-", "@", "\t", "\r"):
           return "'" + value
       return value
   ```

   Применять внутри `write_csv()` ко всем строкам с учётом соответствующего заголовка. Пользователь не должен иметь возможность отключить защиту для Company, Comment, URL, e-mail и отчётов.

5. **Дедупликация по телефону уже объединяет разные реальные компании**

   Места: [converter.py:939–942](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:939), [converter.py:1292–1297](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1292), [index.html:118–121](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/templates/index.html:118).

   По умолчанию сравниваются последние девять цифр. В реальном результате группа №55:

   ```text
   Source rows: 294, 321
   Names:
     Supliestetica Oriental SRL.
     Dermclar República Dominicana
   Matched on: phone
   ```

   Это явно разные названия, но в `crm_import_merged.csv` они превращаются в одну запись. Сырой файл остаётся, однако UI рекомендует именно merged-файл.

   Минимальное исправление:

   ```python
   DEFAULT_OPTIONS["dedupe_by_phone"] = False
   ```

   И убрать `checked` у `opt-dup-phone`. Совпадение телефона показывать как кандидат на ручную проверку, а не как автоматический union. Если функция всё же включена — сравнивать валидный полный E.164, не последние девять цифр.

6. **Нормализация придумывает международные номера**

   Место: [converter.py:262–289](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:262).

   Любому номеру из восьми и более цифр без страны добавляется `+`. В реальном файле так изменены **108 номеров**. Например, национальный мексиканский номер может стать номером с несуществующим или чужим country code.

   Без известной страны номер нельзя превращать в E.164. Следует использовать `phonenumbers` и ISO-регион:

   ```python
   import phonenumbers

   def normalize_phone(raw, default_region=None):
       try:
           parsed = phonenumbers.parse(raw, default_region or None)
       except phonenumbers.NumberParseException:
           return raw, "invalid"

       if not phonenumbers.is_valid_number(parsed):
           return raw, "invalid"

       value = phonenumbers.format_number(
           parsed, phonenumbers.PhoneNumberFormat.E164
       )
       return value, "ok"
   ```

   Поле UI надо изменить с числового `52` на регион `MX`, `CO`, `IT` и т. п. Без региона и без исходного `+` оставлять значение неизменным и включать его в отчёт.

7. **Входные CSV тихо портятся из-за неоднозначного определения кодировки**

   Место: [converter.py:659–668](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:659).

   После UTF-8 код пробует CP1251 раньше Latin-1/CP1252. Почти любые CP1252-байты успешно декодируются как CP1251, поэтому ошибки нет — данные просто становятся другими:

   ```text
   Clínica México → Clнnica Mйxico
   José           → Josй
   ```

   Автоматически надёжно различить CP1251 и CP1252 нельзя. Нужен явный выбор input encoding перед upload:

   ```python
   INPUT_ENCODINGS = {"utf-8-sig", "utf-8", "cp1252", "cp1251"}

   def read_csv_bytes(raw_bytes, encoding):
       if encoding not in INPUT_ENCODINGS:
           raise ValueError("Unsupported input encoding.")
       try:
           text = raw_bytes.decode(encoding, errors="strict")
       except UnicodeDecodeError as exc:
           raise ValueError(
               f"File is not valid {encoding}; choose another encoding."
           ) from exc
       ...
   ```

   Детектор вроде `charset-normalizer` можно использовать только как подсказку; найденная кодировка должна показываться оператору.

8. **Невалидный API вызывает 500, а прямой запуск включает публичный debugger**

   Места: [app.py:225–241](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:225), [converter.py:1019–1026](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1019), [app.py:365–366](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:365).

   Mapping `{"x":"company_lead"}` возвращает HTML `500`. Фактический traceback:

   ```text
   File ".../app.py", line 251, in convert
     records, fixes = converter.build_records(...)

   File ".../converter.py", line 1026, in build_records
     int(k.split(':')[0])

   ValueError: invalid literal for int() with base 10: 'x'
   ```

   Прямой запуск показал:

   ```text
   Debug mode: on
   Running on all addresses (0.0.0.0)
   Debugger is active!
   Debugger PIN: ...
   ```

   Требуется строгая валидация:

   ```python
   if not request.is_json:
       return jsonify(error="Content-Type must be application/json"), 415

   data = request.get_json(silent=True)
   if not isinstance(data, dict):
       return jsonify(error="JSON object required"), 400

   mapping = data.get("mapping")
   options = data.get("options", {})
   if not isinstance(mapping, dict) or not isinstance(options, dict):
       return jsonify(error="Invalid mapping or options"), 400

   allowed_sources = set(session["labels"])
   if set(mapping) - allowed_sources:
       return jsonify(error="Unknown source column"), 400
   if any(v not in converter.TARGET_BY_ID for v in mapping.values()):
       return jsonify(error="Unknown target field"), 400
   if "company_lead" not in mapping.values():
       return jsonify(error="Company/Lead Name must be mapped"), 400
   ```

   Не принимать от клиента `source_value`; boolean options проверять именно как `bool`. Development entry point:

   ```python
   if __name__ == "__main__":
       app.run(
           host=os.environ.get("HOST", "127.0.0.1"),
           port=PORT,
           debug=False,
       )
   ```

9. **Нет воспроизводимого deployment-набора, а файлы PII создаются с широкими правами**

   Места: [requirements.txt:1–3](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/requirements.txt:1), [app.py:27–29](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:27), [app.py:134–138](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:134).

   `requirements.txt` содержит `>=`, gunicorn отсутствует. `UPLOAD_DIR` захардкожен внутри каталога приложения. Фактически создаются:

   ```text
   drwxr-xr-x uploads/<session>
   -rw-r--r-- source.xlsx
   ```

   На общей машине это опасно, если родительский каталог доступен другим пользователям.

   Минимум:

   ```text
   Flask==3.1.3
   openpyxl==3.1.5
   gunicorn==26.0.0
   ```

   Лучше отдельный `requirements.lock` с транзитивными версиями и hashes.

   Код:

   ```python
   UPLOAD_DIR = os.environ.get(
       "UPLOAD_DIR",
       os.path.join(BASE_DIR, "uploads"),
   )
   os.makedirs(UPLOAD_DIR, mode=0o700, exist_ok=True)
   ```

   После сохранения файла выставлять `0o600`; systemd unit должен иметь `UMask=0077`. Исходный код в `/opt` должен быть read-only для service user, данные — в `/var/lib/crm-import-converter/uploads`.

# ВАЖНОЕ — стоит починить скоро

- **Артефакты одной сессии перезаписываются.** Фиксированные имена заданы в [app.py:43–49](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:43), запись идёт напрямую в [app.py:266–276](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:266). После второго `/api/convert` тот же download URL изменил SHA-256 и стал возвращать другое содержимое. При параллельных запросах возможен частично записанный или чужой вариант. Нужны `conversion_id`, отдельная директория на каждую сборку, временный файл + `os.replace()` и per-session lock.

- **Колонки сверх лимита теряются молча.** [converter.py:637](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:637), [converter.py:683–688](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:683). Тест с 301 колонкой вернул `200`, 300 колонок и ни одного warning. Вместо среза должно быть:

  ```python
  if ws.max_column > MAX_COLS:
      raise TableTooLargeError(...)
  if len(data[0]) > MAX_COLS:
      raise TableTooLargeError(...)
  ```

- **Можно получить формально успешный, но бесполезный CSV.** Проверяется лишь наличие любого mapping в [app.py:238–239](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:238). Mapping только Email → Comment дал `200`, 436 строк и `Company Name=0`, `Lead Name=0`. Следует требовать `company_lead` либо отдельное подтверждение опасного режима.

- **Мердж меняет содержимое комментариев.** `_merge_group()` сначала разбивает все поля по запятым: [converter.py:1277–1278](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1277), [converter.py:1372–1384](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:1372). Фактический пример:

  ```text
  "Call Monday, urgent"
  "Call Tuesday, urgent"
  →
  "Call Monday, urgent, Call Tuesday"
  ```

  `_split_multi` надо применять только к phone/e-mail/URL; Comment и другие text-поля объединять целыми значениями.

- **Union-find даёт транзитивные объединения.** Одна строка может совпасть со второй по имени, вторая с третьей по e-mail — и все три становятся одной компанией. Для CRM безопаснее требовать два сигнала либо помечать такие цепочки только как кандидаты.

- **Читается только активный лист.** [converter.py:634–636](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:634). Остальные листы игнорируются без предупреждения. Нужно показывать список листов или отклонять multi-sheet workbook.

- **Формулы XLSX не исполняются, но могут теряться.** `data_only=True` безопасен с точки зрения выполнения кода, однако формула без cached result превращается в `None`. Нужно предупреждать о formula cells либо требовать пересохранить файл с вычисленными значениями.

- **Логирование почти отсутствует.** Ошибка чтения возвращается клиенту вместе с текстом исключения в [app.py:140–147](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:140), но не логируется. Нужны `app.logger.exception`, request ID, длительность, размеры и статусы — без содержимого полей и полного session token. Gunicorn access/error logs направить в journald.

- **`/healthz` — только liveness.** [app.py:113–117](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:113). Добавить `/readyz`, проверяющий доступность и свободное место в `UPLOAD_DIR`.

- **Скачивания допускают хранение в кэше.** `send_file()` в [app.py:355](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:355) отдаёт `no-cache`, но не `private, no-store`. Для файлов с контактами нужен `Cache-Control: private, no-store`.

- **Email regex ограничен ASCII.** [converter.py:136](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/converter.py:136). Международные адреса/IDN могут не распознаться и исчезнуть из e-mail-поля.

- **Нет HTTP/API-тестов.** Все 40 тестов проверяют `converter.py`; не проверяются Flask routes, 413, TTL, права файлов, конкурентность, кодировки входа и отчёты с формулами.

# МЕЛОЧИ

- В download-ответе фактически получилось `Content-Type: text/csv; charset=utf-8; charset=utf-8` из-за передачи charset внутри `mimetype` в [app.py:355–356](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:355). Передавать `mimetype="text/csv"` и назначать заголовок один раз.

- `.txt` разрешён в [app.py:31](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/app.py:31), но сообщение об ошибке на строке 132 его не перечисляет.

- README ориентирован в основном на Windows-команды. Systemd unit, nginx-конфиг и Linux installation отсутствуют.

- `.gitignore` не содержит `.venv/`, `.pytest_cache/`, coverage/build-артефактов.

- `healthz` раскрывает число активных сессий; небольшой, но ненужный leakage для публичного endpoint.

Проверенные положительные моменты:

- Реального path traversal через имя upload или download route не найдено.
- Коллизии имён между разными uploads практически исключены UUIDv4 и отдельными директориями.
- В JS пользовательские строки проходят через `textContent` или `escapeHtml`: [app.js:51–55](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/static/app.js:51), [app.js:132–179](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/static/app.js:132), [app.js:294–340](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/static/app.js:294). Эксплуатируемого DOM XSS не найдено.
- `openpyxl` не выполняет формулы или VBA, код не делает `eval`, `exec` и не открывает URL. Прямого RCE из XLSX не найдено; главный риск парсера — потребление ресурсов.
- Стабильного file descriptor leak не обнаружено.
- UTF-8 BOM и выбранные выходные CP1251/CP1252 реализованы работоспособно; проблема относится к определению **входной** кодировки.

# План деплоя на `ms1`

## 1. Сначала внести обязательные правки

Минимальный безопасный scope до выкладки:

1. ZIP/uncompressed-size validation, streaming CSV, меньшие row/column limits.
2. Защита формул во всех CSV.
3. `dedupe_by_phone=False`, нормализация через валидный E.164.
4. Явная входная кодировка.
5. Валидация JSON/mapping/options и `debug=False`.
6. Реальный TTL на каждом запросе, правильный session cap.
7. `UPLOAD_DIR` через environment, права `0700/0600`.
8. Pinned requirements с gunicorn.
9. HTTP-тесты для всех перечисленных негативных сценариев.

До выноса сессий из памяти использовать строго один sync worker.

## 2. Размещение

```text
/opt/crm-import-converter/current/       код, root:crm-import, 0750
/opt/crm-import-converter/venv/          venv
/var/lib/crm-import-converter/uploads/   данные, crm-import:crm-import, 0700
/etc/crm-import-converter.env            environment, root:crm-import, 0640
```

Отдельный системный пользователь без shell. Postgres сервису пока не нужен — не выдавать ему DB credentials.

`/etc/crm-import-converter.env`:

```ini
UPLOAD_DIR=/var/lib/crm-import-converter/uploads
SESSION_TTL_SECONDS=3600
MAX_SESSIONS=10
MAX_ROWS=50000
MAX_COLS=100
PUBLIC_ORIGIN=https://crm.example.com
```

## 3. systemd

`/etc/systemd/system/crm-import-converter.service`:

```ini
[Unit]
Description=CRM Import Converter
After=network.target

[Service]
Type=simple
User=crm-import
Group=crm-import
WorkingDirectory=/opt/crm-import-converter/current
EnvironmentFile=/etc/crm-import-converter.env

ExecStart=/opt/crm-import-converter/venv/bin/gunicorn \
  --workers 1 \
  --threads 1 \
  --worker-class sync \
  --bind 127.0.0.1:5011 \
  --timeout 120 \
  --graceful-timeout 30 \
  --access-logfile - \
  --error-logfile - \
  --capture-output \
  app:app

Restart=on-failure
RestartSec=3
TimeoutStopSec=40
UMask=0077

StateDirectory=crm-import-converter
StateDirectoryMode=0700
ReadWritePaths=/var/lib/crm-import-converter

NoNewPrivileges=true
PrivateTmp=true
PrivateDevices=true
ProtectSystem=strict
ProtectHome=true
ProtectKernelTunables=true
ProtectKernelModules=true
ProtectControlGroups=true
RestrictSUIDSGID=true
LockPersonality=true
CapabilityBoundingSet=
RestrictAddressFamilies=AF_INET AF_INET6 AF_UNIX

MemoryMax=512M
TasksMax=64
LimitNOFILE=1024

[Install]
WantedBy=multi-user.target
```

Порт `5011` предварительно проверить через `ss -ltnp`; наружу он не открывается.

Независимую страховочную очистку можно задать через `/etc/tmpfiles.d/crm-import-converter.conf`:

```text
d /var/lib/crm-import-converter/uploads 0700 crm-import crm-import 6h -
```

Это не заменяет проверку TTL в приложении.

## 4. nginx + HTTPS

Из-за абсолютных `/api/...`, `/download/...` и `/static/...` в [app.js:70](/Users/loschev/PycharmProjects/crm-import-audit/crm-import-converter/static/app.js:70) сервис сейчас надо размещать на **отдельном hostname с `location /`**. Под `/crm-import/` он без дополнительных правок сломается.

Пример vhost:

```nginx
limit_req_zone $binary_remote_addr zone=crm_upload:10m rate=2r/m;
limit_req_zone $binary_remote_addr zone=crm_api:10m rate=20r/m;
limit_conn_zone $binary_remote_addr zone=crm_conn:10m;

upstream crm_import_converter {
    server 127.0.0.1:5011;
    keepalive 4;
}

server {
    listen 80;
    server_name crm.example.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl http2;
    server_name crm.example.com;

    ssl_certificate     /etc/letsencrypt/live/crm.example.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/crm.example.com/privkey.pem;

    auth_basic "CRM Import";
    auth_basic_user_file /etc/nginx/.htpasswd-crm-import;

    client_max_body_size 10m;
    client_body_timeout 30s;

    add_header X-Content-Type-Options nosniff always;
    add_header Referrer-Policy no-referrer always;
    add_header X-Frame-Options DENY always;
    add_header Content-Security-Policy "default-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'" always;
    add_header Strict-Transport-Security "max-age=31536000" always;

    location = /healthz {
        auth_basic off;
        allow 127.0.0.1;
        allow ::1;
        deny all;

        proxy_pass http://crm_import_converter;
        include /etc/nginx/proxy_params;
    }

    location = /api/upload {
        limit_req zone=crm_upload burst=2 nodelay;
        limit_conn crm_conn 2;

        proxy_pass http://crm_import_converter;
        include /etc/nginx/proxy_params;
        proxy_request_buffering on;
        proxy_connect_timeout 5s;
        proxy_read_timeout 120s;
        proxy_send_timeout 120s;
    }

    location = /api/convert {
        limit_req zone=crm_api burst=5 nodelay;

        proxy_pass http://crm_import_converter;
        include /etc/nginx/proxy_params;
        proxy_connect_timeout 5s;
        proxy_read_timeout 120s;
        proxy_send_timeout 120s;
    }

    location / {
        proxy_pass http://crm_import_converter;
        include /etc/nginx/proxy_params;
        proxy_connect_timeout 5s;
        proxy_read_timeout 120s;
    }
}
```

Размер nginx и Flask должен совпадать. Если nginx ограничивает раньше приложения, настроить JSON `error_page 413`, иначе frontend попытается разобрать HTML как JSON.

## 5. Запуск и приёмка

```bash
python3.13 -m venv /opt/crm-import-converter/venv
/opt/crm-import-converter/venv/bin/pip install --require-hashes \
  -r /opt/crm-import-converter/current/requirements.lock

sudo nginx -t
sudo systemctl daemon-reload
sudo systemctl enable --now crm-import-converter
sudo systemctl reload nginx

curl --fail http://127.0.0.1:5011/healthz
journalctl -u crm-import-converter -n 100 --no-pager
```

Перед открытием DNS/TLS повторить acceptance-набор:

- `40/40 passed`;
- реальный XLSX: `436 → 355`, четыре корректных CSV;
- malformed mapping → JSON `400`, не `500`;
- CP1252 остаётся CP1252;
- формула экранируется и в result, и в reports;
- zip bomb отклоняется до `openpyxl`;
- 301 колонка отклоняется, а не обрезается;
- expired session недоступна без следующего upload;
- сервис запущен с одним worker либо уже использует общее хранилище сессий.
