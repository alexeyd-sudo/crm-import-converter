(function () {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const dzFilename = document.getElementById('dz-filename');
  const uploadForm = document.getElementById('upload-form');
  const uploadBtn = document.getElementById('upload-btn');
  const uploadError = document.getElementById('upload-error');

  const panels = {
    upload: document.getElementById('step-upload'),
    scan: document.getElementById('step-scan'),
    map: document.getElementById('step-map'),
    result: document.getElementById('step-result'),
  };
  const dots = Array.from(document.querySelectorAll('.step-dot'));

  const KIND_LABEL = {
    email: 'e-mail', phone: 'phone', url: 'website/link', text: 'text',
  };

  let state = {
    sessionId: null,
    sourceColumns: [],
    targetFields: [],
    result: null,
  };

  function showPanel(name, stepNum) {
    Object.values(panels).forEach(p => p.classList.remove('active'));
    panels[name].classList.add('active');
    dots.forEach(d => {
      const n = Number(d.dataset.step);
      d.classList.toggle('active', n === stepNum);
      d.classList.toggle('done', n < stepNum);
    });
  }

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      onFileChosen();
    }
  });
  fileInput.addEventListener('change', onFileChosen);

  function onFileChosen() {
    const f = fileInput.files[0];
    dzFilename.textContent = f ? f.name : '';
    uploadBtn.disabled = !f;
  }

  uploadForm.addEventListener('submit', async e => {
    e.preventDefault();
    const f = fileInput.files[0];
    if (!f) return;
    uploadError.textContent = '';
    uploadBtn.disabled = true;

    showPanel('scan', 2);

    const fd = new FormData();
    fd.append('file', f);

    try {
      const resp = await fetch('/api/upload', { method: 'POST', body: fd });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Upload failed');

      state.sessionId = data.session_id;
      state.sourceColumns = data.source_columns;
      state.targetFields = data.target_fields;

      renderWarnings(data.warnings);
      renderEncodings(data.encodings, (data.defaults || {}).csv_encoding);
      renderPhoneSummary(data.phone_summary, data.row_count);
      renderMappingTable(data.source_columns, data.target_fields, data.suggested_mapping || {});
      showPanel('map', 3);
    } catch (err) {
      showPanel('upload', 1);
      uploadError.textContent = err.message;
      uploadBtn.disabled = false;
    }
  });

  function renderWarnings(warnings) {
    const el = document.getElementById('warnings');
    el.innerHTML = '';
    (warnings || []).forEach(w => {
      const div = document.createElement('div');
      div.className = 'warn';
      div.textContent = w;
      el.appendChild(div);
    });
  }

  function renderEncodings(encodings, defaultId) {
    const sel = document.getElementById('opt-encoding');
    const hint = document.getElementById('enc-hint');
    if (!encodings || sel.options.length) return;
    encodings.forEach(e => {
      const opt = document.createElement('option');
      opt.value = e.id;
      opt.textContent = e.label;
      opt.dataset.hint = e.hint || '';
      sel.appendChild(opt);
    });
    sel.value = defaultId || encodings[0].id;
    const update = () => {
      const opt = sel.options[sel.selectedIndex];
      hint.textContent = opt ? opt.dataset.hint : '';
    };
    sel.addEventListener('change', update);
    update();
  }

  function renderPhoneSummary(summary, rowCount) {
    const el = document.getElementById('phone-summary');
    el.innerHTML = '';
    const info = document.createElement('div');
    info.className = 'hint';
    info.style.marginBottom = '10px';
    info.textContent = `Rows in the file: ${rowCount}.` + (summary.length
      ? ` Phone columns found (${summary.length}):`
      : ' No phone columns detected.');
    el.appendChild(info);

    summary.forEach(s => {
      const chip = document.createElement('span');
      chip.className = 'phone-chip';
      const st = s.stats;
      chip.innerHTML = `<span>${escapeHtml(s.header)}</span>` +
        (s.max_count > 1 ? `<span class="badge">up to ${s.max_count} numbers/cell</span>` : '') +
        `<span class="badge ok" title="the number already had a “+”">✓${st.ok}</span>` +
        (st.fixed ? `<span class="badge fixed" title="“+” added automatically">+added ${st.fixed}</span>` : '') +
        (st.cc_added ? `<span class="badge fixed" title="default country code applied">country code ${st.cc_added}</span>` : '') +
        (st.invalid ? `<span class="badge invalid" title="too few digits to trust">⚠ ${st.invalid}</span>` : '') +
        (st.foreign_emails ? `<span class="badge invalid">e-mails inside: ${st.foreign_emails}</span>` : '') +
        (st.foreign_urls ? `<span class="badge invalid">links inside: ${st.foreign_urls}</span>` : '');
      el.appendChild(chip);
    });
  }

  function renderMappingTable(columns, targets, suggested) {
    const el = document.getElementById('mapping-table');
    el.innerHTML = '';

    columns.forEach(col => {
      const row = document.createElement('div');
      row.className = 'mapping-row';

      const left = document.createElement('div');
      const label = document.createElement('div');
      label.className = 'col-label';
      label.textContent = col.label;

      const kind = col.kind || 'text';
      const badge = document.createElement('span');
      badge.className = 'kind-badge kind-' + kind;
      badge.textContent = KIND_LABEL[kind] || kind;
      const a = col.analysis;
      if (a && a.nonempty) {
        const parts = Object.entries(a.counts)
          .filter(([, v]) => v)
          .map(([k, v]) => `${KIND_LABEL[k] || k}: ${v}`);
        badge.title = `Non-empty cells: ${a.nonempty} of ${a.sampled}. ` + parts.join(', ');
      }
      label.appendChild(badge);
      left.appendChild(label);

      if (col.samples && col.samples.length) {
        const samples = document.createElement('div');
        samples.className = 'col-samples';
        samples.innerHTML = col.samples.map(s => `<span>${escapeHtml(s)}</span>`).join('');
        left.appendChild(samples);
      }
      row.appendChild(left);

      const select = document.createElement('select');
      select.dataset.key = col.key;
      const noneOpt = document.createElement('option');
      noneOpt.value = '';
      noneOpt.textContent = '— do not use —';
      select.appendChild(noneOpt);
      targets.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = t.label;
        select.appendChild(opt);
      });
      if (suggested[col.key]) {
        select.value = suggested[col.key];
        select.classList.add('autofilled');
        select.addEventListener('change', () => select.classList.remove('autofilled'), { once: true });
      }
      row.appendChild(select);

      el.appendChild(row);
    });
  }

  function collectOptions() {
    return {
      autofix_types: document.getElementById('opt-autofix').checked,
      skip_rows_without_name: document.getElementById('opt-skip-noname').checked,
      sanitize_formulas: document.getElementById('opt-formulas').checked,
      default_country_code: document.getElementById('opt-cc').value.trim(),
      csv_encoding: document.getElementById('opt-encoding').value,
      transliterate: document.getElementById('opt-translit').checked,
      dedupe_by_name: document.getElementById('opt-dup-name').checked,
      dedupe_by_email: document.getElementById('opt-dup-email').checked,
      dedupe_by_phone: document.getElementById('opt-dup-phone').checked,
      dedupe_by_website: document.getElementById('opt-dup-site').checked,
    };
  }

  document.getElementById('convert-btn').addEventListener('click', async () => {
    const convertError = document.getElementById('convert-error');
    convertError.textContent = '';

    const mapping = {};
    document.querySelectorAll('#mapping-table select').forEach(sel => {
      if (sel.value) mapping[sel.dataset.key] = sel.value;
    });

    if (!Object.keys(mapping).length) {
      convertError.textContent = 'Map at least one column to a CRM field.';
      return;
    }

    try {
      const resp = await fetch('/api/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_id: state.sessionId,
          mapping,
          options: collectOptions(),
        }),
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Could not build the file');

      state.result = data;
      renderResult(data);
      showPanel('result', 4);
    } catch (err) {
      convertError.textContent = err.message;
    }
  });

  function renderResult(data) {
    const summary = document.getElementById('result-summary');
    summary.innerHTML =
      `<p>Done. Source rows: <b>${data.raw.row_count}</b>, ` +
      `after merging duplicates: <b>${data.merged.row_count}</b> ` +
      `(${data.duplicates.collapsed_rows} rows collapsed into ${data.duplicates.group_count} groups).</p>` +
      `<p class="hint">Values moved between fields: <b>${data.fixes.moved}</b>, ` +
      `dropped as duplicates: <b>${data.fixes.dropped_duplicates}</b>.</p>` +
      `<p class="hint">Encoding: <b>${escapeHtml(data.encoding.label)}</b>` +
      (data.encoding.transliterate ? ', accents transliterated (é→e)' : '') + '.</p>' +
      (data.encoding.replaced
        ? `<div class="warn">Characters that did not fit the chosen encoding: ` +
          `<b>${data.encoding.replaced}</b> (rows affected: ${data.encoding.affected_rows}). ` +
          `They were replaced with the closest Latin letter or with “?”. ` +
          `Choose UTF-8 with BOM to keep every character.</div>`
        : '');

    document.getElementById('download-merged').href = data.merged.download_url;
    document.getElementById('download-raw').href = data.raw.download_url;
    document.getElementById('download-dups').href = data.duplicates.download_url;
    document.getElementById('download-fixes').href = data.fixes.download_url;

    document.getElementById('dl-merged-meta').innerHTML = fillGrid(data.merged.filled, data.merged.row_count);
    document.getElementById('dl-raw-meta').innerHTML = fillGrid(data.raw.filled, data.raw.row_count);

    renderDetails(data);
    renderPreview('merged');
  }

  function fillGrid(filled, total) {
    const cells = Object.entries(filled)
      .filter(([h]) => h !== 'Source')
      .filter(([, n]) => n > 0)
      .map(([h, n]) => `<div class="fill-cell"><b>${n}</b>${escapeHtml(h)}</div>`)
      .join('');
    return `<div class="dl-count">${total} rows</div><div class="fill-grid">${cells}</div>`;
  }

  function renderDetails(data) {
    const el = document.getElementById('result-details');
    let html = '';

    if (data.duplicates.groups.length) {
      html += '<details class="report"><summary>Merged duplicates ' +
        `(${data.duplicates.group_count}, showing the first ${data.duplicates.groups.length})</summary>` +
        '<table class="preview"><thead><tr><th>Company</th><th>Rows</th>' +
        '<th>Source rows</th><th>Matched on</th><th>Source names</th></tr></thead><tbody>';
      data.duplicates.groups.forEach(g => {
        html += `<tr><td>${escapeHtml(g.name)}</td><td>${g.size}</td>` +
          `<td>${g.rows.join(', ')}</td><td>${escapeHtml(g.matched_on.join(', '))}</td>` +
          `<td>${escapeHtml(g.names.join(' | '))}</td></tr>`;
      });
      html += '</tbody></table></details>';
    }

    if (data.fixes.items.length) {
      html += '<details class="report"><summary>Moved values ' +
        `(${data.fixes.moved + data.fixes.dropped_duplicates}, showing the first ${data.fixes.items.length})</summary>` +
        '<table class="preview"><thead><tr><th>Row</th><th>Source column</th>' +
        '<th>Type</th><th>Value</th><th>Moved to</th></tr></thead><tbody>';
      data.fixes.items.forEach(f => {
        const to = f.action === 'duplicate' ? '<i>duplicate, dropped</i>' : escapeHtml(f.to);
        html += `<tr><td>${f.row}</td><td>${escapeHtml(f.column)}</td>` +
          `<td>${escapeHtml(KIND_LABEL[f.kind] || f.kind)}</td>` +
          `<td>${escapeHtml(f.value)}</td><td>${to}</td></tr>`;
      });
      html += '</tbody></table></details>';
    }

    el.innerHTML = html;
  }

  function renderPreview(which) {
    const data = state.result;
    if (!data) return;
    const previewEl = document.getElementById('result-preview');
    const rows = data[which].preview;
    if (!rows.length) { previewEl.innerHTML = ''; return; }
    let html = '<table class="preview"><thead><tr>' +
      data.headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr></thead><tbody>';
    rows.forEach(r => {
      html += '<tr>' + r.map(c => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>';
    });
    html += '</tbody></table>';
    previewEl.innerHTML = html;
  }

  document.querySelectorAll('.preview-switch .btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.preview-switch .btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderPreview(btn.dataset.preview);
    });
  });

  document.getElementById('restart-btn').addEventListener('click', () => {
    state = { sessionId: null, sourceColumns: [], targetFields: [], result: null };
    fileInput.value = '';
    dzFilename.textContent = '';
    uploadBtn.disabled = true;
    uploadError.textContent = '';
    showPanel('upload', 1);
  });

  function escapeHtml(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
})();
