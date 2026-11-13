'use strict';

(() => {
  const MAX_FILE_BYTES = 25 * 1024 * 1024;
  const MAX_VISIBLE_ROWS = 200;
  // Filled from the selected workbook so additional horizontal facility groups appear automatically.
  let FACILITIES = [];
  const el = {
    file: document.getElementById('excel-file'), status: document.getElementById('file-status'),
    message: document.getElementById('load-message'), clear: document.getElementById('clear-data'),
    mode: document.getElementById('aggregation-mode'), form: document.getElementById('search-form'),
    days: document.getElementById('order-days'), targets: document.getElementById('facility-targets'),
    facilityList: document.getElementById('facility-list'), selectAll: document.getElementById('select-all-facilities'),
    clearFacilities: document.getElementById('clear-facilities'),
    query: document.getElementById('query-input'), search: document.querySelector('.primary-button'),
    answer: document.getElementById('answer-card'), count: document.getElementById('result-count'),
    sort: document.getElementById('sort-label'), head: document.getElementById('result-head'),
    body: document.getElementById('result-body'), download: document.getElementById('download-csv'),
    toast: document.getElementById('toast')
  };
  let records = [];
  let makers = [];
  let selectedFacilities = new Set();
  let lastResults = [];
  let lastSearch = null;
  let loadSequence = 0;
  let toastTimer = 0;

  const normalize = (value) => String(value ?? '').normalize('NFKC').toLocaleLowerCase('ja-JP').replace(/[\s　]/g, '');
  const textOf = (value) => String(value ?? '').replace(/[\r\n]+/g, ' ').trim();
  function numberOf(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const source = String(value ?? '').trim().replace(/[，,\s　]/g, '');
    if (!source || /^[ー－–—-]+$/.test(source)) return 0;
    const parsed = Number(source.replace(/個|施設|ヶ所|か所/g, ''));
    return Number.isFinite(parsed) ? parsed : 0;
  }
  function indexOf(headers, names) {
    return headers.findIndex((value) => names.includes(normalize(value)));
  }
  function findHeaderRow(rows) {
    return rows.findIndex((row) => {
      const values = row.map(normalize);
      return values.includes('商品コード') && values.includes('商品名');
    });
  }

  function parseRows(rows) {
    const headerRow = findHeaderRow(rows);
    if (headerRow < 0) throw new Error('「商品コード」と「商品名」がある見出し行を見つけられませんでした。');
    const headers = rows[headerRow].map(textOf);
    const codeCol = indexOf(headers, ['商品コード']);
    const nameCol = indexOf(headers, ['商品名']);
    const makerCol = indexOf(headers, ['メーカー名', 'メーカー']);
    const specCol = indexOf(headers, ['規格']);
    const productNoCol = indexOf(headers, ['製品番号']);
    const totalCol = indexOf(headers, ['全施設合計']);
    const facilityColumns = new Map();

    headers.forEach((value, index) => {
      const header = String(value ?? '').normalize('NFKC').replace(/[\s　]/g, '');
      const normalizedHeader = normalize(header);
      const kind = normalizedHeader.endsWith('(施設)') ? 'facility'
        : normalizedHeader.endsWith('(支店)') ? 'branch'
          : normalizedHeader.endsWith('施設') ? 'facility'
            : normalizedHeader.endsWith('支店') ? 'branch' : null;
      if (!kind) return;
      const suffixLength = /\((施設|支店)\)$/.test(normalizedHeader) ? 4 : 2;
      const label = header.slice(0, -suffixLength);
      if (!label) return;
      const key = normalize(label);
      const pair = facilityColumns.get(key) || { name: label, facility: null, branch: null };
      pair[kind] = index;
      facilityColumns.set(key, pair);
    });

    const required = [['商品コード', codeCol], ['商品名', nameCol], ['メーカー名', makerCol]];
    const missing = required.filter(([, index]) => index < 0).map(([label]) => label);
    if (missing.length) throw new Error(`必須列が見つかりません：${missing.join('、')}`);
    if (!facilityColumns.size) throw new Error('施設別の列（例：施設A(施設)、施設A(支店)）が見つかりません。');

    const parsed = [];
    for (const row of rows.slice(headerRow + 1)) {
      const code = textOf(row[codeCol]);
      const name = textOf(row[nameCol]);
      if (!code && !name) continue;
      const facilityValues = {};
      for (const [, pair] of facilityColumns) {
        facilityValues[pair.name] = {
          facility: pair.facility === null ? 0 : numberOf(row[pair.facility]),
          branch: pair.branch === null ? 0 : numberOf(row[pair.branch])
        };
      }
      parsed.push({
        code, name, maker: textOf(row[makerCol]),
        spec: specCol < 0 ? '' : textOf(row[specCol]),
        productNo: productNoCol < 0 ? '' : textOf(row[productNoCol]),
        sourceTotal: totalCol < 0 ? 0 : numberOf(row[totalCol]),
        hasSourceTotal: totalCol >= 0,
        facilityValues
      });
    }
    if (!parsed.length) throw new Error('見出し行の下に商品データがありません。');
    return parsed;
  }

  function unitsAt(record, facility, mode = el.mode.value) {
    const values = record.facilityValues[facility] || { facility: 0, branch: 0 };
    if (mode === 'facility') return values.facility;
    if (mode === 'branch') return values.branch;
    return values.facility + values.branch;
  }
  function totalFor(record, mode = el.mode.value) {
    return [...new Set([...FACILITIES, ...Object.keys(record.facilityValues)])]
      .reduce((sum, facility) => sum + unitsAt(record, facility, mode), 0);
  }
  function countFor(record, mode = el.mode.value) {
    return [...new Set([...FACILITIES, ...Object.keys(record.facilityValues)])]
      .filter((facility) => selectedFacilities.has(facility))
      .filter((facility) => unitsAt(record, facility, mode) > 0).length;
  }
  function includedFacilities(record) {
    return FACILITIES.filter((facility) => selectedFacilities.has(facility) && unitsAt(record, facility) > 0);
  }
  function monthlyAverageFor(record) {
    if (!selectedFacilities.size) return 0;
    const excluded = FACILITIES.filter((facility) => !selectedFacilities.has(facility))
      .reduce((sum, facility) => sum + unitsAt(record, facility), 0);
    const base = el.mode.value === 'both' && record.hasSourceTotal ? record.sourceTotal : totalFor(record);
    return Math.max(0, base - excluded);
  }
  function daysToOrder() {
    const days = Number(el.days.value);
    return Number.isInteger(days) && days >= 1 && days <= 365 ? days : null;
  }
  function orderQuantityFor(record) {
    const days = daysToOrder();
    if (days === null) return 0;
    return Math.ceil(Math.max(0, monthlyAverageFor(record)) * days / 30 - 1e-9);
  }
  function orderFacilityFor(record) {
    return includedFacilities(record).sort((a, b) => unitsAt(record, b) - unitsAt(record, a))[0] || '';
  }
  const fmt = (value) => Number(value).toLocaleString('ja-JP', { maximumFractionDigits: 2 });

  function setReady(ready) {
    el.query.disabled = !ready;
    el.search.disabled = !ready;
    el.clear.disabled = !ready;
    el.days.disabled = !ready;
    el.targets.disabled = !ready;
    el.selectAll.disabled = !ready;
    el.clearFacilities.disabled = !ready;
    if (!ready) el.download.disabled = true;
  }
  function setStatus(message, kind = 'idle') {
    el.status.classList.toggle('is-loaded', kind === 'loaded');
    el.status.classList.toggle('is-error', kind === 'error');
    el.status.classList.toggle('has-error-dot', kind === 'error');
    el.status.querySelector('.status-dot').classList.toggle('muted', kind === 'idle');
    el.status.lastElementChild.textContent = message;
  }
  function showAnswer(title, description) {
    el.answer.querySelector('.answer-title').textContent = title;
    el.answer.querySelector('.answer-description').textContent = description;
  }
  function toast(message) {
    el.toast.textContent = message;
    el.toast.classList.add('visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.toast.classList.remove('visible'), 2800);
  }

  function clearData(resetPicker = true) {
    loadSequence += 1;
    records = [];
    makers = [];
    selectedFacilities = new Set();
    FACILITIES = [];
    lastResults = [];
    lastSearch = null;
    setReady(false);
    setStatus('ファイルが選択されていません');
    el.message.textContent = '1シート目から商品一覧を読み込みます。見出し行は自動で判定します。';
    el.count.textContent = '0件';
    el.sort.textContent = '表示順：全施設合計の降順';
    el.query.value = '';
    el.download.disabled = true;
    el.days.value = '10';
    el.facilityList.replaceChildren();
    renderHead();
    renderRows([], '検索結果がここに表示されます');
    showAnswer('Excelを読み込むと検索を始められます', '商品コード、商品名、メーカー名や施設別の使用数で絞り込めます。');
    if (resetPicker) el.file.value = '';
  }

  async function loadFile(file) {
    if (!file) return;
    el.file.value = '';
    clearData(false);
    const requestId = ++loadSequence;
    const extension = file.name.toLocaleLowerCase('ja-JP').split('.').pop();
    if (!['xlsx', 'xls'].includes(extension)) {
      setStatus('Excel形式（.xlsx / .xls）を選択してください', 'error');
      el.message.textContent = 'このファイル形式には対応していません。';
      return;
    }
    if (file.size > MAX_FILE_BYTES) {
      setStatus('ファイルが上限（25 MB）を超えています', 'error');
      el.message.textContent = 'ファイルサイズを25 MB以下にしてください。';
      return;
    }
    if (!window.XLSX) {
      setStatus('Excel読込ライブラリを利用できません', 'error');
      el.message.textContent = 'SheetJSを読み込めませんでした。インターネット接続を確認して再読み込みしてください。';
      return;
    }
    setStatus('読み込み中…');
    el.message.textContent = 'ファイルをブラウザ内で解析しています。';
    try {
      const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellFormula: false, cellText: false, cellDates: false });
      if (requestId !== loadSequence) return;
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error('ワークシートがありません。');
      const values = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true, blankrows: false });
      records = parseRows(values);
      FACILITIES = [...new Set(records.flatMap((record) => Object.keys(record.facilityValues)))].sort((a, b) => a.localeCompare(b, 'ja', { numeric: true }));
      selectedFacilities = new Set(FACILITIES);
      makers = [...new Set(records.map((record) => record.maker).filter(Boolean))];
      lastResults = [];
      lastSearch = null;
      setReady(true);
      el.days.disabled = false;
      el.targets.disabled = false;
      el.selectAll.disabled = false;
      el.clearFacilities.disabled = false;
      renderFacilityChoices();
      setStatus(`${file.name} · ${records.length.toLocaleString('ja-JP')}件`, 'loaded');
      el.message.textContent = `「${sheetName}」から${records.length.toLocaleString('ja-JP')}件の商品を読み込みました。データはページを閉じると消去されます。`;
      renderHead();
      el.query.focus();
      showAnswer('検索条件を入力してください', '施設の集計対象を切り替えると、合計数と施設数を再計算します。');
      renderRows([], '検索結果がここに表示されます');
    } catch (error) {
      if (requestId !== loadSequence) return;
      records = [];
      setReady(false);
      const message = error instanceof Error ? error.message : 'ファイルを読み込めませんでした。';
      setStatus('読み込みに失敗しました', 'error');
      el.message.textContent = message;
      toast(message);
    }
  }

  function parseQuery(original) {
    const query = normalize(original);
    const facilityMatch = FACILITIES.slice().sort((a, b) => b.length - a.length)
      .find((name) => query.includes(normalize(name)));
    const fallbackFacility = query.match(/施設([a-z0-9]+)/i);
    const facility = facilityMatch || (fallbackFacility ? `施設${fallbackFacility[1].toUpperCase()}` : null);
    const codeMatch = original.toUpperCase().match(/\b[A-Z][A-Z0-9-]{4,}\b/);
    const countMatch = query.match(/(?:使用施設数|施設数)(?:が|は)?(\d+)(?:施設|か所|ヶ所)?(?:以上|より多い)/) || query.match(/(\d+)(?:施設|か所|ヶ所)以上/);
    const countMin = countMatch ? Number(countMatch[1]) : null;
    const facilityPattern = facility ? escapeRegExp(normalize(facility)) : '施設[a-z0-9]+';
    const quantityMatch = query.match(/(?:月)?(\d+(?:\.\d+)?)(?:個|点|本|箱|セット)以上/) || (!countMatch && facility ? query.match(new RegExp(`${facilityPattern}.*?(\\d+(?:\\.\\d+)?)以上`)) : null);
    const quantityMin = quantityMatch ? Number(quantityMatch[1]) : null;
    const rankMatch = query.match(/(?:トップ|top|上位)(\d+)/i);
    const ranking = Boolean(rankMatch || /ランキング|よく使われる|よく使用|上位商品/.test(query));
    const rankLimit = rankMatch ? Math.min(Math.max(Number(rankMatch[1]), 1), 1000) : 10;
    const allTotal = /全施設合計|全施設|総合計/.test(query);
    const sortByFacility = Boolean(facility && ranking && !allTotal);
    let remaining = query;
    for (const name of FACILITIES) remaining = remaining.replace(normalize(name), ' ');
    remaining = remaining
      .replace(/施設[a-z0-9]+/g, ' ').replace(/使用施設数|施設数/g, ' ')
      .replace(/\d+(?:\.\d+)?(?:個|点|本|箱|セット|施設|か所|ヶ所)?以上/g, ' ')
      .replace(/(?:トップ|top|上位)\d+/g, ' ')
      .replace(/ランキング|よく使われる|よく使用|上位商品|全施設合計|全施設|総合計|月/g, ' ')
      .replace(/使用されている|使用される|使われている|使われる|されている|商品コード|コード|商品名|メーカー名|メーカー|製品番号|商品|製品|使用|使う|以上|で|の|が|は|を|に|と|から|まで|検索|探して|ください|お願い/g, ' ').trim();
    const makerTerms = [];
    const makerHint = /メーカー|製品/.test(query);
    for (const maker of makers) {
      const full = normalize(maker);
      const short = full.replace(/(株式会社|有限会社|合同会社|（株）|\(株\))/g, '');
      const match = [full, short].find((term) => term.length >= 2 && (remaining.includes(term) || term.includes(remaining)));
      if (match) {
        makerTerms.push(match);
        remaining = remaining.includes(match) ? remaining.replace(match, ' ') : '';
        break;
      }
    }
    if (!makerTerms.length && makerHint && remaining) {
      makerTerms.push(remaining.replace(/(製品|メーカー)$/g, '').trim());
      remaining = '';
    }
    return { original: original.trim(), facility, code: codeMatch?.[0] ?? null, countMin, quantityMin, ranking, rankLimit, sortByFacility, makerTerms, productTerm: remaining.replace(/製品|メーカー/g, '').trim() };
  }

  function matches(record, condition) {
    if (condition.code && !normalize(record.code).includes(normalize(condition.code))) return false;
    if (condition.makerTerms.some((term) => !normalize(record.maker).includes(term))) return false;
    if (condition.productTerm && !normalize(record.name).includes(condition.productTerm)) return false;
    if (condition.countMin !== null && countFor(record) < condition.countMin) return false;
    if (condition.facility && condition.quantityMin !== null && unitsAt(record, condition.facility) < condition.quantityMin) return false;
    if (!condition.facility && condition.quantityMin !== null && monthlyAverageFor(record) < condition.quantityMin) return false;
    if (condition.facility && unitsAt(record, condition.facility) <= 0) return false;
    return true;
  }

  function search(queryText) {
    if (!records.length) return toast('先にExcelファイルを読み込んでください。');
    if (!queryText.trim()) { el.query.focus(); return toast('検索条件を入力してください。'); }
    if (daysToOrder() === null) {
      el.days.focus();
      return toast('発注日数は1〜365日の整数で入力してください。');
    }
    const condition = parseQuery(queryText);
    let found = records.filter((record) => matches(record, condition));
    const score = (record) => condition.sortByFacility ? unitsAt(record, condition.facility) : monthlyAverageFor(record);
    found.sort((a, b) => score(b) - score(a) || a.code.localeCompare(b.code, 'ja'));
    if (condition.ranking) found = found.slice(0, condition.rankLimit);
    lastResults = found;
    lastSearch = condition;
    renderRows(found, '条件に一致する商品はありませんでした');
    const modeLabel = modeName();
    const facilityText = condition.facility ? `、${condition.facility}で判定` : '';
    const summary = condition.ranking
      ? `${condition.sortByFacility ? condition.facility : '計算対象の月平均出荷数'}の上位${condition.rankLimit}件を表示しています（${found.length.toLocaleString('ja-JP')}件）。`
      : `条件に一致する商品は${found.length.toLocaleString('ja-JP')}件です${facilityText}。`;
    let detail = `検索条件：「${condition.original}」 · 集計対象：${modeLabel} · 発注日数：${daysToOrder()}日 · 計算対象：${selectedFacilities.size}/${FACILITIES.length}施設`;
    const filters = [];
    if (condition.countMin !== null) filters.push(`使用施設数${condition.countMin}以上`);
    if (condition.quantityMin !== null) filters.push(`${condition.facility || '計算対象'}月平均出荷数${condition.quantityMin}以上`);
    if (filters.length) detail += ` · ${filters.join('、')}`;
    showAnswer(summary, detail);
    el.sort.textContent = `表示順：${condition.sortByFacility ? `${condition.facility}使用数` : '対象月平均出荷数'}の降順 · ${modeLabel}`;
    el.download.disabled = !found.some((record) => orderQuantityFor(record) > 0 && orderFacilityFor(record));
  }

  function modeName() { return el.mode.value === 'facility' ? '施設のみ' : el.mode.value === 'branch' ? '支店のみ' : '施設と支店の合算'; }
  function escapeRegExp(value) { return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function renderFacilityChoices() {
    el.facilityList.replaceChildren();
    if (!FACILITIES.length) {
      const note = document.createElement('span'); note.className = 'setting-help'; note.textContent = '対象施設の列がありません。'; el.facilityList.append(note); return;
    }
    const fragment = document.createDocumentFragment();
    for (const facility of FACILITIES) {
      const label = document.createElement('label'); label.className = 'facility-option';
      const checkbox = document.createElement('input'); checkbox.type = 'checkbox'; checkbox.checked = selectedFacilities.has(facility); checkbox.value = facility;
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) selectedFacilities.add(facility); else selectedFacilities.delete(facility);
        refreshCurrentSearch();
      });
      const text = document.createElement('span'); text.textContent = facility;
      label.append(checkbox, text); fragment.append(label);
    }
    el.facilityList.append(fragment);
  }
  function refreshCurrentSearch() {
    if (lastSearch) {
      if (daysToOrder() === null) {
        el.download.disabled = true;
        renderRows(lastResults, '条件に一致する商品はありませんでした');
        showAnswer('発注日数を確認してください', '1〜365日の整数を入力すると発注数量を再計算します。');
      } else search(lastSearch.original);
    } else if (records.length) {
      showAnswer('計算条件を更新しました', `${modeName()}・${daysToOrder() ?? '未入力'}日分・計算対象${selectedFacilities.size}施設で計算します。`);
    }
  }
  function renderHead() {
    el.head.replaceChildren();
    const labels = ['順位', '商品コード', '商品名', 'メーカー', '対象月平均出荷数', '発注数', '発注施設', '使用施設数', ...FACILITIES];
    labels.forEach((label, index) => {
      const th = document.createElement('th');
      th.scope = 'col';
      th.textContent = label;
      if (index === 4 || index === 5 || index === 7 || index >= 8) th.className = 'numeric';
      el.head.append(th);
    });
  }
  function renderRows(items, emptyMessage) {
    el.body.replaceChildren();
    el.count.textContent = `${items.length.toLocaleString('ja-JP')}件`;
    if (!items.length) {
      const row = document.createElement('tr'); row.className = 'empty-row';
      const cell = document.createElement('td'); cell.colSpan = 8 + FACILITIES.length; cell.textContent = emptyMessage;
      row.append(cell); el.body.append(row); return;
    }
    const fragment = document.createDocumentFragment();
    for (const [index, record] of items.slice(0, MAX_VISIBLE_ROWS).entries()) {
      const row = document.createElement('tr');
      addCell(row, String(index + 1), 'rank-number');
      addCell(row, record.code);
      const product = document.createElement('td'); product.className = 'product-name';
      const name = document.createElement('span'); name.className = 'product-title'; name.textContent = record.name || '（商品名なし）'; product.append(name);
      if (record.spec) { const spec = document.createElement('span'); spec.className = 'product-spec'; spec.textContent = record.spec; product.append(spec); }
      row.append(product);
      addCell(row, record.maker || '—');
      addCell(row, fmt(monthlyAverageFor(record)), 'numeric');
      addCell(row, fmt(orderQuantityFor(record)), 'numeric quantity-cell');
      addCell(row, orderFacilityFor(record) || '—');
      addCell(row, fmt(countFor(record)), 'numeric');
      for (const facility of FACILITIES) addCell(row, fmt(unitsAt(record, facility)), 'numeric facility-quantity');
      fragment.append(row);
    }
    el.body.append(fragment);
    if (items.length > MAX_VISIBLE_ROWS) {
      const row = document.createElement('tr'); const cell = document.createElement('td');
      cell.colSpan = 8 + FACILITIES.length; cell.className = 'table-footnote';
      cell.textContent = `${MAX_VISIBLE_ROWS}件まで表示しています。CSVには発注数量が1以上の商品を出力します。`;
      row.append(cell); el.body.append(row);
    }
  }
  function addCell(row, value, className = '') {
    const td = document.createElement('td'); td.textContent = value; if (className) td.className = className; row.append(td);
  }
  function csvCell(value) {
    let text = String(value ?? '');
    if (typeof value === 'string' && /^[\s\u0000-\u001f]*[=+\-@]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g, '""')}"`;
  }
  function downloadCsv() {
    if (!lastResults.length || !lastSearch) return;
    const header = ['商品コード', '商品名', '規格', '製品番号', 'メーカー名', '発注数', '発注施設'];
    const rows = [header];
    for (const record of lastResults) {
      const orderQuantity = orderQuantityFor(record);
      const orderFacility = orderFacilityFor(record);
      if (orderQuantity < 1 || !orderFacility) continue;
      rows.push([record.code, record.name, record.spec, record.productNo, record.maker, orderQuantity, orderFacility]);
    }
    if (rows.length === 1) return toast('発注数量が1以上の商品がありません。日数や計算対象の施設を確認してください。');
    const blob = new Blob([`\uFEFF${rows.map((row) => row.map(csvCell).join(',')).join('\r\n')}`], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const date = new Date();
    const dateStamp = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    const anchor = document.createElement('a'); anchor.href = url; anchor.download = `発注一覧_${dateStamp}.csv`; anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  el.file.addEventListener('change', (event) => loadFile(event.currentTarget.files?.[0]));
  el.clear.addEventListener('click', () => clearData());
  el.form.addEventListener('submit', (event) => { event.preventDefault(); search(el.query.value); });
  el.download.addEventListener('click', downloadCsv);
  el.mode.addEventListener('change', () => {
    renderHead();
    refreshCurrentSearch();
  });
  el.days.addEventListener('input', refreshCurrentSearch);
  el.selectAll.addEventListener('click', () => {
    selectedFacilities = new Set(FACILITIES);
    renderFacilityChoices();
    refreshCurrentSearch();
  });
  el.clearFacilities.addEventListener('click', () => {
    selectedFacilities.clear();
    renderFacilityChoices();
    refreshCurrentSearch();
  });
  document.querySelectorAll('.example-chip').forEach((button) => button.addEventListener('click', () => {
    el.query.value = button.dataset.query ?? '';
    if (records.length) search(el.query.value); else toast('先にExcelファイルを読み込んでください。');
  }));
  renderHead();
  if (!window.XLSX) el.message.textContent = 'Excel読込ライブラリの読み込みにはインターネット接続が必要です。';
})();
