/*
 * 医療材料売上検索システム
 * Excelの読込と検索はブラウザ内で完結します。データ永続化・外部送信はありません。
 */
'use strict';

(() => {
  const MAX_FILE_BYTES = 25 * 1024 * 1024;
  const MAX_VISIBLE_ROWS = 200;
  const DEFAULT_RANK_LIMIT = 10;

  const el = {
    file: document.getElementById('excel-file'),
    fileStatus: document.getElementById('file-status'),
    loadMessage: document.getElementById('load-message'),
    clear: document.getElementById('clear-data'),
    form: document.getElementById('search-form'),
    query: document.getElementById('query-input'),
    searchButton: document.querySelector('.primary-button'),
    answer: document.getElementById('answer-card'),
    resultCount: document.getElementById('result-count'),
    sortLabel: document.getElementById('sort-label'),
    resultBody: document.getElementById('result-body'),
    download: document.getElementById('download-csv'),
    toast: document.getElementById('toast')
  };

  let records = [];
  let lastResults = [];
  let lastSearch = null;
  let toastTimer = 0;
  let loadSequence = 0;

  function normalize(value) {
    return String(value ?? '').normalize('NFKC').toLocaleLowerCase('ja-JP').replace(/[\s　]/g, '');
  }

  function cellText(value) {
    return String(value ?? '').replace(/[\r\n]+/g, ' ').trim();
  }

  function numeric(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = String(value ?? '').trim().replace(/[，,\s　]/g, '');
    if (!text || /^[ー－–—-]+$/.test(text)) return 0;
    const parsed = Number(text.replace(/個|施設|ヶ所|か所/g, ''));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function colIndex(headers, predicate) {
    return headers.findIndex((value) => predicate(normalize(value)));
  }

  function isHeaderRow(row) {
    const values = row.map(normalize);
    return values.includes('商品コード') && values.includes('商品名');
  }

  function headerInfo(headers) {
    const code = colIndex(headers, (h) => h === '商品コード');
    const name = colIndex(headers, (h) => h === '商品名');
    const maker = colIndex(headers, (h) => h === 'メーカー名' || h === 'メーカー');
    const spec = colIndex(headers, (h) => h === '規格');
    const productNo = colIndex(headers, (h) => h === '製品番号');
    const total = colIndex(headers, (h) => h === '全施設合計');
    const usedCount = colIndex(headers, (h) => h === '使用施設数');
    const facilities = new Map();

    headers.forEach((value, index) => {
      const header = normalize(value).replace(/[()（）]/g, '');
      const match = header.match(/^(施設[A-ZＡ-Ｚ0-9０-９]+)(施設|支店)$/i);
      if (!match) return;
      const facility = match[1].toLocaleUpperCase('ja-JP').replace(/[Ａ-Ｚ]/g, (letter) => String.fromCharCode(letter.charCodeAt(0) - 0xFEE0));
      const item = facilities.get(facility) || { facility: null, branch: null };
      if (match[2] === '施設') item.facility = index;
      if (match[2] === '支店') item.branch = index;
      facilities.set(facility, item);
    });

    return { code, name, maker, spec, productNo, total, usedCount, facilities };
  }

  function parseRows(aoa) {
    const headerRowIndex = aoa.findIndex(isHeaderRow);
    if (headerRowIndex < 0) {
      throw new Error('「商品コード」と「商品名」がある見出し行を見つけられませんでした。');
    }

    const headers = aoa[headerRowIndex].map(cellText);
    const columns = headerInfo(headers);
    const required = [
      ['商品コード', columns.code], ['商品名', columns.name],
      ['メーカー名', columns.maker], ['全施設合計', columns.total],
      ['使用施設数', columns.usedCount]
    ];
    const missing = required.filter(([, index]) => index < 0).map(([label]) => label);
    if (missing.length) throw new Error(`必須列が見つかりません：${missing.join('、')}`);
    if (!columns.facilities.size) throw new Error('施設別の列（例：施設A(施設)、施設A(支店)）が見つかりません。');

    const parsed = [];
    for (const row of aoa.slice(headerRowIndex + 1)) {
      const code = cellText(row[columns.code]);
      const name = cellText(row[columns.name]);
      if (!code && !name) continue;
      const facilityUnits = {};
      for (const [facility, pair] of columns.facilities) {
        facilityUnits[facility] = numeric(row[pair.facility]) + numeric(row[pair.branch]);
      }
      parsed.push({
        code,
        name,
        maker: cellText(row[columns.maker]),
        spec: columns.spec >= 0 ? cellText(row[columns.spec]) : '',
        productNo: columns.productNo >= 0 ? cellText(row[columns.productNo]) : '',
        total: numeric(row[columns.total]),
        usedCount: numeric(row[columns.usedCount]),
        facilityUnits
      });
    }
    if (!parsed.length) throw new Error('見出しの下に商品データがありません。');
    return parsed;
  }

  function showToast(message) {
    el.toast.textContent = message;
    el.toast.classList.add('visible');
    window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => el.toast.classList.remove('visible'), 2800);
  }

  function setReady(ready) {
    el.query.disabled = !ready;
    el.searchButton.disabled = !ready;
    el.clear.disabled = !ready;
    if (!ready) el.download.disabled = true;
  }

  function setFileStatus(text, kind) {
    el.fileStatus.classList.toggle('is-loaded', kind === 'loaded');
    el.fileStatus.classList.toggle('is-error', kind === 'error');
    const dot = el.fileStatus.querySelector('.status-dot');
    dot.classList.toggle('muted', kind === 'idle');
    el.fileStatus.classList.toggle('has-error-dot', kind === 'error');
    el.fileStatus.lastElementChild.textContent = text;
  }

  async function loadFile(file) {
    if (!file) return;
    el.file.value = '';
    clearData(false);
    const loadId = ++loadSequence;
    const extension = file.name.toLocaleLowerCase('ja-JP').split('.').pop();
    if (!['xlsx', 'xls'].includes(extension)) {
      setFileStatus('Excel形式（.xlsx / .xls）を選択してください', 'error');
      el.loadMessage.textContent = 'このファイル形式には対応していません。';
      return;
    }
    if (file.size > MAX_FILE_BYTES) {
      setFileStatus('ファイルが上限（25 MB）を超えています', 'error');
      el.loadMessage.textContent = 'ファイルサイズを25 MB以下にしてください。';
      return;
    }
    if (!window.XLSX) {
      setFileStatus('Excel読込ライブラリを利用できません', 'error');
      el.loadMessage.textContent = 'SheetJSを読み込めませんでした。インターネット接続を確認して再読み込みしてください。';
      return;
    }
    setFileStatus('読み込み中…', 'idle');
    el.loadMessage.textContent = 'ファイルをブラウザ内で解析しています。';
    try {
      const buffer = await file.arrayBuffer();
      if (loadId !== loadSequence) return;
      const workbook = window.XLSX.read(buffer, { type: 'array', cellFormula: false, cellText: false, cellDates: false });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) throw new Error('ワークシートがありません。');
      const worksheet = workbook.Sheets[firstSheetName];
      const aoa = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true, blankrows: false });
      records = parseRows(aoa);
      uniqueMakers = [...new Set(records.map((record) => record.maker).filter(Boolean))];
      lastResults = [];
      lastSearch = null;
      setReady(true);
      setFileStatus(`${file.name} · ${records.length.toLocaleString('ja-JP')}件`, 'loaded');
      el.loadMessage.textContent = `「${firstSheetName}」から${records.length.toLocaleString('ja-JP')}件の商品を読み込みました。データはページを閉じると消去されます。`;
      el.query.focus();
      showAnswer('検索条件を入力してください', '自然文の検索に対応しています。検索結果は全施設合計の多い順に表示します。');
      renderResults([], '検索結果がここに表示されます');
    } catch (error) {
      if (loadId !== loadSequence) return;
      records = [];
      setReady(false);
      const message = error instanceof Error ? error.message : 'ファイルを読み込めませんでした。';
      setFileStatus('読み込みに失敗しました', 'error');
      el.loadMessage.textContent = message;
      showToast(message);
    }
  }

  function clearData(resetPicker = true) {
    loadSequence += 1;
    records = [];
    lastResults = [];
    lastSearch = null;
    setReady(false);
    setFileStatus('ファイルが選択されていません', 'idle');
    el.loadMessage.textContent = '1シート目から商品一覧を読み込みます。見出し行は自動で判定します。';
    el.resultCount.textContent = '0件';
    el.sortLabel.textContent = '表示順：全施設合計の降順';
    el.query.value = '';
    renderResults([], '検索結果がここに表示されます');
    showAnswer('Excelを読み込むと検索を始められます', '商品コード、商品名、メーカー名や施設別の使用数で絞り込めます。');
    if (resetPicker) el.file.value = '';
  }

  function showAnswer(title, description) {
    const titleNode = el.answer.querySelector('.answer-title');
    const descNode = el.answer.querySelector('.answer-description');
    titleNode.textContent = title;
    descNode.textContent = description;
  }

  function parseQuery(original) {
    const query = normalize(original);
    const facilityMatches = [...query.matchAll(/施設([a-eａ-ｅ])/gi)];
    const facility = facilityMatches.length
      ? `施設${facilityMatches[0][1].toUpperCase().replace(/[Ａ-Ｅ]/g, (letter) => String.fromCharCode(letter.charCodeAt(0) - 0xFEE0))}`
      : null;
    const codeMatch = original.toUpperCase().match(/\b[A-Z][A-Z0-9-]{4,}\b/);

    const countMatch = query.match(/(?:使用施設数|施設数)(?:が|は)?(\d+)(?:施設|か所|ヶ所)?(?:以上|より多い)/)
      || query.match(/(\d+)(?:施設|か所|ヶ所)以上/);
    const countMin = countMatch ? Number(countMatch[1]) : null;
    const quantityMatch = query.match(/(?:月)?(\d+(?:\.\d+)?)(?:個|点|本|箱|セット)以上/)
      || (!countMatch && facility ? query.match(/施設[a-e].*?(\d+(?:\.\d+)?)以上/) : null);
    const quantityMin = quantityMatch ? Number(quantityMatch[1]) : null;
    const rankMatch = query.match(/(?:トップ|top|上位)(\d+)/i);
    const ranking = Boolean(rankMatch || /ランキング|よく使われる|よく使用|上位商品/.test(query));
    const rankLimit = rankMatch ? Math.min(Math.max(Number(rankMatch[1]), 1), 1000) : DEFAULT_RANK_LIMIT;
    const allTotal = /全施設合計|全施設|総合計/.test(query);
    const sortByFacility = Boolean(facility && ranking && !allTotal);

    let remaining = query
      .replace(/施設[a-e]/g, ' ')
      .replace(/使用施設数|施設数/g, ' ')
      .replace(/\d+(?:\.\d+)?(?:個|点|本|箱|セット|施設|か所|ヶ所)?以上/g, ' ')
      .replace(/(?:トップ|top|上位)\d+/g, ' ')
      .replace(/ランキング|よく使われる|よく使用|上位商品|全施設合計|全施設|総合計|月/g, ' ')
      .replace(/使用されている|使用される|使われている|使われる|されている|商品コード|コード|商品名|メーカー名|メーカー|製品番号|商品|製品|使用|使う|以上|で|の|が|は|を|に|と|から|まで|検索|探して|ください|お願い/g, ' ')
      .trim();

    const makerHint = /(?:メーカー|製品|メーカー名)/.test(query);
    const makerTerms = [];
    if (remaining) {
      for (const maker of uniqueMakers) {
        const candidate = normalize(maker);
        const shortName = candidate.replace(/(株式会社|有限会社|合同会社|（株）|\(株\))/g, '');
        const matched = [candidate, shortName].find((term) => term.length >= 2 && (remaining.includes(term) || term.includes(remaining)));
        if (matched) {
          makerTerms.push(matched);
          remaining = remaining.includes(matched) ? remaining.replace(matched, ' ') : '';
          break;
        }
      }
      if (!makerTerms.length && makerHint) {
        const term = remaining.replace(/(製品|メーカー)$/g, '').trim();
        if (term) makerTerms.push(term);
        remaining = '';
      }
    }
    const productTerm = remaining.replace(/製品|メーカー/g, '').trim();
    return { query: original.trim(), facility, code: codeMatch?.[0] ?? null, countMin, quantityMin, ranking, rankLimit, sortByFacility, makerTerms, productTerm };
  }

  let uniqueMakers = [];

  function recordMatches(record, conditions) {
    if (conditions.code && !normalize(record.code).includes(normalize(conditions.code))) return false;
    if (conditions.makerTerms.some((term) => !normalize(record.maker).includes(term))) return false;
    if (conditions.productTerm && !normalize(record.name).includes(conditions.productTerm)) return false;
    if (conditions.countMin !== null && record.usedCount < conditions.countMin) return false;
    if (conditions.facility && conditions.quantityMin !== null && (record.facilityUnits[conditions.facility] ?? 0) < conditions.quantityMin) return false;
    if (!conditions.facility && conditions.quantityMin !== null && record.total < conditions.quantityMin) return false;
    if (conditions.facility && conditions.quantityMin === null && (record.facilityUnits[conditions.facility] ?? 0) <= 0) return false;
    return true;
  }

  function search(original) {
    if (!records.length) {
      showToast('先にExcelファイルを読み込んでください。');
      return;
    }
    if (!original.trim()) {
      showToast('検索条件を入力してください。');
      el.query.focus();
      return;
    }

    const conditions = parseQuery(original);
    let matches = records.filter((record) => recordMatches(record, conditions));
    const score = (record) => conditions.sortByFacility
      ? (record.facilityUnits[conditions.facility] ?? 0)
      : record.total;
    matches.sort((a, b) => score(b) - score(a) || a.code.localeCompare(b.code, 'ja'));
    if (conditions.ranking) matches = matches.slice(0, conditions.rankLimit);

    lastResults = matches;
    lastSearch = conditions;
    renderResults(matches, '条件に一致する商品はありませんでした');

    const facilityDescription = conditions.facility ? `、${conditions.facility}の施設・支店合計で判定` : '';
    const countDescription = matches.length.toLocaleString('ja-JP');
    const filteredWord = conditions.ranking ? '該当商品の上位' : '条件に一致する商品';
    let summary = `${filteredWord}は${countDescription}件です${facilityDescription}。`;
    if (conditions.ranking) summary = `${conditions.sortByFacility ? conditions.facility : '全施設合計'}の使用数上位${conditions.rankLimit}件を表示しています（該当${countDescription}件）。`;
    let detail = `検索条件：「${conditions.query}」`;
    const filterDescriptions = [];
    if (conditions.makerTerms.length) filterDescriptions.push('メーカー名');
    if (conditions.productTerm) filterDescriptions.push('商品名');
    if (conditions.code) filterDescriptions.push('商品コード');
    if (conditions.countMin !== null) filterDescriptions.push(`使用施設数 ${conditions.countMin}以上`);
    if (conditions.quantityMin !== null) filterDescriptions.push(`${conditions.facility || '全施設'}使用数 ${conditions.quantityMin}以上`);
    if (filterDescriptions.length) detail += ` · 適用条件：${filterDescriptions.join('、')}`;
    showAnswer(summary, detail);
    el.sortLabel.textContent = `表示順：${conditions.sortByFacility ? `${conditions.facility}使用数` : '全施設合計'}の降順`;
    el.download.disabled = matches.length === 0;
  }

  function renderResults(items, emptyMessage) {
    el.resultBody.replaceChildren();
    el.resultCount.textContent = `${items.length.toLocaleString('ja-JP')}件`;
    if (!items.length) {
      const row = document.createElement('tr');
      row.className = 'empty-row';
      const cell = document.createElement('td');
      cell.colSpan = 6;
      cell.textContent = emptyMessage;
      row.append(cell);
      el.resultBody.append(row);
      return;
    }

    const fragment = document.createDocumentFragment();
    for (const [index, record] of items.slice(0, MAX_VISIBLE_ROWS).entries()) {
      const row = document.createElement('tr');
      const rank = document.createElement('td');
      rank.className = 'rank-number';
      rank.textContent = String(index + 1);
      row.append(rank);
      appendCell(row, record.code);
      const productCell = document.createElement('td');
      productCell.className = 'product-name';
      const productName = document.createElement('span');
      productName.className = 'product-title';
      productName.textContent = record.name || '（商品名なし）';
      productCell.append(productName);
      if (record.spec) {
        const specification = document.createElement('span');
        specification.className = 'product-spec';
        specification.textContent = record.spec;
        productCell.append(specification);
      }
      row.append(productCell);
      appendCell(row, record.maker || '—');
      appendCell(row, formatNumber(record.total), 'numeric quantity-cell');
      appendCell(row, formatNumber(record.usedCount), 'numeric');
      fragment.append(row);
    }
    el.resultBody.append(fragment);
    if (items.length > MAX_VISIBLE_ROWS) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 6;
      cell.className = 'table-footnote';
      cell.textContent = `${MAX_VISIBLE_ROWS}件まで表示しています。CSVには${items.length.toLocaleString('ja-JP')}件すべてを出力します。`;
      row.append(cell);
      el.resultBody.append(row);
    }
  }

  function appendCell(row, text, className = '') {
    const cell = document.createElement('td');
    cell.textContent = text;
    if (className) cell.className = className;
    row.append(cell);
  }

  function formatNumber(value) {
    return Number(value).toLocaleString('ja-JP', { maximumFractionDigits: 2 });
  }

  function csvEscape(value) {
    let text = String(value ?? '');
    if (typeof value === 'string' && /^[\s\u0000-\u001f]*[=+\-@]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g, '""')}"`;
  }

  function downloadCsv() {
    if (!lastResults.length || !lastSearch) return;
    const rows = [
      ['商品コード', '商品名', '規格', '製品番号', 'メーカー名', '全施設合計', '使用施設数', ...[...new Set(records.flatMap((row) => Object.keys(row.facilityUnits)))]]
    ];
    const facilityNames = rows[0].slice(7);
    for (const record of lastResults) {
      rows.push([
        record.code, record.name, record.spec, record.productNo, record.maker,
        record.total, record.usedCount, ...facilityNames.map((facility) => record.facilityUnits[facility] ?? 0)
      ]);
    }
    const csv = `\uFEFF${rows.map((row) => row.map(csvEscape).join(',')).join('\r\n')}`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `検索結果_${new Date().toISOString().slice(0, 10)}.csv`;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  el.file.addEventListener('change', (event) => loadFile(event.currentTarget.files?.[0]));
  el.clear.addEventListener('click', () => clearData());
  el.form.addEventListener('submit', (event) => {
    event.preventDefault();
    search(el.query.value);
  });
  el.download.addEventListener('click', downloadCsv);
  document.querySelectorAll('.example-chip').forEach((button) => {
    button.addEventListener('click', () => {
      el.query.value = button.dataset.query ?? '';
      if (records.length) search(el.query.value);
      else showToast('先にExcelファイルを読み込んでください。');
    });
  });

  if (!window.XLSX) {
    el.loadMessage.textContent = 'Excel読込ライブラリを読み込み中です。ネットワーク接続が必要です。';
  }
})();
