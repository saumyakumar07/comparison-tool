// ─── State ───────────────────────────────────────────────
let sheets = { 1: null, 2: null };
let results = [];
let activeFilter = 'all';

// ─── File Loading ─────────────────────────────────────────
function loadFile(num, file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (data.length < 2) { toast('Sheet ' + num + ' appears to be empty.', 'error'); return; }
      sheets[num] = data;
      populateHeaders(num, data[0]);
      document.getElementById('fname' + num).textContent = '✓ ' + file.name;
      document.getElementById('fname' + num).classList.remove('hidden');
      document.getElementById('drop' + num).classList.add('loaded');
      checkReady();
      updateSteps();
      toast('Sheet ' + num + ' loaded — ' + (data.length - 1) + ' rows', 'success');
    } catch (err) {
      toast('Failed to read file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateHeaders(num, headers) {
  const selIds = num === 1
    ? ['s1-id', 's1-q', 's1-r']
    : ['s2-q', 's2-r'];

  selIds.forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    sel.innerHTML = '<option value="">— select —</option>';
    headers.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = h || 'Column ' + colLetter(i);
      sel.appendChild(opt);
    });
  });

  if (num === 1) {
    const idSel = document.getElementById('s1-id');
    const qSel  = document.getElementById('s1-q');
    const rSel  = document.getElementById('s1-r');
    autoSelect(idSel, headers, ['id', 'code', 'no', 'num', 'sr', '#', 'ref']);
    autoSelect(qSel,  headers, ['question', 'q', 'query', 'item', 'description', 'text']);
    autoSelect(rSel,  headers, ['response', 'answer', 'result', 'reply', 'status', 'value']);
    if (idSel.value === '') idSel.value = '0';
    if (qSel.value  === '') qSel.value  = headers.length > 1 ? '1' : '0';
    if (rSel.value  === '') rSel.value  = headers.length > 2 ? '2' : '1';
  } else {
    const qSel = document.getElementById('s2-q');
    const rSel = document.getElementById('s2-r');
    autoSelect(qSel, headers, ['question', 'q', 'query', 'item', 'description', 'text']);
    autoSelect(rSel, headers, ['response', 'answer', 'result', 'reply', 'status', 'value']);
    if (qSel.value === '') qSel.value = headers.length > 1 ? '1' : '0';
    if (rSel.value === '') rSel.value = headers.length > 2 ? '2' : '1';
  }
}

function autoSelect(sel, headers, keywords) {
  for (const kw of keywords) {
    const idx = headers.findIndex(h => String(h).toLowerCase().includes(kw));
    if (idx !== -1) { sel.value = idx; return; }
  }
}

function colLetter(i) {
  let s = '';
  i++;
  while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

function checkReady() {
  const ready = sheets[1] && sheets[2];
  document.getElementById('process-btn').disabled = !ready;
}

// ─── Drag & Drop ──────────────────────────────────────────
function onDragOver(e, num) {
  e.preventDefault();
  document.getElementById('drop' + num).classList.add('drag-over');
}
function onDragLeave(e, num) {
  document.getElementById('drop' + num).classList.remove('drag-over');
}
function onDrop(e, num) {
  e.preventDefault();
  document.getElementById('drop' + num).classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) loadFile(num, file);
}

// ─── Text Normalization ───────────────────────────────────
function normalize(text) {
  return String(text ?? '').trim().toLowerCase().replace(/\s+/g, ' ').replace(/['']/g, "'").replace(/[""]/g, '"');
}

// Bigram Dice coefficient — O(n), safe for 1000+ rows
function getBigrams(s) {
  const map = new Map();
  for (let i = 0; i < s.length - 1; i++) {
    const b = s[i] + s[i + 1];
    map.set(b, (map.get(b) || 0) + 1);
  }
  return map;
}

function similarityScore(a, b) {
  a = normalize(a); b = normalize(b);
  if (a === b) return 100;
  if (a.length < 2 || b.length < 2) return 0;
  const ba = getBigrams(a), bb = getBigrams(b);
  let intersection = 0;
  for (const [bg, cnt] of ba) {
    if (bb.has(bg)) intersection += Math.min(cnt, bb.get(bg));
  }
  return Math.round(2 * intersection / (a.length - 1 + b.length - 1) * 100);
}

// ─── Main Processing ──────────────────────────────────────
function setProgress(pct, label, status) {
  document.getElementById('proc-fill').style.width = pct + '%';
  if (label) document.getElementById('proc-label').textContent = label;
  if (status) document.getElementById('proc-status').textContent = status;
}

function yield_() { return new Promise(r => setTimeout(r, 0)); }

function extractDomain(qid, delim) {
  const s = String(qid ?? '').trim();
  if (!delim || !s) return s || 'Unknown';
  const idx = s.indexOf(delim);
  return idx > 0 ? s.slice(0, idx).trim() : s;
}

async function processSheets() {
  const s1idIdx = parseInt(document.getElementById('s1-id').value);
  const s1qIdx  = parseInt(document.getElementById('s1-q').value);
  const s1rIdx  = parseInt(document.getElementById('s1-r').value);
  const s2qIdx  = parseInt(document.getElementById('s2-q').value);
  const s2rIdx  = parseInt(document.getElementById('s2-r').value);
  const domainDelim = document.getElementById('domain-delim').value;

  if (isNaN(s1qIdx) || isNaN(s1rIdx) || isNaN(s2qIdx) || isNaN(s2rIdx)) {
    toast('Please select Question and Response columns for both sheets.', 'error'); return;
  }

  // Lock UI
  const btn = document.getElementById('process-btn');
  btn.disabled = true;
  btn.textContent = '⏳ Processing...';
  document.getElementById('processing-overlay').classList.remove('hidden');
  document.getElementById('results-section').classList.add('hidden');
  setProgress(0, 'Building lookup map...', 'Reading Sheet 2');
  await yield_();

  const useFuzzy = document.getElementById('opt-fuzzy').checked;
  const threshold = parseInt(document.getElementById('opt-threshold').value);

  const data1 = sheets[1].slice(1).filter(r => r.some(c => c !== ''));
  const data2 = sheets[2].slice(1).filter(r => r.some(c => c !== ''));

  // Build lookup map from Sheet 2
  const map2 = new Map();
  const fuzzyKeys = []; // [{key, bigrams, entry}]
  data2.forEach(row => {
    const q = row[s2qIdx] ?? '';
    const r = row[s2rIdx] ?? '';
    const key = normalize(q);
    if (!map2.has(key)) {
      const entry = { original: q, response: r };
      map2.set(key, entry);
      if (useFuzzy) fuzzyKeys.push({ key, bigrams: getBigrams(key), entry });
    }
  });

  setProgress(5, 'Processing questions...', `0 / ${data1.length}`);
  await yield_();

  results = [];
  const CHUNK = 40;

  for (let i = 0; i < data1.length; i += CHUNK) {
    const chunk = data1.slice(i, i + CHUNK);

    chunk.forEach(row => {
      const q1 = row[s1qIdx] ?? '';
      const r1 = row[s1rIdx] ?? '';
      const normQ1 = normalize(q1);
      let matchedEntry = null;
      let matchScore = 0;

      // Exact match first (O(1))
      if (map2.has(normQ1)) {
        matchedEntry = map2.get(normQ1);
        matchScore = 100;
      } else if (useFuzzy && normQ1.length >= 2) {
        // Fuzzy: bigram Dice against pre-computed keys
        const qBigrams = getBigrams(normQ1);
        let bestScore = 0, bestEntry = null;
        for (const fk of fuzzyKeys) {
          const lenRatio = Math.min(normQ1.length, fk.key.length) / Math.max(normQ1.length, fk.key.length);
          if (lenRatio < (threshold - 10) / 100) continue;
          let intersection = 0;
          for (const [bg, cnt] of qBigrams) {
            if (fk.bigrams.has(bg)) intersection += Math.min(cnt, fk.bigrams.get(bg));
          }
          const score = Math.round(2 * intersection / (normQ1.length - 1 + fk.key.length - 1) * 100);
          if (score > bestScore) { bestScore = score; bestEntry = fk.entry; }
          if (bestScore === 100) break;
        }
        if (bestScore >= threshold) { matchedEntry = bestEntry; matchScore = bestScore; }
      }

      const r2Raw = matchedEntry ? matchedEntry.response : null;
      // Treat 'Not Attempted' in manual response as 'No'
      const r2 = r2Raw !== null && normalize(r2Raw) === 'not attempted' ? 'No' : r2Raw;
      let status;
      if (r2 === null) status = 'unmatched';
      else if (normalize(r1) === normalize(r2)) status = 'match';
      else status = 'mismatch';

      const qid = !isNaN(s1idIdx) ? (row[s1idIdx] ?? '') : '';
      const domain = extractDomain(qid, domainDelim);
      results.push({
        num: results.length + 1,
        question: q1, aiResponse: r1,
        manualResponse: r2 ?? '—',
        status, matchScore,
        qid: String(qid), domain
      });
    });

    const pct = Math.min(95, 5 + Math.round((i + CHUNK) / data1.length * 90));
    setProgress(pct, 'Processing questions...', `${Math.min(i + CHUNK, data1.length)} / ${data1.length}`);
    await yield_();
  }

  setProgress(100, 'Done!', `${results.length} rows processed`);
  await yield_();

  // Restore UI
  btn.disabled = false;
  btn.textContent = '⚡ Generate Comparison';
  document.getElementById('processing-overlay').classList.add('hidden');

  renderResults();
  updateSteps();
}

// ─── Render ───────────────────────────────────────────────
function renderResults() {
  const total = results.length;
  const matched = results.filter(r => r.status === 'match').length;
  const mismatched = results.filter(r => r.status === 'mismatch').length;
  const unmatched = results.filter(r => r.status === 'unmatched').length;

  document.getElementById('stat-total').textContent = total;
  document.getElementById('stat-matched').textContent = matched;
  document.getElementById('stat-mismatched').textContent = mismatched;
  document.getElementById('stat-unmatched').textContent = unmatched;

  const pct = v => total ? Math.round(v / total * 100) : 0;
  ['matched', 'mismatched', 'unmatched'].forEach(k => {
    const val = k === 'matched' ? matched : k === 'mismatched' ? mismatched : unmatched;
    const p = pct(val);
    document.getElementById('pct-' + k).textContent = p + '%';
    document.getElementById('bar-' + k).style.width = p + '%';
  });

  // Populate domain filter dropdown
  const domains = [...new Set(results.map(r => r.domain))].sort();
  const domSel = document.getElementById('domain-filter');
  domSel.innerHTML = '<option value="">All Domains</option>';
  domains.forEach(d => {
    const opt = document.createElement('option');
    opt.value = d; opt.textContent = d;
    domSel.appendChild(opt);
  });

  document.getElementById('results-section').classList.remove('hidden');
  applyFilter();
  toast(`Done! ${total} rows across ${domains.length} domain(s).`, 'success');
}

function applyFilter() {
  const filter = activeFilter;
  const search = (document.getElementById('search-box').value || '').toLowerCase();
  const domainFilter = document.getElementById('domain-filter').value;
  const tbody = document.getElementById('result-tbody');
  tbody.innerHTML = '';
  let shown = 0;

  results.forEach(row => {
    if (filter !== 'all' && row.status !== filter) return;
    if (domainFilter && row.domain !== domainFilter) return;
    if (search && !row.question.toLowerCase().includes(search)) return;
    shown++;

    const tr = document.createElement('tr');
    if (row.status === 'mismatch') tr.classList.add('row-mismatch');
    if (row.status === 'unmatched') tr.classList.add('row-unmatched');

    const statusHtml = row.status === 'match'
      ? `<span class="badge badge-match">✓ Match</span>`
      : row.status === 'mismatch'
        ? `<span class="badge badge-mismatch">✗ Mismatch</span>`
        : `<span class="badge badge-unmatched">? Unmatched</span>`;

    tr.innerHTML = `
      <td class="row-num">${row.num}</td>
      <td><span class="badge badge-neutral">${escHtml(row.domain)}</span></td>
      <td style="font-size:0.78rem;color:var(--text-muted);white-space:nowrap">${escHtml(row.qid)}</td>
      <td class="q-col">${escHtml(row.question)}${row.matchScore < 100 && row.matchScore > 0 ? `<br><span style="font-size:0.7rem;color:var(--text-muted)">fuzzy: ${row.matchScore}%</span>` : ''}</td>
      <td class="r-col">${escHtml(row.aiResponse)}</td>
      <td class="r-col">${escHtml(row.manualResponse)}</td>
      <td>${statusHtml}</td>
    `;
    tbody.appendChild(tr);
  });

  const empty = document.getElementById('empty-state');
  empty.classList.toggle('hidden', shown > 0);
}

function setFilter(f, el) {
  activeFilter = f;
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  applyFilter();
}

function escHtml(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ─── Export ───────────────────────────────────────────────
function buildDomainSheet(domainRows, highlight) {
  const cols = ['A', 'B', 'C', 'D', 'E', 'F'];
  const header = [['#', 'Question ID', 'Question', 'AI Response', 'Manual Response', 'Comparison']];
  const rows = domainRows.map((r, i) => [
    i + 1, r.qid, r.question, r.aiResponse, r.manualResponse,
    r.status === 'match' ? 1 : 0
  ]);

  const ws = XLSX.utils.aoa_to_sheet(header.concat(rows));
  ws['!cols'] = [{ wch: 5 }, { wch: 15 }, { wch: 55 }, { wch: 22 }, { wch: 22 }, { wch: 12 }];

  // Overwrite Comparison column (F) with IF formula; keep cached value so sheet opens correctly
  domainRows.forEach((r, i) => {
    const ri = i + 2;
    ws[`F${ri}`] = { t: 'n', f: `IF(D${ri}=E${ri},1,0)`, v: r.status === 'match' ? 1 : 0 };
  });

  // Header row style
  cols.forEach(c => {
    const addr = c + '1';
    if (!ws[addr]) return;
    ws[addr].s = {
      fill: { fgColor: { rgb: '22263A' } },
      font: { bold: true, color: { rgb: 'E8EAF6' } },
      alignment: { horizontal: 'center' }
    };
  });

  // Row highlight
  if (highlight) {
    domainRows.forEach((r, i) => {
      const rowIdx = i + 2;
      cols.forEach(col => {
        const addr = col + rowIdx;
        if (!ws[addr]) ws[addr] = { v: '', t: 's' };
        let fill = null;
        if (r.status === 'mismatch')   fill = { fgColor: { rgb: 'FEE2E2' } };
        else if (r.status === 'unmatched') fill = { fgColor: { rgb: 'FEF3C7' } };
        else if (r.status === 'match') fill = { fgColor: { rgb: 'DCFCE7' } };
        if (fill) ws[addr].s = { fill };
      });
    });
  }

  // Total row with COUNTIF formula
  const lastDataRow = domainRows.length + 1;
  const totalMatched = domainRows.filter(r => r.status === 'match').length;
  const totalRowIdx = domainRows.length + 2;
  XLSX.utils.sheet_add_aoa(ws, [['', '', 'Total', '', '', totalMatched]], { origin: -1 });
  ws[`F${totalRowIdx}`] = { t: 'n', f: `COUNTIF(F2:F${lastDataRow},1)`, v: totalMatched };
  cols.forEach(col => {
    const addr = col + totalRowIdx;
    if (!ws[addr]) ws[addr] = { v: '', t: 's' };
    ws[addr].s = {
      fill: { fgColor: { rgb: '22263A' } },
      font: { bold: true, color: { rgb: 'E8EAF6' } },
      alignment: { horizontal: col === 'F' ? 'center' : 'left' }
    };
  });

  return ws;
}

function safeSheetName(name) {
  // Excel sheet names: max 31 chars, no special chars
  return String(name).replace(/[:\\\/\?\*\[\]]/g, '_').slice(0, 31);
}

function exportExcel() {
  if (!results.length) { toast('Nothing to export yet.', 'error'); return; }

  const highlight  = document.getElementById('opt-highlight').checked;
  const addSummary = document.getElementById('opt-summary').checked;

  const wb = XLSX.utils.book_new();

  // Group results by domain
  const domainMap = new Map();
  results.forEach(r => {
    if (!domainMap.has(r.domain)) domainMap.set(r.domain, []);
    domainMap.get(r.domain).push(r);
  });

  // One sheet per domain — also collect info for summary cross-sheet formulas
  const usedNames = new Set();
  const domainInfo = [];
  const isAnswered = r => String(r.aiResponse ?? '').trim() !== '' && r.aiResponse !== '—';
  const normV = v => String(v ?? '').trim().toLowerCase();

  for (const [domain, domainRows] of domainMap) {
    let sheetName = safeSheetName(domain);
    if (usedNames.has(sheetName)) {
      let n = 2;
      while (usedNames.has(sheetName + '_' + n)) n++;
      sheetName = sheetName + '_' + n;
    }
    usedNames.add(sheetName);
    XLSX.utils.book_append_sheet(wb, buildDomainSheet(domainRows, highlight), sheetName);

    // Pre-compute cached values for summary (used as fallback v in formula cells)
    const total       = domainRows.length;
    const answered    = domainRows.filter(isAnswered).length;
    const notAnswered = total - answered;
    const correct     = domainRows.filter(r => r.status === 'match').length;
    const accuracy    = answered > 0 ? parseFloat((correct / answered * 100).toFixed(2)) : 0;
    const yesNo       = domainRows.filter(r => normV(r.manualResponse) === 'yes' && normV(r.aiResponse) === 'no').length;
    const yesPartial  = domainRows.filter(r => normV(r.manualResponse) === 'yes' && normV(r.aiResponse) === 'partial').length;
    domainInfo.push({ domain, sheetName, rowCount: total, cached: { total, answered, notAnswered, correct, accuracy, yesNo, yesPartial } });
  }

  // Summary sheet with formulas referencing each domain sheet
  if (addSummary) {
    const sumHeaders = [
      'Domain', 'Total Questions', 'Total Answered by AI', 'Total Not Answered by AI',
      'Correct Responses of AI (against Manual Responses)', 'Accuracy (%)',
      'Manual Yes → AI No', 'Manual Yes → AI Partial'
    ];

    // Build with cached values first so sheet opens correctly before Excel recalculates
    const sumData = [sumHeaders];
    domainInfo.forEach(d => {
      const c = d.cached;
      sumData.push([d.domain, c.total, c.answered, c.notAnswered, c.correct, c.accuracy, c.yesNo, c.yesPartial]);
    });
    const totalRowIdx = domainInfo.length + 3;
    sumData.push(['', '', '', '', '', '', '', '']);

    // Grand cached totals for TOTAL row
    let gTotal = 0, gAnswered = 0, gNotAnswered = 0, gCorrect = 0, gYesNo = 0, gYesPartial = 0;
    domainInfo.forEach(d => {
      gTotal += d.cached.total; gAnswered += d.cached.answered; gNotAnswered += d.cached.notAnswered;
      gCorrect += d.cached.correct; gYesNo += d.cached.yesNo; gYesPartial += d.cached.yesPartial;
    });
    const gAccuracy = gAnswered > 0 ? parseFloat((gCorrect / gAnswered * 100).toFixed(2)) : 0;
    sumData.push(['TOTAL', gTotal, gAnswered, gNotAnswered, gCorrect, gAccuracy, gYesNo, gYesPartial]);

    const ws2 = XLSX.utils.aoa_to_sheet(sumData);

    // Overwrite data rows with cross-sheet formulas
    domainInfo.forEach((d, i) => {
      const row  = i + 2;
      const sn   = `'${d.sheetName}'`;
      const last = d.rowCount + 1;
      const tRow = d.rowCount + 2;
      const c    = d.cached;

      ws2[`B${row}`] = { t: 'n', f: `COUNTA(${sn}!C2:C${last})`,                                          v: c.total };
      ws2[`C${row}`] = { t: 'n', f: `COUNTIF(${sn}!D2:D${last},"<>")`,                                    v: c.answered };
      ws2[`D${row}`] = { t: 'n', f: `B${row}-C${row}`,                                                     v: c.notAnswered };
      ws2[`E${row}`] = { t: 'n', f: `${sn}!F${tRow}`,                                                      v: c.correct };
      ws2[`F${row}`] = { t: 'n', f: `IF(C${row}=0,0,ROUND(E${row}/C${row}*100,2))`,                       v: c.accuracy };
      ws2[`G${row}`] = { t: 'n', f: `COUNTIFS(${sn}!E2:E${last},"Yes",${sn}!D2:D${last},"No")`,           v: c.yesNo };
      ws2[`H${row}`] = { t: 'n', f: `COUNTIFS(${sn}!E2:E${last},"Yes",${sn}!D2:D${last},"Partial")`,      v: c.yesPartial };
    });

    // TOTAL row formulas — SUM across all domain rows
    const lastDomainRow = domainInfo.length + 1;
    ws2[`B${totalRowIdx}`] = { t: 'n', f: `SUM(B2:B${lastDomainRow})`,                                     v: gTotal };
    ws2[`C${totalRowIdx}`] = { t: 'n', f: `SUM(C2:C${lastDomainRow})`,                                     v: gAnswered };
    ws2[`D${totalRowIdx}`] = { t: 'n', f: `SUM(D2:D${lastDomainRow})`,                                     v: gNotAnswered };
    ws2[`E${totalRowIdx}`] = { t: 'n', f: `SUM(E2:E${lastDomainRow})`,                                     v: gCorrect };
    ws2[`F${totalRowIdx}`] = { t: 'n', f: `IF(C${totalRowIdx}=0,0,ROUND(E${totalRowIdx}/C${totalRowIdx}*100,2))`, v: gAccuracy };
    ws2[`G${totalRowIdx}`] = { t: 'n', f: `SUM(G2:G${lastDomainRow})`,                                     v: gYesNo };
    ws2[`H${totalRowIdx}`] = { t: 'n', f: `SUM(H2:H${lastDomainRow})`,                                     v: gYesPartial };

    ws2['!cols'] = [
      { wch: 12 }, { wch: 16 }, { wch: 22 }, { wch: 26 }, { wch: 50 }, { wch: 14 }, { wch: 20 }, { wch: 22 }
    ];

    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(c => {
      const h = c + '1';
      if (ws2[h]) ws2[h].s = {
        fill: { fgColor: { rgb: '22263A' } },
        font: { bold: true, color: { rgb: 'E8EAF6' } },
        alignment: { wrapText: true, horizontal: 'center' }
      };
      const t = c + totalRowIdx;
      if (ws2[t]) ws2[t].s = {
        fill: { fgColor: { rgb: '22263A' } },
        font: { bold: true, color: { rgb: 'E8EAF6' } }
      };
    });

    XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
  }

  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `comparison_${date}.xlsx`);
  toast(`Exported ${domainMap.size} domain sheet(s)!`, 'success');
}

function pct(v, total) {
  return total ? (Math.round(v / total * 1000) / 10) + '%' : '0%';
}

// ─── Steps UI ─────────────────────────────────────────────
function updateSteps() {
  const s1 = !!sheets[1], s2 = !!sheets[2];
  const done = !!results.length;

  document.getElementById('step-1').className = 'step ' + (s1 && s2 ? 'done' : 'active');
  document.getElementById('step-2').className = 'step ' + (done ? 'done' : s1 && s2 ? 'active' : '');
  document.getElementById('step-3').className = 'step ' + (done ? 'active' : '');
}

// ─── Toast ────────────────────────────────────────────────
let toastTimer;
function toast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'show ' + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.className = '', 3000);
}
