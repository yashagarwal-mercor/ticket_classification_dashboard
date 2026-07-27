const MODEL = "claude-haiku-4-5";
const BATCH_SIZE = 30;
const CONCURRENCY = 5;
const CATCHALL_WORKFLOW = "Unclassified / Other";
const API_BASE = ""; // same-origin: server.py serves both the page and /api/*

let parsedRows = [];       // raw spreadsheet rows as objects, keyed by header
let headers = [];
let rubric = [];           // [{name, category, description}]
let cancelRequested = false;
let dashboardRows = [];    // [{company, workflow, category, hours, touches, first_touch}]

function showError(msg) {
  const el = document.getElementById('errorBanner');
  el.textContent = msg;
  el.style.display = msg ? 'block' : 'none';
}

function parseBoolish(v) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim().toLowerCase();
  if (['1', 'true', 'yes', 'y'].includes(s)) return 1;
  if (['0', 'false', 'no', 'n'].includes(s)) return 0;
  const n = parseFloat(s);
  return isNaN(n) ? null : n; // pass raw counts through unforced (aggregated rows may have counts > 1)
}

// --- Spreadsheet parsing ---

function parseCSV(text) {
  const rows = [];
  let row = [], field = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else { field += c; }
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else { field += c; }
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  if (!rows.length) return { headers: [], objects: [] };
  const hdrs = rows[0];
  const objects = rows.slice(1).map(r => {
    const o = {};
    hdrs.forEach((h, i) => { o[h] = r[i] !== undefined ? r[i] : ''; });
    return o;
  });
  return { headers: hdrs, objects };
}

function handleFile(file) {
  const reader = new FileReader();
  const isCSV = /\.csv$/i.test(file.name);
  reader.onload = (e) => {
    try {
      if (isCSV) {
        const { headers: h, objects } = parseCSV(e.target.result);
        headers = h; parsedRows = objects;
      } else {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const objects = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        headers = objects.length ? Object.keys(objects[0]) : [];
        parsedRows = objects;
      }
      onFileParsed();
    } catch (err) {
      showError('Could not parse the file: ' + err.message);
    }
  };
  if (isCSV) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

function onFileParsed() {
  document.getElementById('rowCountHint').textContent = `${parsedRows.length.toLocaleString()} rows detected, ${headers.length} columns.`;
  populateColumnSelectOptions();
  // Everything is visible up front: the column mapping (optionally auto-filled
  // by AI), the user-authored workflow rubric, and the classify button.
  document.getElementById('analyzePanel').style.display = 'block';
  document.getElementById('reviewPanel').style.display = 'block';
  document.getElementById('rubricPanel').style.display = 'block';
  document.getElementById('classifyPanel').style.display = 'block';
  if (!rubric.length) rubric = [{ name: '', category: 'General', description: '' }];
  renderRubricTable();
}

function populateColumnSelectOptions() {
  const optionalSelects = ['colCompany', 'colFirstRes', 'colTicketCount'];
  const requiredSelects = ['colHours', 'colTouches'];
  [...optionalSelects, ...requiredSelects].forEach(id => {
    const sel = document.getElementById(id);
    sel.innerHTML = (optionalSelects.includes(id) ? '<option value="">(none)</option>' : '') +
      headers.map(h => `<option value="${h}">${h}</option>`).join('');
  });
  const checklist = document.getElementById('textColsChecklist');
  checklist.innerHTML = headers.map(h => `
    <label><input type="checkbox" value="${escapeAttr(h)}" class="text-col-check"> ${h}</label>
  `).join('');
}

function escapeAttr(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

// --- AI analysis (column mapping + rubric) ---

async function runAnalyze() {
  if (!parsedRows.length) { showError('Upload a ticket spreadsheet first.'); return; }
  showError('');
  const statusEl = document.getElementById('analyzeStatus');
  const btn = document.getElementById('analyzeBtn');
  btn.disabled = true;
  statusEl.textContent = 'Analyzing...';
  statusEl.style.color = 'var(--text-dim)';

  const sampleRows = parsedRows.slice(0, 20);
  try {
    const resp = await fetch(`${API_BASE}/api/infer`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ headers, sampleRows }),
    });
    const body = await resp.json();
    if (!body.ok) throw new Error(body.error || 'Analysis failed');
    const parsed = JSON.parse(body.text);
    applyColumnMapping(parsed.column_mapping);
    statusEl.textContent = 'Column mapping filled in — review it below ✓';
    statusEl.style.color = 'var(--low)';
  } catch (err) {
    statusEl.textContent = '';
    showError('AI column mapping failed: ' + err.message + ' — you can still map the columns manually below.');
  } finally {
    btn.disabled = false;
  }
}

function applyColumnMapping(mapping) {
  if (!mapping) return;
  const setIfPresent = (id, value) => {
    const sel = document.getElementById(id);
    if (value && [...sel.options].some(o => o.value === value)) sel.value = value;
  };
  setIfPresent('colCompany', mapping.company_column);
  setIfPresent('colHours', mapping.hours_column);
  setIfPresent('colTouches', mapping.touches_column);
  setIfPresent('colFirstRes', mapping.first_resolution_column);
  setIfPresent('colTicketCount', mapping.ticket_count_column);

  const textCols = new Set(mapping.text_columns || []);
  document.querySelectorAll('.text-col-check').forEach(cb => {
    cb.checked = textCols.has(cb.value);
  });
}

document.getElementById('analyzeBtn').addEventListener('click', runAnalyze);
document.getElementById('fileInput').addEventListener('change', (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});

// --- Rubric table (AI-generated, editable, or replace via bulk paste) ---

function renderRubricTable() {
  const tbody = document.getElementById('rubricTableBody');
  tbody.innerHTML = rubric.map((r, i) => `
    <tr>
      <td class="col-name"><input data-i="${i}" data-f="name" value="${escapeAttr(r.name)}"></td>
      <td class="col-cat"><input data-i="${i}" data-f="category" value="${escapeAttr(r.category)}"></td>
      <td><input data-i="${i}" data-f="description" value="${escapeAttr(r.description)}"></td>
      <td class="col-actions"><button class="btn danger small" data-remove="${i}">&times;</button></td>
    </tr>
  `).join('');
  tbody.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', () => {
      rubric[+inp.dataset.i][inp.dataset.f] = inp.value;
    });
  });
  tbody.querySelectorAll('[data-remove]').forEach(btn => {
    btn.addEventListener('click', () => {
      rubric.splice(+btn.dataset.remove, 1);
      renderRubricTable();
    });
  });
}

document.getElementById('parseRubricBtn').addEventListener('click', () => {
  const text = document.getElementById('rubricPaste').value;
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const parsed = lines.map(line => {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length >= 3) return { name: parts[0], category: parts[1], description: parts.slice(2).join(' | ') };
    if (parts.length === 2) return { name: parts[0], category: 'General', description: parts[1] };
    return { name: parts[0], category: 'General', description: '' };
  }).filter(r => r.name);
  if (parsed.length) { rubric = parsed; renderRubricTable(); }
});

document.getElementById('addRubricRowBtn').addEventListener('click', () => {
  rubric.push({ name: '', category: 'General', description: '' });
  renderRubricTable();
});

// --- Thresholds ---

function getThresholds() {
  return {
    lowFrr: parseFloat(document.getElementById('thLowFrr').value) / 100,
    lowAht: parseFloat(document.getElementById('thLowAht').value),
    lowTouches: parseFloat(document.getElementById('thLowTouches').value),
    medFrr: parseFloat(document.getElementById('thMedFrr').value) / 100,
    medAht: parseFloat(document.getElementById('thMedAht').value),
  };
}

document.getElementById('resetThresholdsBtn').addEventListener('click', () => {
  document.getElementById('thLowFrr').value = 70;
  document.getElementById('thLowAht').value = 30;
  document.getElementById('thLowTouches').value = 1.7;
  document.getElementById('thMedFrr').value = 50;
  document.getElementById('thMedAht').value = 60;
  if (dashboardRows.length) render();
});

function tierFor(frr, aht, touchesPerTicket) {
  const t = getThresholds();
  if (frr >= t.lowFrr && aht <= t.lowAht && touchesPerTicket <= t.lowTouches) return 'low';
  if (frr >= t.medFrr && aht <= t.medAht) return 'medium';
  return 'high';
}

// --- Classification ---

function buildBatchPrompt(chunk) {
  const workflowList = rubric.map(r => `- ${r.name}: ${r.description}`).join('\n');
  const lines = chunk.map((row, i) => `${i}. ${row.text}`).join('\n');
  const system = `You are classifying helpdesk/support tickets into a fixed set of workflow categories.

Here are the ${rubric.length} valid workflows:
${workflowList}

For each ticket, pick the single best-matching workflow name from the list above,
copied EXACTLY as written. If a ticket genuinely does not fit any of these well,
respond with the literal string "NONE" instead of forcing a fit.

Rate your confidence in each answer as "high", "medium", or "low".`;
  const user = `Classify each of these tickets:\n\n${lines}`;
  return { system, user };
}

const RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    classifications: {
      type: "array",
      items: {
        type: "object",
        properties: {
          index: { type: "integer" },
          workflow: { type: "string" },
          confidence: { type: "string", enum: ["high", "medium", "low"] },
        },
        required: ["index", "workflow", "confidence"],
        additionalProperties: false,
      },
    },
  },
  required: ["classifications"],
  additionalProperties: false,
};

async function classifyBatch(chunk, validNames) {
  const { system, user } = buildBatchPrompt(chunk);
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, max_tokens: 4096, system, user, schema: RESPONSE_SCHEMA }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Classification failed');
  const parsed = JSON.parse(body.text);
  const result = new Map();
  for (const item of parsed.classifications || []) {
    if (item.index == null || item.index >= chunk.length) continue;
    let wf = item.workflow;
    if (wf !== "NONE" && !validNames.has(wf)) wf = "NONE";
    result.set(item.index, { workflow: wf, confidence: item.confidence });
  }
  return result;
}

async function runClassification() {
  cancelRequested = false;
  showError('');

  if (!rubric.length) { showError('Add at least one workflow to the rubric.'); return; }
  if (!parsedRows.length) { showError('Upload a ticket spreadsheet first.'); return; }

  const companyCol = document.getElementById('colCompany').value;
  const hoursCol = document.getElementById('colHours').value;
  const touchesCol = document.getElementById('colTouches').value;
  const firstResCol = document.getElementById('colFirstRes').value;
  const ticketCountCol = document.getElementById('colTicketCount').value;
  const textCols = [...document.querySelectorAll('.text-col-check:checked')].map(cb => cb.value);

  if (!hoursCol || !touchesCol) { showError('Map the required columns (Hours, Touches).'); return; }
  if (!textCols.length) { showError('Select at least one classification text column.'); return; }

  const tickets = parsedRows.map(r => ({
    company: companyCol ? String(r[companyCol] ?? 'All') : 'All',
    text: textCols.map(c => `${c}: ${r[c] ?? ''}`).join(' | '),
    hours: parseFloat(r[hoursCol]) || 0,
    touches: parseFloat(r[touchesCol]) || 0,
    firstTouchResolved: firstResCol ? parseBoolish(r[firstResCol]) : null,
    ticketCount: ticketCountCol ? (parseFloat(r[ticketCountCol]) || 0) : 1,
  }));

  const validNames = new Set(rubric.map(r => r.name));

  const chunks = [];
  for (let i = 0; i < tickets.length; i += BATCH_SIZE) chunks.push(tickets.slice(i, i + BATCH_SIZE));

  const classifications = new Array(tickets.length).fill(null);
  let completed = 0;

  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('classifyBtn').disabled = true;

  const updateProgress = () => {
    const pct = Math.round((completed / chunks.length) * 100);
    document.getElementById('progressBarInner').style.width = pct + '%';
    document.getElementById('progressText').textContent =
      `Classifying batch ${completed} of ${chunks.length} (${tickets.length.toLocaleString()} rows total)...`;
  };
  updateProgress();

  let chunkIdx = 0;
  let hadErrors = 0;
  let firstErrorMessage = null;

  function describeError(err) {
    if (err && err.message) return err.message;
    return String(err);
  }

  async function worker() {
    while (chunkIdx < chunks.length) {
      if (cancelRequested) return;
      const myIdx = chunkIdx++;
      const chunk = chunks[myIdx];
      const offset = myIdx * BATCH_SIZE;
      try {
        const result = await classifyBatch(chunk, validNames);
        for (const [i, val] of result.entries()) classifications[offset + i] = val;
      } catch (err) {
        hadErrors++;
        if (!firstErrorMessage) firstErrorMessage = describeError(err);
        console.error('Batch failed', myIdx, err);
      }
      completed++;
      updateProgress();
    }
  }

  await Promise.all(Array.from({ length: CONCURRENCY }, worker));

  document.getElementById('progressWrap').style.display = 'none';
  document.getElementById('classifyBtn').disabled = false;

  if (cancelRequested) {
    showError('Classification cancelled.');
    return;
  }

  if (hadErrors === chunks.length) {
    showError(`Every batch failed — nothing was classified. First error: ${firstErrorMessage}. Check the server console (running server.py) and try again.`);
    return;
  }

  const catByName = new Map(rubric.map(r => [r.name, r.category || 'General']));
  dashboardRows = tickets.map((t, i) => {
    const c = classifications[i];
    const workflow = (c && c.workflow && c.workflow !== 'NONE') ? c.workflow : CATCHALL_WORKFLOW;
    const category = catByName.get(workflow) || 'Other';
    let first_touch;
    if (t.firstTouchResolved != null) {
      first_touch = t.firstTouchResolved;
    } else {
      first_touch = (t.ticketCount <= 1 && t.touches === 1) ? 1 : 0;
    }
    return { company: t.company, workflow, category, hours: t.hours, touches: t.touches, first_touch, ticketCount: t.ticketCount };
  });

  if (hadErrors) showError(`${hadErrors} of ${chunks.length} batches failed and were left Unclassified. First error: ${firstErrorMessage}`);

  showDashboard();
}

document.getElementById('classifyBtn').addEventListener('click', runClassification);
document.getElementById('cancelBtn').addEventListener('click', () => { cancelRequested = true; });

document.getElementById('startOverBtn').addEventListener('click', () => {
  document.getElementById('dashboardSection').style.display = 'none';
  document.getElementById('setupSection').style.display = 'block';
});

// --- Dashboard (aggregation, filtering, rendering) ---

const companyFilter = document.getElementById('companyFilter');
const tierFilterEl = document.getElementById('tierFilter');
const tableBody = document.getElementById('tableBody');
const summaryStats = document.getElementById('summaryStats');
let sortKey = 'tickets';
let sortDir = -1;

function aggregateRows(rows) {
  const byWf = new Map();
  for (const r of rows) {
    if (!byWf.has(r.workflow)) byWf.set(r.workflow, { workflow: r.workflow, category: r.category, tickets: 0, hours: 0, touches: 0, first_touch: 0 });
    const b = byWf.get(r.workflow);
    b.tickets += r.ticketCount; b.hours += r.hours; b.touches += r.touches; b.first_touch += r.first_touch;
  }
  const out = [];
  for (const b of byWf.values()) {
    const aht = b.tickets > 0 ? (b.hours / b.tickets) * 60 : 0;
    const frr = b.tickets > 0 ? b.first_touch / b.tickets : 0;
    const touchesPerTicket = b.tickets > 0 ? b.touches / b.tickets : 0;
    out.push({ workflow: b.workflow, category: b.category, tickets: b.tickets, hours: b.hours, aht, frr, touches: touchesPerTicket, tier: tierFor(frr, aht, touchesPerTicket) });
  }
  return out;
}

// Catch-all rows always sort to the bottom, regardless of ticket volume: the
// app's own unmatched-ticket bucket (CATCHALL_WORKFLOW) plus any rubric entry
// the user categorized as "Other" (e.g. "Long-tail / Other", "Unclassified /
// Untagged Ticket") — these are deliberately low-signal buckets, not workflows
// worth surfacing near the top just because they happen to have volume.
function isCatchAll(r) {
  return r.workflow === CATCHALL_WORKFLOW || (r.category || '').trim().toLowerCase() === 'other';
}

function sortAggRows(agg) {
  const categoryTotals = new Map();
  for (const r of agg) categoryTotals.set(r.category, (categoryTotals.get(r.category) || 0) + r.tickets);
  agg.sort((a, b) => {
    const aCatch = isCatchAll(a), bCatch = isCatchAll(b);
    if (aCatch !== bCatch) return aCatch ? 1 : -1;
    if (a.category !== b.category) return categoryTotals.get(b.category) - categoryTotals.get(a.category);
    const va = a[sortKey], vb = b[sortKey];
    if (typeof va === 'string') return va.localeCompare(vb) * sortDir;
    return (va - vb) * sortDir;
  });
  return agg;
}

function populateCompanyFilter() {
  const companies = Array.from(new Set(dashboardRows.map(r => r.company))).sort();
  companyFilter.innerHTML = '<option value="all">All Companies</option>' + companies.map(c => `<option value="${c}">${c}</option>`).join('');
  companyFilter.value = 'all';
}

function getFilteredRows() {
  const company = companyFilter.value;
  return company === 'all' ? dashboardRows : dashboardRows.filter(r => r.company === company);
}

function render() {
  const filteredRaw = getFilteredRows();
  let agg = aggregateRows(filteredRaw);

  const tier = tierFilterEl.value;
  if (tier !== 'all') agg = agg.filter(r => r.tier === tier);

  const totalTickets = agg.reduce((s, r) => s + r.tickets, 0);
  const totalHours = agg.reduce((s, r) => s + r.hours, 0);
  const totalCompanies = new Set(filteredRaw.map(r => r.company)).size;

  for (const r of agg) {
    r.ticketPct = totalTickets > 0 ? (r.tickets / totalTickets * 100) : 0;
    r.hoursPct = totalHours > 0 ? (r.hours / totalHours * 100) : 0;
  }
  agg = sortAggRows(agg);

  const tierCounts = { low: 0, medium: 0, high: 0 };
  for (const r of agg) tierCounts[r.tier] += r.tickets;
  const tierPct = t => totalTickets > 0 ? (tierCounts[t] / totalTickets * 100) : 0;

  summaryStats.innerHTML = `
    <div class="stat"><div class="num">${totalTickets.toLocaleString()}</div><div class="lbl">Total Tickets</div></div>
    <div class="stat"><div class="num">${totalHours.toLocaleString(undefined, {maximumFractionDigits: 0})}</div><div class="lbl">Total Hours</div></div>
    <div class="stat"><div class="num">${agg.length}</div><div class="lbl">Workflows Shown</div></div>
    <div class="stat"><div class="num">${totalCompanies}</div><div class="lbl">Companies</div></div>
    <div class="stat tier-stat tier-stat-low"><div class="num">${tierCounts.low.toLocaleString()}</div><div class="lbl">Low-Complexity Tickets (${tierPct('low').toFixed(0)}%)</div></div>
    <div class="stat tier-stat tier-stat-medium"><div class="num">${tierCounts.medium.toLocaleString()}</div><div class="lbl">Medium-Complexity Tickets (${tierPct('medium').toFixed(0)}%)</div></div>
    <div class="stat tier-stat tier-stat-high"><div class="num">${tierCounts.high.toLocaleString()}</div><div class="lbl">High-Complexity Tickets (${tierPct('high').toFixed(0)}%)</div></div>
  `;

  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  tableBody.innerHTML = agg.map(r => `
    <tr>
      <td class="cat-name">${r.category}</td>
      <td class="wf-name">${r.workflow}${descByName.get(r.workflow) ? `<div class="wf-desc">${descByName.get(r.workflow)}</div>` : ''}</td>
      <td class="num-cell">${r.tickets.toLocaleString()}</td>
      <td class="num-cell">${r.ticketPct.toFixed(1)}%</td>
      <td class="num-cell">${r.hours.toLocaleString(undefined, {maximumFractionDigits: 1})}</td>
      <td class="num-cell">${r.hoursPct.toFixed(1)}%</td>
      <td class="num-cell">${r.aht.toFixed(1)}</td>
      <td class="num-cell">${(r.frr * 100).toFixed(0)}%</td>
      <td class="num-cell">${r.touches.toFixed(2)}</td>
      <td><span class="tier tier-${r.tier}">${r.tier}</span></td>
    </tr>
  `).join('');

  window.__currentAgg = agg;
}

document.querySelectorAll('table.dash thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    if (sortKey === key) sortDir *= -1; else { sortKey = key; sortDir = -1; }
    render();
  });
});
companyFilter.addEventListener('change', render);
tierFilterEl.addEventListener('change', render);

function showDashboard() {
  document.getElementById('setupSection').style.display = 'none';
  document.getElementById('dashboardSection').style.display = 'block';
  populateCompanyFilter();
  render();
}

// --- Export: clipboard TSV ---

function buildTSV() {
  const headersRow = ['Category', 'Workflow', 'Tickets', 'Total Hours', 'AHT (min)', 'FRR', 'Touches/Ticket', 'Complexity'];
  const lines = [headersRow.join('\t')];
  for (const r of window.__currentAgg || []) {
    lines.push([r.category, r.workflow, r.tickets, r.hours.toFixed(1), r.aht.toFixed(1), (r.frr * 100).toFixed(0) + '%', r.touches.toFixed(2), r.tier].join('\t'));
  }
  return lines.join('\n');
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

document.getElementById('exportBtn').addEventListener('click', async () => {
  const tsv = buildTSV();
  try { await navigator.clipboard.writeText(tsv); showToast('Copied to clipboard — paste into Excel'); return; } catch (e) {}
  try {
    const ta = document.createElement('textarea');
    ta.value = tsv; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.focus(); ta.select();
    const ok = document.execCommand('copy'); document.body.removeChild(ta);
    if (ok) { showToast('Copied to clipboard — paste into Excel'); return; }
  } catch (e) {}
  const modalTa = document.getElementById('modalTextarea');
  modalTa.value = tsv;
  document.getElementById('modalBackdrop').classList.add('show');
  modalTa.focus(); modalTa.select();
});

// --- Export: multi-sheet Excel (SpreadsheetML), one tab per company ---

function xmlEscape(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function sanitizeSheetName(name) {
  let s = String(name).replace(/[\\\/\?\*\[\]:]/g, ' ').trim();
  if (!s) s = 'Sheet';
  return s.slice(0, 31);
}

function rowsForCompany(company) {
  const filtered = company === null ? dashboardRows : dashboardRows.filter(r => r.company === company);
  let agg = aggregateRows(filtered);
  const tier = tierFilterEl.value;
  if (tier !== 'all') agg = agg.filter(r => r.tier === tier);
  return sortAggRows(agg);
}

function sheetXML(sheetName, rows) {
  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  const hdrs = ['Category', 'Workflow', 'Description', 'Tickets', 'Total Hours', 'AHT (min)', 'FRR', 'Touches/Ticket', 'Complexity'];
  const headerRow = '<Row>' + hdrs.map(h => `<Cell ss:StyleID="Header"><Data ss:Type="String">${xmlEscape(h)}</Data></Cell>`).join('') + '</Row>';
  const dataRows = rows.map(r => {
    const cells = [
      `<Cell><Data ss:Type="String">${xmlEscape(r.category)}</Data></Cell>`,
      `<Cell><Data ss:Type="String">${xmlEscape(r.workflow)}</Data></Cell>`,
      `<Cell><Data ss:Type="String">${xmlEscape(descByName.get(r.workflow) || '')}</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.tickets}</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.hours.toFixed(1)}</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.aht.toFixed(1)}</Data></Cell>`,
      `<Cell><Data ss:Type="String">${(r.frr * 100).toFixed(0)}%</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.touches.toFixed(2)}</Data></Cell>`,
      `<Cell><Data ss:Type="String">${xmlEscape(r.tier)}</Data></Cell>`,
    ].join('');
    return `<Row>${cells}</Row>`;
  }).join('');
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${headerRow}${dataRows}</Table></Worksheet>`;
}

function buildWorkbookXML() {
  const companies = Array.from(new Set(dashboardRows.map(r => r.company))).sort();
  const usedNames = new Set();
  const uniqueName = (base) => {
    let name = sanitizeSheetName(base), n = 2;
    while (usedNames.has(name)) name = sanitizeSheetName(base).slice(0, 28) + ' ' + n++;
    usedNames.add(name);
    return name;
  };
  let sheets = sheetXML(uniqueName('All Companies'), rowsForCompany(null));
  for (const company of companies) sheets += sheetXML(uniqueName(company), rowsForCompany(company));
  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles><Style ss:ID="Header"><Font ss:Bold="1"/></Style></Styles>
 ${sheets}
</Workbook>`;
}

document.getElementById('exportCsvBtn').addEventListener('click', async () => {
  const xml = buildWorkbookXML();
  const fileLabel = `workflows_by_company.xls`;
  let downloadAttempted = false;
  try {
    const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileLabel;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    downloadAttempted = true;
  } catch (e) {}

  if (downloadAttempted) { showToast(`Workbook download started (${fileLabel}) — a tab per company. If Excel warns about the file format, choose "Yes, open it".`); return; }

  const flat = buildTSV();
  try { await navigator.clipboard.writeText(flat); showToast('Download blocked here — copied a flattened summary to clipboard instead'); return; } catch (e) {}
  const modalTa = document.getElementById('modalTextarea');
  modalTa.value = flat;
  document.getElementById('modalBackdrop').classList.add('show');
  modalTa.focus(); modalTa.select();
});
