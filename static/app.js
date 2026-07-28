const MODEL = "claude-haiku-4-5";
const BATCH_SIZE = 12;      // notes bundles are long (full Dutch email threads) — smaller batches than the old 30
const CONCURRENCY = 5;
const CATCHALL_WORKFLOW = "Unclassified / Other";
const API_BASE = ""; // same-origin: server.py serves both the page and /api/*

// Time-entry column mapping: select id -> mapping key. The app accepts ONLY the
// one-row-per-time-entry format; rows are grouped into ticket records before anything else.
const COL_SELECTS = {
  colInstance: 'instance', colTicketNumber: 'ticketnumber', colStartDate: 'startdate',
  colSummaryNotes: 'summarynotes', colInternalNotes: 'internalnotes', colKpiLabel: 'kpiLabel',
  colHoursWorked: 'hoursWorked', colHoursToBill: 'hoursToBill', colNonBillable: 'nonbillable',
  colRoles: 'roles', colBillingCodes: 'billingcodes', colContractType: 'contracttype',
};
const REQUIRED_COLS = ['instance', 'ticketnumber', 'startdate', 'summarynotes', 'hoursWorked'];

// Mapping-file schema: select id <-> file key (e.g. colInstance <-> instance_column).
// Used to load a saved mapping (upload) and to export the current one (download).
// applyColumnMapping() consumes the *_column keys; getColumnMapping() reads the short keys.
const MAPPING_FILE_KEYS = {
  colInstance: 'instance_column', colTicketNumber: 'ticketnumber_column', colStartDate: 'startdate_column',
  colSummaryNotes: 'summarynotes_column', colInternalNotes: 'internalnotes_column', colKpiLabel: 'kpi_label_column',
  colHoursWorked: 'hours_worked_column', colHoursToBill: 'hours_to_bill_column', colNonBillable: 'nonbillable_column',
  colRoles: 'roles_column', colBillingCodes: 'billingcodes_column', colContractType: 'contracttype_column',
};

let parsedRows = [];       // raw time-entry rows as objects, keyed by header
let headers = [];
let ticketRecords = [];    // canonical ticket records (one per instance+ticketnumber) — see groupTimeEntriesIntoTickets
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
  ticketRecords = [];
  document.getElementById('groupSummary').style.display = 'none';
  document.getElementById('groupStatus').textContent = '';
  document.getElementById('rowCountHint').textContent = `${parsedRows.length.toLocaleString()} time-entry rows detected, ${headers.length} columns.`;
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
  const optionsHtml = '<option value="">(none)</option>' +
    headers.map(h => `<option value="${escapeAttr(h)}">${xmlEscape(h)}</option>`).join('');
  Object.keys(COL_SELECTS).forEach(id => {
    document.getElementById(id).innerHTML = optionsHtml;
  });
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
  setIfPresent('colInstance', mapping.instance_column);
  setIfPresent('colTicketNumber', mapping.ticketnumber_column);
  setIfPresent('colStartDate', mapping.startdate_column);
  setIfPresent('colSummaryNotes', mapping.summarynotes_column);
  setIfPresent('colInternalNotes', mapping.internalnotes_column);
  setIfPresent('colKpiLabel', mapping.kpi_label_column);
  setIfPresent('colHoursWorked', mapping.hours_worked_column);
  setIfPresent('colHoursToBill', mapping.hours_to_bill_column);
  setIfPresent('colNonBillable', mapping.nonbillable_column);
  setIfPresent('colRoles', mapping.roles_column);
  setIfPresent('colBillingCodes', mapping.billingcodes_column);
  setIfPresent('colContractType', mapping.contracttype_column);
}

// Read current dropdown selections into a mapping keyed by the short names groupTimeEntriesIntoTickets expects.
function getColumnMapping() {
  const m = {};
  for (const [id, key] of Object.entries(COL_SELECTS)) m[key] = document.getElementById(id).value;
  return m;
}

document.getElementById('analyzeBtn').addEventListener('click', runAnalyze);
document.getElementById('fileInput').addEventListener('change', (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});

// Upload a saved mapping file (.json). Accepts either a bare {..._column: value}
// object or the infer-style {column_mapping: {...}} wrapper.
document.getElementById('mappingFileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('analyzeStatus');
  const reader = new FileReader();
  reader.onload = (ev) => {
    let mapping;
    try {
      const parsed = JSON.parse(ev.target.result);
      mapping = parsed && parsed.column_mapping ? parsed.column_mapping : parsed;
    } catch (err) {
      showError('Could not read mapping file: ' + err.message + ' — expected a .json mapping.');
      return;
    }
    showError('');
    applyColumnMapping(mapping); // only sets a dropdown when the header exists in this file
    const wanted = Object.values(MAPPING_FILE_KEYS).map(k => mapping[k]).filter(Boolean);
    const missing = wanted.filter(h => !headers.includes(h));
    if (missing.length) {
      statusEl.textContent = `Mapping loaded — ${missing.length} mapped column(s) not found in this file; review step 4.`;
      statusEl.style.color = 'var(--high)';
    } else {
      statusEl.textContent = 'Mapping loaded — review it below ✓';
      statusEl.style.color = 'var(--low)';
    }
  };
  reader.readAsText(file);
});

// Download the current step-4 dropdown selections as a reusable mapping file.
document.getElementById('downloadMappingBtn').addEventListener('click', () => {
  const mapping = {};
  for (const [id, fileKey] of Object.entries(MAPPING_FILE_KEYS)) {
    mapping[fileKey] = document.getElementById(id).value || null;
  }
  const json = JSON.stringify({ column_mapping: mapping }, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'column_mapping.json';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  showToast('Mapping downloaded (column_mapping.json)');
});

// --- Ticket-record grouping (time entries -> one record per instance+ticketnumber) ---

// Parse the export's date format explicitly. The Autotask export uses US M/D/YYYY
// (optionally with time), which Date.parse handles inconsistently across engines/locales —
// so we parse it directly and only fall back to Date.parse for other shapes.
function parseEntryDate(s) {
  if (!s) return NaN;
  const str = String(s).trim();
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)).getTime();
  const t = Date.parse(str);
  return isNaN(t) ? NaN : t;
}

// Canonical ticket record (ENT-1998 §4 Stage 1). Groups many time-entry rows into one
// ticket keyed by instance + ticketnumber (ticketnumber alone collides across companies).
function groupTimeEntriesIntoTickets(rows, m) {
  const byKey = new Map();
  const val = (r, key) => (m[key] ? String(r[m[key]] ?? '').trim() : '');
  for (const r of rows) {
    const instance = val(r, 'instance');
    const tn = val(r, 'ticketnumber');
    if (!instance && !tn) continue; // skip blank rows
    const key = JSON.stringify([instance, tn]); // collision-safe compound key
    let rec = byKey.get(key);
    if (!rec) { rec = { ticket_id: tn, company: instance || 'Unknown', entries: [] }; byKey.set(key, rec); }
    rec.entries.push({
      start: val(r, 'startdate'),
      summary: val(r, 'summarynotes'),
      internal: val(r, 'internalnotes'),
      hoursWorked: m.hoursWorked ? (parseFloat(r[m.hoursWorked]) || 0) : 0,
      hoursToBill: m.hoursToBill ? (parseFloat(r[m.hoursToBill]) || 0) : 0,
      nonbillable: m.nonbillable ? (parseBoolish(r[m.nonbillable]) === 1) : false,
      role: val(r, 'roles'),
      billingCode: val(r, 'billingcodes'),
      kpi: val(r, 'kpiLabel'),
      contract: val(r, 'contracttype'),
    });
  }
  const distinct = arr => Array.from(new Set(arr.filter(Boolean)));
  const records = [];
  for (const rec of byKey.values()) {
    const entries = rec.entries.slice().sort((a, b) => {
      const da = parseEntryDate(a.start), db = parseEntryDate(b.start);
      if (isNaN(da) && isNaN(db)) return 0;
      if (isNaN(da)) return 1;
      if (isNaN(db)) return -1;
      return da - db;
    });
    const noteParts = [];
    for (const e of entries) { if (e.summary) noteParts.push(e.summary); if (e.internal) noteParts.push(e.internal); }
    const nbCount = entries.filter(e => e.nonbillable).length;
    let nonbillable_flag = 'none';
    if (nbCount > 0) nonbillable_flag = (nbCount === entries.length) ? 'fully' : 'partial';
    records.push({
      ticket_id: rec.ticket_id,
      company: rec.company,
      notes: noteParts.join('\n---\n'),
      hours: entries.reduce((s, e) => s + e.hoursWorked, 0),
      hours_to_bill: entries.reduce((s, e) => s + e.hoursToBill, 0),
      touches: entries.length,
      roles: distinct(entries.map(e => e.role)),
      billing_codes: distinct(entries.map(e => e.billingCode)),
      labels: distinct(entries.map(e => e.kpi)),
      contract_types: distinct(entries.map(e => e.contract)),
      nonbillable_hours: entries.reduce((s, e) => s + (e.nonbillable ? e.hoursWorked : 0), 0),
      nonbillable_flag,
      has_notes: noteParts.length > 0,
    });
  }
  return records;
}

// Validate required mapping, run grouping, store ticketRecords, and show a summary.
// Returns true on success. Both classification and non-billable analysis depend on this.
function runGrouping() {
  const m = getColumnMapping();
  const missing = REQUIRED_COLS.filter(k => !m[k]);
  if (missing.length) {
    const labels = { instance: 'Company/instance', ticketnumber: 'Ticket number', startdate: 'Start date/time', summarynotes: 'Summary notes', hoursWorked: 'Hours worked' };
    showError('Map the required columns first: ' + missing.map(k => labels[k]).join(', ') + '.');
    return false;
  }
  showError('');
  ticketRecords = groupTimeEntriesIntoTickets(parsedRows, m);
  if (!ticketRecords.length) { showError('Grouping produced 0 ticket records — check the Company/Ticket-number mapping.'); return false; }

  const totalHours = ticketRecords.reduce((s, r) => s + r.hours, 0);
  const nbHours = ticketRecords.reduce((s, r) => s + r.nonbillable_hours, 0);
  const withNotes = ticketRecords.filter(r => r.has_notes).length;
  const companies = new Set(ticketRecords.map(r => r.company)).size;
  const nbPct = totalHours > 0 ? (nbHours / totalHours * 100) : 0;
  const notesPct = ticketRecords.length > 0 ? (withNotes / ticketRecords.length * 100) : 0;

  const el = document.getElementById('groupSummary');
  el.style.display = 'block';
  el.innerHTML =
    `<strong>${ticketRecords.length.toLocaleString()}</strong> ticket records from ` +
    `<strong>${parsedRows.length.toLocaleString()}</strong> time entries across ` +
    `<strong>${companies}</strong> companies. ` +
    `Total ${totalHours.toLocaleString(undefined, {maximumFractionDigits: 0})} h; ` +
    `non-billable ${nbHours.toLocaleString(undefined, {maximumFractionDigits: 0})} h (${nbPct.toFixed(1)}%). ` +
    `${notesPct.toFixed(0)}% of tickets have ≥1 note.`;

  const statusEl = document.getElementById('groupStatus');
  statusEl.textContent = 'Grouped ✓';
  statusEl.style.color = 'var(--low)';
  return true;
}

document.getElementById('groupBtn').addEventListener('click', runGrouping);

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

// Parse "Name | Category | Description" lines (Category optional) into rubric rows.
function parseRubricText(text) {
  return text.split('\n').map(l => l.trim()).filter(Boolean).map(line => {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length >= 3) return { name: parts[0], category: parts[1], description: parts.slice(2).join(' | ') };
    if (parts.length === 2) return { name: parts[0], category: 'General', description: parts[1] };
    return { name: parts[0], category: 'General', description: '' };
  }).filter(r => r.name);
}

function applyParsedRubric(parsed) {
  if (!parsed.length) return false;
  rubric = parsed;
  renderRubricTable();
  return true;
}

document.getElementById('parseRubricBtn').addEventListener('click', () => {
  applyParsedRubric(parseRubricText(document.getElementById('rubricPaste').value));
});

// Upload a rubric file (.txt/.csv) — one "Name | Category | Description" per line.
document.getElementById('rubricFileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('rubricFileStatus');
  const reader = new FileReader();
  reader.onload = (ev) => {
    const parsed = parseRubricText(ev.target.result);
    if (applyParsedRubric(parsed)) {
      statusEl.textContent = `Loaded ${parsed.length} workflow${parsed.length === 1 ? '' : 's'} ✓`;
      statusEl.style.color = 'var(--low)';
    } else {
      statusEl.textContent = 'No workflows found — expected "Name | Category | Description" per line.';
      statusEl.style.color = 'var(--high)';
    }
  };
  reader.readAsText(file);
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
  if (!parsedRows.length) { showError('Upload a time-entry export first.'); return; }
  if (!ticketRecords.length && !runGrouping()) return; // group first (also validates required columns)

  // Classify the grouped ticket records. The classification signal is the concatenated
  // engineer notes (mostly Dutch), prefixed with the label/contract tags for extra context.
  const tickets = ticketRecords.map(rec => {
    const tags = [
      rec.labels.length ? `Labels: ${rec.labels.join(', ')}` : '',
      rec.contract_types.length ? `Contract: ${rec.contract_types.join(', ')}` : '',
    ].filter(Boolean).join(' | ');
    return {
      ticket_id: rec.ticket_id,
      company: rec.company,
      text: (tags ? tags + '\n' : '') + (rec.notes || '(no notes)'),
      hours: rec.hours,
      touches: rec.touches,
      firstTouchResolved: null,
      ticketCount: 1,
      record: rec,
    };
  });

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
    if (t.record) { t.record.workflow = workflow; t.record.category = category; } // let the non-billable pivot be workflow-aware
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
  setDashboardView('workflow');
}

// --- Dashboard view switching (Workflow vs Non-Billable) ---

function setDashboardView(view) {
  const isNb = view === 'nonbillable';
  document.getElementById('workflowView').style.display = isNb ? 'none' : 'block';
  document.getElementById('nonBillableView').style.display = isNb ? 'block' : 'none';
  // Workflow-specific export buttons only make sense on the workflow view (NB export lands in Phase 5).
  document.getElementById('exportCsvBtn').style.display = isNb ? 'none' : '';
  document.getElementById('exportBtn').style.display = isNb ? 'none' : '';
  document.getElementById('viewWorkflowBtn').className = 'btn small' + (isNb ? ' secondary' : '');
  document.getElementById('viewNonBillableBtn').className = 'btn small' + (isNb ? '' : ' secondary');
  if (isNb) renderNonBillable();
}

// Reachable straight after grouping (no classification / no API calls).
function showNonBillable() {
  if (!ticketRecords.length && !runGrouping()) return;
  document.getElementById('setupSection').style.display = 'none';
  document.getElementById('dashboardSection').style.display = 'block';
  populateNbCompanyFilter();
  setDashboardView('nonbillable');
}

document.getElementById('viewWorkflowBtn').addEventListener('click', () => {
  if (!dashboardRows.length) { showError('Classify tickets first to see the workflow dashboard.'); return; }
  showError('');
  setDashboardView('workflow');
});
document.getElementById('viewNonBillableBtn').addEventListener('click', () => setDashboardView('nonbillable'));
document.getElementById('nonBillableBtn').addEventListener('click', showNonBillable);

// --- Feature B: non-billable analysis (aggregation over ticket records, no LLM) ---

const NB_GROUP_DEFS = {
  company_workflow: { cols: ['Company', 'Workflow'], parts: r => [r.company, r.workflow || '(not classified)'], showWfPct: true },
  company: { cols: ['Company'], parts: r => [r.company], showWfPct: false },
  workflow: { cols: ['Workflow'], parts: r => [r.workflow || '(not classified)'], showWfPct: true },
  contract: { cols: ['Contract type'], parts: r => [r.contract_types.join(' + ') || '(none)'], showWfPct: false },
  role: { cols: ['Role / team'], parts: r => [r.roles.join(' + ') || '(none)'], showWfPct: false },
};

function nbScopeFilter(records, scope) {
  let recs = records.filter(r => r.nonbillable_hours > 0);
  if (scope === 'fully') recs = recs.filter(r => r.nonbillable_flag === 'fully');
  else if (scope === 'partial') recs = recs.filter(r => r.nonbillable_flag === 'partial');
  return recs;
}

// Build grouped non-billable rows. Denominator for the "% of group hours" metric is the
// group's TOTAL hours (billable + non-billable) over all records — so for a Workflow grouping
// this is exactly ENT-1998 §3's "% of that workflow's total hours that is non-billable".
function buildNbTable(allRecords, def, scope) {
  const inScope = nbScopeFilter(allRecords, scope);
  const totalNb = inScope.reduce((s, r) => s + r.nonbillable_hours, 0);
  const keyOf = r => def.parts(r).join('␟');
  const groupAllHours = new Map();
  for (const r of allRecords) { const k = keyOf(r); groupAllHours.set(k, (groupAllHours.get(k) || 0) + r.hours); }
  const groups = new Map();
  for (const r of inScope) {
    const k = keyOf(r);
    let g = groups.get(k);
    if (!g) { g = { key: k, parts: def.parts(r), fully: 0, partial: 0, nbHours: 0 }; groups.set(k, g); }
    if (r.nonbillable_flag === 'fully') g.fully++; else if (r.nonbillable_flag === 'partial') g.partial++;
    g.nbHours += r.nonbillable_hours;
  }
  const rows = [...groups.values()].map(g => {
    const groupHours = groupAllHours.get(g.key) || 0;
    return {
      parts: g.parts,
      fully: g.fully,
      partial: g.partial,
      tickets: g.fully + g.partial,
      nbHours: g.nbHours,
      pctOfTotalNb: totalNb > 0 ? (g.nbHours / totalNb * 100) : 0,
      pctOfGroupHoursNb: groupHours > 0 ? (g.nbHours / groupHours * 100) : 0,
    };
  });
  rows.sort((a, b) => b.nbHours - a.nbHours);
  return { rows, totalNb };
}

function populateNbCompanyFilter() {
  const el = document.getElementById('nbCompanyFilter');
  const companies = Array.from(new Set(ticketRecords.map(r => r.company))).sort();
  el.innerHTML = '<option value="all">All Companies</option>' +
    companies.map(c => `<option value="${escapeAttr(c)}">${xmlEscape(c)}</option>`).join('');
  el.value = 'all';
}

function renderNonBillable() {
  const company = document.getElementById('nbCompanyFilter').value;
  const groupBy = document.getElementById('nbGroupBy').value;
  const scope = document.getElementById('nbScope').value;
  const def = NB_GROUP_DEFS[groupBy] || NB_GROUP_DEFS.company_workflow;

  const records = company === 'all' ? ticketRecords : ticketRecords.filter(r => r.company === company);
  const { rows, totalNb } = buildNbTable(records, def, scope);

  // Summary tiles
  const inScopeAll = nbScopeFilter(records, 'all');
  const fully = inScopeAll.filter(r => r.nonbillable_flag === 'fully').length;
  const partial = inScopeAll.filter(r => r.nonbillable_flag === 'partial').length;
  const totalHoursAll = records.reduce((s, r) => s + r.hours, 0);
  const nbHoursAll = records.reduce((s, r) => s + r.nonbillable_hours, 0);
  const nbPctAll = totalHoursAll > 0 ? (nbHoursAll / totalHoursAll * 100) : 0;
  const companiesAffected = new Set(inScopeAll.map(r => r.company)).size;
  document.getElementById('nbSummaryStats').innerHTML = `
    <div class="stat"><div class="num">${nbHoursAll.toLocaleString(undefined, {maximumFractionDigits: 0})}</div><div class="lbl">Non-Billable Hours</div></div>
    <div class="stat"><div class="num">${(fully + partial).toLocaleString()}</div><div class="lbl">Non-Billable Tickets</div></div>
    <div class="stat"><div class="num">${fully.toLocaleString()}</div><div class="lbl">Fully Non-Billable</div></div>
    <div class="stat"><div class="num">${partial.toLocaleString()}</div><div class="lbl">Partially Non-Billable</div></div>
    <div class="stat"><div class="num">${nbPctAll.toFixed(1)}%</div><div class="lbl">of All Hours</div></div>
    <div class="stat"><div class="num">${companiesAffected.toLocaleString()}</div><div class="lbl">Companies Affected</div></div>
  `;

  // Header
  const headCells = def.cols.map(c => `<th>${xmlEscape(c)}</th>`).join('') +
    '<th class="num-cell">NB Tickets</th>' +
    '<th class="num-cell">Fully</th>' +
    '<th class="num-cell">Partial</th>' +
    '<th class="num-cell">NB Hours</th>' +
    '<th class="num-cell">% of Total NB</th>' +
    (def.showWfPct ? '<th class="num-cell">% of Hours NB</th>' : '');
  document.getElementById('nbTableHead').innerHTML = `<tr>${headCells}</tr>`;

  // Body
  document.getElementById('nbTableBody').innerHTML = rows.map(r => {
    const dimCells = r.parts.map(p => `<td>${xmlEscape(p)}</td>`).join('');
    return `<tr>
      ${dimCells}
      <td class="num-cell">${r.tickets.toLocaleString()}</td>
      <td class="num-cell">${r.fully.toLocaleString()}</td>
      <td class="num-cell">${r.partial.toLocaleString()}</td>
      <td class="num-cell">${r.nbHours.toLocaleString(undefined, {maximumFractionDigits: 1})}</td>
      <td class="num-cell">${r.pctOfTotalNb.toFixed(1)}%</td>
      ${def.showWfPct ? `<td class="num-cell">${r.pctOfGroupHoursNb.toFixed(1)}%</td>` : ''}
    </tr>`;
  }).join('') || `<tr><td colspan="${def.cols.length + 6}" style="text-align:center; color:var(--text-dim);">No non-billable tickets in scope.</td></tr>`;
}

['nbCompanyFilter', 'nbGroupBy', 'nbScope'].forEach(id => {
  document.getElementById(id).addEventListener('change', renderNonBillable);
});

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
