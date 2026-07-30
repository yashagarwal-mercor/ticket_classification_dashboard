const MODEL = "claude-sonnet-5"; // bulk classification, reconciliation extraction, Playbook steps
const MODEL_SUMMARY = "claude-sonnet-5"; // one-shot quality tasks (theme summary, reconciliation)
const BATCH_SIZE = 12;      // notes bundles are long (full Dutch email threads) — smaller batches than the old 30
// The Batch API caps a single submission at 256MB. In the batch JSONL the rubric system
// prompt is repeated per request, so a full run's requests are split across several batches,
// each kept under this budget (headroom left for UTF-8/JSON-escaping inflation).
const MAX_BATCH_BYTES = 150 * 1024 * 1024;
const CONCURRENCY = 5;
const CATCHALL_WORKFLOW = "Unclassified / Other";
const COVERAGE_GAP_THRESHOLD = 0.10; // >10% unclassified triggers a rubric coverage-gap warning (§2.3)
const THEME_SAMPLE_SIZE = 40;        // # of unclassified notes sampled for theme summarization
const API_BASE = ""; // same-origin: server.py serves both the page and /api/*

// --- Playbook view (Company x Workflow: description, common steps, responsible tier) ---
const PLAYBOOK_MIN_EVIDENCE = 5;  // groups with fewer notes-bearing tickets skip the steps call -> "Insufficient evidence"
const PLAYBOOK_SAMPLE = 25;       // max note excerpts sent per (company, workflow) steps-synthesis call
const TIER_LEVELS = ['Tier 1 (Front-line)', 'Tier 2 (Escalation)', 'Tier 3 (Specialist Engineering)', 'Security / SOC', 'Management / Admin', 'Sales / Account', 'Other / Unclear'];

// Tickets-export (optional second file) column mapping. Carries ticket-level fields
// (title, description, issue/sub-issue type) that enrich the classification text, plus
// the support-queue label that is the source for Responsible Engineer Tier.
const META_COL_SELECTS = {
  colMetaInstance: 'instance', colMetaTicketNumber: 'ticketnumber',
  colMetaTitle: 'title', colMetaDescription: 'description',
  colMetaIssueType: 'issueType', colMetaSubIssueType: 'subIssueType',
  colMetaQueue: 'queue',
};
const META_AUTODETECT = {
  instance: ['all_tickets_tasks[instance]'],
  ticketnumber: ['all_tickets_tasks[ticketnumber]'],
  title: ['all_tickets_tasks[title]'],
  description: ['all_tickets_tasks[description]'],
  issueType: ['ticket_issue_type[label]'],
  subIssueType: ['ticket_sub_issue_type[label]'],
  queue: ['ticket_queue[label]'],
};

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
let datasetName = '';      // uploaded source filename, used to label exports
let ticketRecords = [];    // canonical ticket records (one per instance+ticketnumber) — see groupTimeEntriesIntoTickets
let rubric = [];           // [{name, category, description}]
let cancelRequested = false;
let dashboardRows = [];    // [{company, workflow, category, hours, touches, first_touch}]
let coverageInfo = null;   // {total, unclassified, pct, noNotes, themes:[{theme,description,examples}]} — §2.3 coverage gap
let activeBatchIds = [];   // batch ids currently being polled (a run is split across several); enables Cancel-batch
let reconciliation = null;      // Map(workflowName -> verdict) from Feature A reconciliation (Phase 5)
let citedEvidenceById = null;   // Map(uid -> {en, nl}) evidence for cited tickets (rendered + persisted)
let metaHeaders = [];           // headers of the optional Tickets export (for the mapping dropdowns)
let metaByKey = null;           // Map("company/ticketnumber" -> {headerName: value}) — compact,
                                // projected lookup for the Tickets export (only the columns we join on).
                                // We never keep the full rows: the description column alone is ~175M chars
                                // and materializing every row would OOM the tab on the real 79MB export.
let tierByQueueLabel = null;    // Map(raw queue label -> canonical tier bucket), from one-shot LLM normalization
let playbookRows = [];          // built Playbook rows: [{company, workflow, category, tickets, hours, aht, frr, touches, tier, steps, stepsStatus}]

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
  datasetName = file.name || '';
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
  coverageInfo = null;
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

  // Surface a pending Batch API run (from a previous session) so the user knows to re-group to resume.
  const pending = loadBatchState();
  const hint = document.getElementById('batchResumeHint');
  if (pending) {
    const ids = pending.batch_ids || (pending.batch_id ? [pending.batch_id] : []);
    hint.textContent = `A classification run (${ids.length} batch${ids.length === 1 ? '' : 'es'}) is in progress — Group the same file to resume watching it.`;
    hint.style.display = 'inline';
  } else {
    hint.style.display = 'none';
  }
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

// --- Tickets-export ingestion (optional second file) — carries ticket-level fields
// (title/description/issue type) that enrich classification, plus the support-queue label
// used for Responsible Engineer Tier. Joined onto ticketRecords by company + ticket_id. ---

// The Tickets export is huge (~83MB / ~315MB of XML) and OOMs the tab if parsed in-browser
// with SheetJS. Instead we POST the file to the LOCAL server (server.py), which streams it
// with openpyxl and returns a compact "<company>/<ticket>" -> {header: value} lookup keyed
// by header name. The file never leaves the machine (localhost) — DPA-safe.
async function handleMetaFile(file) {
  const statusEl = document.getElementById('metaStatus');
  statusEl.style.color = 'var(--text-dim)';
  statusEl.textContent = `Parsing “${file.name}” on the local server…`;
  metaByKey = null; metaHeaders = [];
  try {
    const cols = encodeURIComponent(JSON.stringify(META_AUTODETECT));
    const resp = await fetch(`${API_BASE}/api/tickets/parse?cols=${cols}`, { method: 'POST', body: file });
    const body = await resp.json();
    if (!body.ok) {
      showError('Could not parse the Tickets export: ' + (body.error || 'unknown error'));
      statusEl.textContent = '';
      return;
    }
    metaHeaders = body.headers || [];
    metaByKey = new Map(Object.entries(body.byKey || {}));
    populateMetaSelectOptions();
    const n = metaByKey.size;
    statusEl.textContent = n
      ? `Tickets export: ${n.toLocaleString()} tickets loaded, ${metaHeaders.length} columns. Joined automatically when you group (step 4).`
      : `Tickets export loaded, but no instance/ticket-number columns were detected — it won't join. Check the file's headers.`;
    statusEl.style.color = n ? 'var(--text-dim)' : 'var(--high)';
  } catch (err) {
    showError('Tickets export upload failed: ' + describeError(err));
    statusEl.textContent = '';
  }
}

function populateMetaSelectOptions() {
  const optionsHtml = '<option value="">(none)</option>' +
    metaHeaders.map(h => `<option value="${escapeAttr(h)}">${xmlEscape(h)}</option>`).join('');
  Object.keys(META_COL_SELECTS).forEach(id => {
    document.getElementById(id).innerHTML = optionsHtml;
  });
  // Auto-detect known TechOne header names; falls back to manual selection otherwise.
  for (const [id, key] of Object.entries(META_COL_SELECTS)) {
    const candidates = META_AUTODETECT[key] || [];
    const match = metaHeaders.find(h => candidates.includes(h));
    if (match) document.getElementById(id).value = match;
  }
}

document.getElementById('metaFileInput').addEventListener('change', (e) => {
  if (e.target.files[0]) handleMetaFile(e.target.files[0]);
});

// Join the Tickets export onto ticketRecords by company + ticket number, attaching
// ticket-level fields (title, description, issue/sub-issue type) used to enrich the
// classification text, plus the raw queue label (normalized into a support tier later,
// in runBuildPlaybook()). Runs automatically at the end of runGrouping() whenever a
// Tickets file has been uploaded — there is no separate "attach" step. Returns the
// number of ticket records that matched a row in the Tickets export.
function joinTicketMetadata() {
  if (!metaByKey || !metaByKey.size) return 0;
  // Each field maps to a source header chosen in the step-4 dropdowns (auto-detected for
  // standard TechOne exports). metaByKey stores values by header name, so we resolve by value.
  const sel = id => document.getElementById(id).value;
  const titleH = sel('colMetaTitle'), descH = sel('colMetaDescription'),
        issueH = sel('colMetaIssueType'), subH = sel('colMetaSubIssueType'),
        queueH = sel('colMetaQueue');
  const pull = (meta, h) => (h && meta[h] != null ? meta[h] : '');
  let matched = 0;
  for (const rec of ticketRecords) {
    const meta = metaByKey.get(`${rec.company}/${rec.ticket_id}`);
    if (!meta) continue;
    matched++;
    rec.title = pull(meta, titleH);
    rec.description = pull(meta, descH);
    rec.issue_type = pull(meta, issueH);
    rec.sub_issue_type = pull(meta, subH);
    rec.queue_label = pull(meta, queueH) || null;
  }
  tierByQueueLabel = null; // stale — Build Playbook re-normalizes against the newly attached labels
  return matched;
}

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

  // Join the optional Tickets export (uploaded in step 1) onto the fresh records.
  const metaMatched = joinTicketMetadata();

  // When a Tickets export is present, the Tickets view defines the analysis universe:
  // scope ticketRecords to only the tickets it contains, so classification (and every
  // downstream view) covers exactly those. Keeps dashboardRows<->ticketRecords 1:1.
  // If it matched nothing (wrong companies/period), don't zero the dataset — warn instead.
  let metaNote = '';
  if (metaByKey && metaByKey.size) {
    if (metaMatched > 0) {
      const before = ticketRecords.length;
      ticketRecords = ticketRecords.filter(r => metaByKey.has(`${r.company}/${r.ticket_id}`));
      const excluded = before - ticketRecords.length;
      metaNote = ` Scoped to the Tickets export: <strong>${ticketRecords.length.toLocaleString()}</strong> tickets` +
        (excluded ? `, excluding <strong>${excluded.toLocaleString()}</strong> time-entry-only tickets not present in it.` : '.');
    } else {
      metaNote = ` <span style="color:var(--high)">Tickets export matched 0 of these records — check it covers the same companies/period. Classifying all tickets (not scoped).</span>`;
    }
  }

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
    `${notesPct.toFixed(0)}% of tickets have ≥1 note.` + metaNote;

  const statusEl = document.getElementById('groupStatus');
  statusEl.textContent = 'Grouped ✓';
  statusEl.style.color = 'var(--low)';
  return true;
}

document.getElementById('groupBtn').addEventListener('click', () => {
  if (!runGrouping()) return;
  // An in-flight batch takes precedence (resume it); otherwise restore the latest
  // completed run for this dataset. A completed run is only overwritten when a new
  // run finishes, so cancelling a run leaves the prior results intact.
  if (!maybeResumeBatch()) maybeRestoreEnriched();
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

// Stable guidance appended to the system prompt. Serves two purposes: it improves
// classification quality, and it keeps the (byte-identical) system prefix above the
// model's minimum cacheable length (4096 tokens on Haiku) so prompt caching kicks in.
const CLASSIFY_GUIDELINES = `Classification guidelines:
- Base your decision on what the engineer actually did, as described in the notes, not only on the label or category tags. The tags are a hint; the notes are the evidence.
- Choose the workflow whose description best matches the primary activity of the ticket. When several could apply, pick the single dominant one — the activity that consumed the most effort or was the reason the ticket existed.
- Do not invent, abbreviate, translate, or paraphrase workflow names. Copy the chosen name exactly, character for character, from the list above — including punctuation, capitalization, slashes, and parentheses.
- Prefer a specific workflow over a generic catch-all when the notes clearly support it. Fall back to a general workflow only when the notes are genuinely generic.
- Only answer "NONE" when the ticket genuinely does not fit any workflow — not merely because the notes are short, informal, or written in Dutch. A brief note like "wachtwoord gereset" still clearly maps to a workflow.

Reading the notes:
- Notes are mostly Dutch, sometimes English, and may mix both. Reason over the Dutch directly; do not translate first.
- A ticket's notes are the concatenation of every time entry on that ticket, in chronological order, separated by "---". Read the whole bundle before deciding — the first entry states the request, later entries show the resolution.
- Notes may contain quoted email threads, greetings, and signatures ("Goedemiddag", "Met vriendelijke groet"); ignore the pleasantries and focus on the technical action.
- Common signals: "doorgezet naar" / "doorgestuurd" = escalated/forwarded; "rechten toegekend" = permissions granted; "wachtwoord/MFA/token reset" = credential reset; "TeamViewer"/"remote" = remote session; "printer" = printer work; "backup"/"restore" = backup work; "licentie" = licensing.
- "Interne uren" (internal hours) tickets are often internal/administrative work rather than a customer request — classify by the actual activity if one fits, else consider whether any workflow genuinely applies.

Choosing confidence:
- "high" when the notes clearly and unambiguously identify the activity and it maps cleanly to one workflow.
- "medium" when the mapping is plausible but the notes are partial, ambiguous between two workflows, or you inferred the activity indirectly.
- "low" when you are guessing from very thin signal (e.g. a one-word note or tags only).

Output rules:
- Return exactly one classification object per input ticket, using the same index shown before each ticket.
- The "workflow" value must be either an exact name from the list above or the literal string "NONE". Never output any other value.`;

function buildSystemPrompt() {
  const workflowList = rubric.map(r => `- ${r.name}: ${r.description}`).join('\n');
  return `You are classifying helpdesk/support tickets into a fixed set of workflow categories.
Each ticket is described by its engineer time-entry notes, which are mostly in Dutch (some English).
Classify based on the meaning of the notes — do NOT translate first, reason over the Dutch directly.

Here are the ${rubric.length} valid workflows:
${workflowList}

For each ticket, pick the single best-matching workflow name from the list above,
copied EXACTLY as written. If a ticket genuinely does not fit any of these well,
respond with the literal string "NONE" instead of forcing a fit.

Rate your confidence in each answer as "high", "medium", or "low".

${CLASSIFY_GUIDELINES}`;
}

function buildUserPrompt(chunk) {
  const lines = chunk.map((row, i) => `${i}. ${row.text}`).join('\n');
  return `Classify each of these tickets:\n\n${lines}`;
}

function buildBatchPrompt(chunk) {
  return { system: buildSystemPrompt(), user: buildUserPrompt(chunk) };
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

// --- Feature A, Stage 1: per-ticket extraction (§2.1.2 / §4 fixed schema) ---
// For each classified ticket, extract a fixed schema from its notes, per workflow
// (the workflow's rubric description is included so the model can spot rubric_gaps).
// Tickets are keyed by uid = "company/ticketnumber" (ticketnumber alone collides
// across companies). Notes are Dutch; all output is English except evidence_nl.

const EXTRACT_BATCH_SIZE = 8; // extraction output is larger per ticket than classification

const EXTRACTION_SCHEMA = {
  type: "object",
  properties: {
    extractions: {
      type: "array",
      items: {
        type: "object",
        properties: {
          ticket_id: { type: "string" },
          main_action: { type: "string" },
          bottleneck: { type: "string", enum: ["none", "customer wait", "vendor/third-party", "internal handoff", "approval", "technical blocker"] },
          bottleneck_detail: { type: "string" },
          customer_contact: { type: "boolean" },
          customer_contact_channel: { type: "string", enum: ["none", "email", "phone", "remote session", "on-site", "unknown"] },
          teams_involved: { type: "array", items: { type: "string" } },
          rubric_gap: { type: "string" },
          evidence_nl: { type: "string" },
          evidence_en: { type: "string" },
        },
        required: ["ticket_id", "main_action", "bottleneck", "bottleneck_detail", "customer_contact", "customer_contact_channel", "teams_involved", "rubric_gap", "evidence_nl", "evidence_en"],
        additionalProperties: false,
      },
    },
  },
  required: ["extractions"],
  additionalProperties: false,
};

// Generic (workflow-agnostic) extraction system prompt — shared across the whole
// extraction pass so it caches and a single-system Batch request works. The specific
// workflow name/description go in the user prompt (buildExtractionUser).
function buildExtractionSystem() {
  return `You are analyzing helpdesk/support tickets. Each request gives you a workflow's name and rubric description, then a batch of tickets classified into that workflow, each with its engineer time-entry notes (mostly Dutch, sometimes English). Reason over the Dutch directly — do NOT translate first. For each ticket, extract this fixed schema:
- main_action: a short English verb phrase for the dominant manual action performed (e.g. "reset password in Entra", "assigned AD group permissions", "escalated to Engineering").
- bottleneck: the main roadblock/wait — one of: none, customer wait, vendor/third-party, internal handoff, approval, technical blocker.
- bottleneck_detail: a short English phrase with specifics (empty string when bottleneck is "none").
- customer_contact: true if the engineer communicated with the customer / end-user on this ticket.
- customer_contact_channel: one of none, email, phone, remote session, on-site, unknown.
- teams_involved: the roles/teams that touched the ticket, including escalations named in the notes (e.g. "doorgezet naar Engineering" -> "Engineering"). Use short English labels.
- rubric_gap: anything the ticket actually involved that the given workflow description does NOT mention; empty string if the ticket fits the description.
- evidence_nl: the single most telling short quote from the notes, verbatim in its original language.
- evidence_en: the English translation of that quote.

Output ALL fields in English EXCEPT evidence_nl (keep it verbatim). Echo ticket_id EXACTLY as given. Return exactly one object per input ticket.`;
}

function buildExtractionUser(chunk, wf) {
  return `Workflow: "${wf.name}"\nWorkflow description: "${wf.description}"\n\nTickets:\n\n` + chunk.map(t =>
    `ticket_id: ${t.uid}\nnotes: ${(t.notes || '').replace(/\s+/g, ' ')}`
  ).join('\n\n');
}

// Parse an extraction response's text into Map(uid -> extraction), dropping any
// echoed id that wasn't sent (anti-hallucination). Shared by the sync + batch paths.
function parseExtractions(text, sentIds) {
  const result = new Map();
  let parsed;
  try { parsed = JSON.parse(text); } catch (e) { return result; }
  for (const e of parsed.extractions || []) {
    if (e && typeof e.ticket_id === 'string' && sentIds.has(e.ticket_id)) result.set(e.ticket_id, e);
  }
  return result;
}

// Extract one chunk of same-workflow tickets synchronously. chunk: [{uid, notes}]. wf: {name, description}.
async function extractBatch(chunk, wf) {
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, max_tokens: 4096, system: buildExtractionSystem(), user: buildExtractionUser(chunk, wf), schema: EXTRACTION_SCHEMA }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Extraction failed');
  return parseExtractions(body.text, new Set(chunk.map(t => t.uid)));
}

// --- Feature A, Stage 2: per-workflow reconcile (§2.1.3 / §2.2) ---
const RECON_MIN_EVIDENCE = 5;  // < this many tickets-with-notes -> "Insufficient evidence" (no LLM call, §2.2)
const RECON_SAMPLE = 40;       // max extractions sent to a single reconcile call (bounds tokens)

const RECONCILE_SCHEMA = {
  type: "object",
  properties: {
    reconciliation_status: { type: "string", enum: ["Aligned", "Minor drift", "Significantly changed"] },
    observed_differences: { type: "string" },
    roadblock: { type: "string" },
    main_action: { type: "string" },
    evidence_ticket_ids: { type: "array", items: { type: "string" } },
    suggested_description_update: { type: "string" },
  },
  required: ["reconciliation_status", "observed_differences", "roadblock", "main_action", "evidence_ticket_ids", "suggested_description_update"],
  additionalProperties: false,
};

// Compact a workflow's extractions into a summary that bounds the reconcile prompt's tokens.
function aggregateWorkflowExtractions(extractionList, total) {
  const bottleneckHist = {}, teamsHist = {};
  let customerContact = 0, gapCount = 0;
  for (const e of extractionList) {
    bottleneckHist[e.bottleneck] = (bottleneckHist[e.bottleneck] || 0) + 1;
    if (e.customer_contact) customerContact++;
    if (e.rubric_gap && e.rubric_gap.trim()) gapCount++;
    for (const t of (e.teams_involved || [])) teamsHist[t] = (teamsHist[t] || 0) + 1;
  }
  const topTeams = Object.entries(teamsHist).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([t, n]) => `${t} (${n})`);
  const sample = extractionList.slice(0, RECON_SAMPLE).map(e => ({
    uid: e.ticket_id, main_action: e.main_action, bottleneck: e.bottleneck,
    bottleneck_detail: e.bottleneck_detail, rubric_gap: e.rubric_gap, evidence_en: e.evidence_en,
  }));
  return { total, withEvidence: extractionList.length, bottleneckHist, customerContact, topTeams, gapCount, sample };
}

function buildReconcileSystem() {
  return `You are reconciling ONE helpdesk workflow's rubric description against what engineers ACTUALLY did, using structured extractions from the tickets classified into it. Judge how far the real process has drifted from the description.

Definitions:
- "Aligned": the notes match the description; the process is as documented.
- "Minor drift": mostly matches, with small additions, tooling changes, or scope nuances.
- "Significantly changed": the real process differs materially — different steps, tooling, or scope, or the description no longer describes what these tickets actually are.

Produce (all in English):
- reconciliation_status: Aligned | Minor drift | Significantly changed.
- observed_differences: 1-3 sentences on what the notes show that the description doesn't cover, or vice versa.
- roadblock: the most frequent/impactful roadblock across these tickets.
- main_action: the dominant manual action across these tickets.
- evidence_ticket_ids: 2-3 ticket ids from the input that best support the verdict (at least 3 when Significantly changed). Only use ids present in the input.
- suggested_description_update: a proposed replacement rubric description ONLY when status is "Significantly changed"; otherwise an empty string.`;
}

function buildReconcileUser(wf, agg) {
  const lines = agg.sample.map(s =>
    `- [${s.uid}] action: ${s.main_action}; bottleneck: ${s.bottleneck}${s.bottleneck_detail ? ' (' + s.bottleneck_detail + ')' : ''}; gap: ${s.rubric_gap || '—'}; evidence: ${s.evidence_en}`
  ).join('\n');
  return `Workflow: "${wf.name}"
Rubric description: "${wf.description}"

Tickets classified into this workflow: ${agg.total} (with notes/extractions: ${agg.withEvidence}).
Bottleneck distribution: ${JSON.stringify(agg.bottleneckHist)}
Customer contact on: ${agg.customerContact}/${agg.withEvidence} tickets.
Top teams/roles involved: ${agg.topTeams.join(', ') || '—'}
Tickets whose notes involved something the description does not mention: ${agg.gapCount}/${agg.withEvidence}

Per-ticket extractions (${agg.sample.length} shown):
${lines}`;
}

async function reconcileWorkflow(wf, agg) {
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL_SUMMARY, max_tokens: 2048, system: buildReconcileSystem(), user: buildReconcileUser(wf, agg), schema: RECONCILE_SCHEMA }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Reconcile failed');
  const v = JSON.parse(body.text);
  const sentUids = new Set(agg.sample.map(s => s.uid));
  const cited = (v.evidence_ticket_ids || []).filter(id => sentUids.has(id)); // drop unknown/hallucinated ids
  return {
    status: v.reconciliation_status,
    observed_differences: v.observed_differences || '',
    roadblock: v.roadblock || '',
    main_action: v.main_action || '',
    evidence_ticket_ids: cited,
    suggested_description_update: v.reconciliation_status === 'Significantly changed' ? (v.suggested_description_update || '') : '',
    total: agg.total, withEvidence: agg.withEvidence,
  };
}

// Extraction, in-browser worker pool. tasks: [{chunk:[{uid,notes}], wf:{name,description}}].
// Returns {extractionsByWf: Map(name->[extraction]), exById: Map(uid->extraction)}.
async function runExtractionSync(tasks, setStatus) {
  const extractionsByWf = new Map(), exById = new Map();
  let ti = 0, done = 0, errors = 0;
  const up = () => setStatus(`Extracting ticket details: ${done}/${tasks.length} batches${errors ? ` (${errors} failed)` : ''}...`);
  up();
  async function worker() {
    while (ti < tasks.length) {
      if (cancelRequested) return;
      const t = tasks[ti++];
      try {
        const m = await extractBatch(t.chunk, t.wf);
        if (!extractionsByWf.has(t.wf.name)) extractionsByWf.set(t.wf.name, []);
        for (const [uid, e] of m) { extractionsByWf.get(t.wf.name).push(e); exById.set(uid, e); }
      } catch (err) { errors++; console.error('extract batch failed', err); }
      done++; up();
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, worker));
  return { extractionsByWf, exById, errors };
}

// Extraction via the Batch API (scales past the relay ceiling; tab-closeable). One batch,
// shared cached system prompt, custom_id = "t<taskIndex>" maps results back to (workflow, chunk).
// In-session polling only (no reload-resume for the reconciliation batch yet).
async function runExtractionBatch(tasks, setStatus) {
  const extractionsByWf = new Map(), exById = new Map();
  const requests = tasks.map((t, i) => ({ custom_id: `t${i}`, user: buildExtractionUser(t.chunk, t.wf) }));
  setStatus(`Submitting ${requests.length} extraction requests to the Batch API...`);
  const createResp = await fetch(`${API_BASE}/api/batch/create`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, max_tokens: 4096, system: buildExtractionSystem(), schema: EXTRACTION_SCHEMA, requests }),
  });
  const created = await createResp.json();
  if (!created.ok) throw new Error(created.error || 'batch create failed');
  const batchId = created.batch_id;

  while (true) {
    if (cancelRequested) return { extractionsByWf, exById, cancelled: true };
    let status;
    try { const r = await fetch(`${API_BASE}/api/batch/status?id=${encodeURIComponent(batchId)}`); status = await r.json(); }
    catch (e) { status = { ok: false, error: describeError(e) }; }
    if (status.ok) {
      const c = status.request_counts || {};
      const doneN = (c.succeeded || 0) + (c.errored || 0) + (c.canceled || 0) + (c.expired || 0);
      const errPart = (c.errored || 0) ? `, ${c.errored} errored` : '';
      setStatus(`Extracting via Batch API: ${doneN}/${requests.length} (${c.succeeded || 0} ok${errPart}, ${status.processing_status}). Safe to close the tab.`);
      if (status.processing_status === 'ended') break;
      if (['canceled', 'canceling', 'expired'].includes(status.processing_status)) throw new Error(`extraction batch ${status.processing_status}`);
    } else {
      if (isPermanentBatchError(status.error)) throw new Error(`extraction batch tracking failed — invalid batch id (${status.error || ''})`);
      setStatus(`Batch status check failed (${status.error || 'unknown'}) — retrying...`);
    }
    await sleep(5000);
  }

  const resultsResp = await fetch(`${API_BASE}/api/batch/results?id=${encodeURIComponent(batchId)}`);
  const body = await resultsResp.json();
  if (!body.ok) throw new Error(body.error || 'batch results failed');
  let errors = 0;
  for (const res of body.results || []) {
    const m = /^t(\d+)$/.exec(res.custom_id || '');
    if (!m) continue;
    const t = tasks[+m[1]];
    if (!t) continue;
    if (!res.ok) { errors++; console.error('extract request failed', res.custom_id, res.error); continue; }
    const map = parseExtractions(res.text, new Set(t.chunk.map(x => x.uid)));
    if (!extractionsByWf.has(t.wf.name)) extractionsByWf.set(t.wf.name, []);
    for (const [uid, e] of map) { extractionsByWf.get(t.wf.name).push(e); exById.set(uid, e); }
  }
  return { extractionsByWf, exById, errors };
}

// Orchestrate Feature A: extract every classified rubric-workflow ticket (with notes),
// then reconcile each of the rubric's workflows (all get a status; <N with notes ->
// "Insufficient evidence", no LLM). Wired to a button + rendered in Stage 5c.
async function runReconciliation() {
  if (!dashboardRows.length) { showError('Classify tickets first, then run reconciliation.'); return; }
  cancelRequested = false;
  showError('');

  const validNames = new Set(rubric.map(r => r.name));
  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  const byWorkflow = new Map(); // name -> [{uid, notes}]
  const totals = new Map();     // name -> classified ticket count
  for (let i = 0; i < dashboardRows.length; i++) {
    const wfName = dashboardRows[i].workflow;
    if (!validNames.has(wfName)) continue; // skip catch-all / non-rubric buckets
    totals.set(wfName, (totals.get(wfName) || 0) + 1);
    const rec = ticketRecords[i];
    if (!rec || !rec.has_notes) continue;
    if (!byWorkflow.has(wfName)) byWorkflow.set(wfName, []);
    byWorkflow.get(wfName).push({ uid: `${rec.company}/${rec.ticket_id}`, notes: rec.notes });
  }

  const statusEl = document.getElementById('reconStatus');
  const btn = document.getElementById('runReconBtn');
  btn.disabled = true;
  const setStatus = (msg, done) => { statusEl.textContent = msg; statusEl.style.color = done ? 'var(--low)' : 'var(--text-dim)'; };

  // Stage 1 — extraction (per-workflow chunks). exById surfaces per-ticket evidence
  // (English + original Dutch) for the tickets the reconcile cites. Runs via the Batch
  // API when the #useBatch toggle is on (scales past the relay ceiling), else in-browser.
  const tasks = [];
  for (const [wfName, tix] of byWorkflow) {
    const wf = { name: wfName, description: descByName.get(wfName) || '' };
    for (let i = 0; i < tix.length; i += EXTRACT_BATCH_SIZE) tasks.push({ chunk: tix.slice(i, i + EXTRACT_BATCH_SIZE), wf });
  }
  let extractionsByWf, exById, extractionErrors = 0;
  try {
    const useBatch = document.getElementById('useBatch').checked;
    const res = useBatch ? await runExtractionBatch(tasks, setStatus) : await runExtractionSync(tasks, setStatus);
    if (res.cancelled || cancelRequested) { setStatus('Reconciliation cancelled.'); btn.disabled = false; return; }
    ({ extractionsByWf, exById } = res);
    extractionErrors = res.errors || 0;
    if (extractionErrors > 0) {
      const unit = useBatch ? 'request' : 'batch';
      showError(`Warning: ${extractionErrors} of ${tasks.length} extraction ${unit}(s) failed — the affected tickets contribute no evidence, so some workflows may show less evidence or "Insufficient evidence". See the browser console for details.`);
    }
  } catch (err) {
    setStatus(''); btn.disabled = false;
    showError('Extraction failed: ' + describeError(err));
    return;
  }

  // Stage 2 — reconcile every rubric workflow (all get a status).
  reconciliation = new Map();
  const llmTasks = [];
  for (const wf of rubric) {
    if (!wf.name) continue;
    const exs = extractionsByWf.get(wf.name) || [];
    const total = totals.get(wf.name) || 0;
    if (exs.length < RECON_MIN_EVIDENCE) {
      reconciliation.set(wf.name, { status: 'Insufficient evidence', total, withEvidence: exs.length, observed_differences: '', roadblock: '', main_action: '', evidence_ticket_ids: [], suggested_description_update: '' });
    } else {
      llmTasks.push({ wf, agg: aggregateWorkflowExtractions(exs, total) });
    }
  }
  let ri = 0, rcDone = 0;
  const upRc = () => setStatus(`Reconciling workflows: ${rcDone}/${llmTasks.length}...`);
  upRc();
  async function rcWorker() {
    while (ri < llmTasks.length) {
      if (cancelRequested) return;
      const { wf, agg } = llmTasks[ri++];
      try { reconciliation.set(wf.name, await reconcileWorkflow(wf, agg)); }
      catch (err) {
        reconciliation.set(wf.name, { status: 'Insufficient evidence', total: agg.total, withEvidence: agg.withEvidence, observed_differences: '(reconcile failed: ' + describeError(err) + ')', roadblock: '', main_action: '', evidence_ticket_ids: [], suggested_description_update: '' });
      }
      rcDone++; upRc();
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, rcWorker));

  // Keep only the cited tickets' evidence (small subset) for rendering + persistence.
  citedEvidenceById = new Map();
  for (const v of reconciliation.values()) {
    for (const uid of (v.evidence_ticket_ids || [])) {
      const e = exById.get(uid);
      if (e && !citedEvidenceById.has(uid)) citedEvidenceById.set(uid, { en: e.evidence_en || '', nl: e.evidence_nl || '' });
    }
  }

  btn.disabled = false;
  const withEv = [...reconciliation.values()].filter(v => v.status !== 'Insufficient evidence').length;
  setStatus(`Reconciliation complete — ${withEv} workflow(s) with enough evidence, ${reconciliation.size} total.`, true);
  render();
  saveEnrichedState();
}

// --- Coverage-gap detection (§2.3): flag when too many tickets fit no workflow ---

function computeCoverage(rows, records) {
  const total = rows.length;
  const unclassified = rows.filter(r => r.workflow === CATCHALL_WORKFLOW).length;
  const noNotes = records.filter(r => !r.has_notes).length;
  return { total, unclassified, pct: total > 0 ? unclassified / total : 0, noNotes, themes: [] };
}

const THEME_SCHEMA = {
  type: "object",
  properties: {
    themes: {
      type: "array",
      items: {
        type: "object",
        properties: {
          theme: { type: "string" },
          description: { type: "string" },
          example_ticket_numbers: { type: "array", items: { type: "string" } },
        },
        required: ["theme", "description", "example_ticket_numbers"],
        additionalProperties: false,
      },
    },
  },
  required: ["themes"],
  additionalProperties: false,
};

// Summarize recurring themes among unclassified tickets — candidate new workflows (§2.3).
// Samples up to THEME_SAMPLE_SIZE unclassified tickets that actually have notes.
async function summarizeUnclassifiedThemes(records) {
  const withNotes = records.filter(r => r.workflow === CATCHALL_WORKFLOW && r.has_notes);
  if (!withNotes.length) return [];
  const sample = withNotes.slice(0, THEME_SAMPLE_SIZE);
  const bundles = sample.map(r => `Ticket ${r.ticket_id}: ${r.notes.replace(/\s+/g, ' ').slice(0, 600)}`).join('\n\n');
  const system = `These support tickets could NOT be matched to any workflow in the rubric. ` +
    `Identify the recurring themes among them — these are candidate new workflows the rubric is missing. ` +
    `Notes are mostly Dutch; reason over them directly and write all output in English. ` +
    `Return 3–8 themes, each with a short English name, a one-sentence description, and 2–3 example ticket numbers drawn only from the input.`;
  const user = `Unclassified tickets:\n\n${bundles}`;
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL_SUMMARY, max_tokens: 2048, system, user, schema: THEME_SCHEMA }),
  });
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Theme summary failed');
  return (JSON.parse(body.text).themes) || [];
}

function describeError(err) {
  return (err && err.message) ? err.message : String(err);
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Build the per-ticket classification inputs from the grouped records. The signal is
// the concatenated engineer notes (mostly Dutch), optionally preceded by the ticket-export
// fields the user checked in the classify panel (title / issue type / description).

// Read the classify-panel checkboxes that pick which fields go into the classifier text.
// Falls back to the shipped defaults (notes + description on) when a checkbox is absent.
function classifyTextOptions() {
  const on = (id, dflt) => { const el = document.getElementById(id); return el ? el.checked : dflt; };
  return {
    notes: on('txtNotes', true),
    description: on('txtDescription', true),
    title: on('txtTitle', false),
    issueType: on('txtIssueType', false),
    subIssueType: on('txtSubIssueType', false),
  };
}

function buildTicketsForClassification() {
  const opt = classifyTextOptions();
  return ticketRecords.map(rec => {
    // Ticket context first (title / issue type / description), then label+contract tags,
    // then the engineer notes — mirrors how a human reads the ticket top-down.
    const parts = [];
    if (opt.title && rec.title) parts.push(`Title: ${rec.title}`);
    const issueBits = [opt.issueType ? rec.issue_type : '', opt.subIssueType ? rec.sub_issue_type : ''].filter(Boolean);
    if (issueBits.length) parts.push(`Issue type: ${issueBits.join(' / ')}`);
    if (opt.description && rec.description) parts.push(`Description: ${rec.description}`);
    // Labels (KPI level 1) and Contract type are deliberately NOT injected here: they're
    // billing/commercial metadata, not descriptive of the work, so they dilute the prompt.
    // rec.labels / rec.contract_types are still kept for the Non-Billable + dashboard views.
    if (opt.notes && rec.notes) parts.push(rec.notes);
    return {
      ticket_id: rec.ticket_id,
      company: rec.company,
      text: parts.length ? parts.join('\n') : '(no notes)',
      hours: rec.hours,
      touches: rec.touches,
      firstTouchResolved: null,
      ticketCount: 1,
      record: rec,
    };
  });
}

function chunkTickets(tickets) {
  const chunks = [];
  for (let i = 0; i < tickets.length; i += BATCH_SIZE) chunks.push(tickets.slice(i, i + BATCH_SIZE));
  return chunks;
}

// Build dashboardRows from per-ticket classifications and stamp workflow/category
// back onto the records (so the non-billable pivot is workflow-aware). Shared by the
// classify paths and the restore-from-localStorage path.
function applyClassifications(tickets, classifications) {
  const catByName = new Map(rubric.map(r => [r.name, r.category || 'General']));
  dashboardRows = tickets.map((t, i) => {
    const c = classifications[i];
    const workflow = (c && c.workflow && c.workflow !== 'NONE') ? c.workflow : CATCHALL_WORKFLOW;
    const category = catByName.get(workflow) || 'Other';
    if (t.record) { t.record.workflow = workflow; t.record.category = category; }
    let first_touch;
    if (t.firstTouchResolved != null) first_touch = t.firstTouchResolved;
    else first_touch = (t.ticketCount <= 1 && t.touches === 1) ? 1 : 0;
    return { company: t.company, workflow, category, hours: t.hours, touches: t.touches, first_touch, ticketCount: t.ticketCount };
  });
}

// Shared post-classification step for both the in-browser and Batch API paths:
// build dashboardRows, run the coverage-gap themes (§2.3), persist, and show the dashboard.
async function finalizeClassification(tickets, classifications) {
  reconciliation = null; citedEvidenceById = null; // a fresh classification invalidates prior reconciliation
  playbookRows = []; // a fresh classification invalidates prior Playbook rows (ticket-index alignment shifts)
  applyClassifications(tickets, classifications);

  coverageInfo = computeCoverage(dashboardRows, ticketRecords);
  if (coverageInfo.pct > COVERAGE_GAP_THRESHOLD) {
    document.getElementById('progressText').textContent = 'Summarizing recurring themes among unclassified tickets...';
    try {
      coverageInfo.themes = await summarizeUnclassifiedThemes(ticketRecords);
    } catch (err) {
      console.error('Theme summary failed', err);
      coverageInfo.themeError = describeError(err);
    }
  }
  saveEnrichedState(); // persist so the completed run survives a reload (Phase 6 slice)
  showDashboard();
}

// --- Persist / restore completed classification results (Phase 6 slice) ---
// Only the per-ticket workflow is stored (as an index into a names table) plus the
// rubric and coverage info; everything else (company/hours/touches/category) is
// reconstructed from the re-grouped records, keeping the payload well under the
// localStorage quota even at 117K tickets (~350KB).
const ENRICHED_LS_KEY = 'ent1998_enriched';

// Serializable snapshot of a completed run. Stores the per-ticket workflow as an
// index into a de-duplicated name list (compact for the full 117K dataset).
function buildEnrichedState() {
  const names = [];
  const idx = new Map();
  const wf = dashboardRows.map(r => {
    if (!idx.has(r.workflow)) { idx.set(r.workflow, names.length); names.push(r.workflow); }
    return idx.get(r.workflow);
  });
  return {
    fingerprint: datasetFingerprint(), names, wf, rubric, coverageInfo,
    reconciliation: reconciliation ? [...reconciliation.entries()] : null,
    citedEvidence: citedEvidenceById ? [...citedEvidenceById.entries()] : null,
  };
}
function saveEnrichedState() {
  try {
    localStorage.setItem(ENRICHED_LS_KEY, JSON.stringify(buildEnrichedState()));
  } catch (e) {
    console.warn('Could not persist enriched results (localStorage quota?):', e);
  }
}
function loadEnrichedState() { try { return JSON.parse(localStorage.getItem(ENRICHED_LS_KEY) || 'null'); } catch (e) { return null; } }

// Restore a completed run for the currently-grouped dataset (no API calls). Treats the
// stored state as untrusted: validates the fingerprint/shape and coerces any workflow
// name not in the current rubric to the catch-all.
function maybeRestoreEnriched() {
  const st = loadEnrichedState();
  if (!st || st.fingerprint !== datasetFingerprint()) return false;
  return applyEnrichedState(st);
}

// Apply a completed-run snapshot (from localStorage or an imported file) to the
// currently-grouped dataset. Treats `st` as untrusted: validates shape, requires
// the per-ticket wf[] to line up with ticketRecords (it maps by index), and
// coerces any workflow name not in the current rubric to the catch-all.
function applyEnrichedState(st) {
  if (!st || !Array.isArray(st.names) || !Array.isArray(st.wf) || st.wf.length !== ticketRecords.length) return false;

  if ((!rubric.length || (rubric.length === 1 && !rubric[0].name)) && Array.isArray(st.rubric) && st.rubric.length) {
    rubric = st.rubric;
    renderRubricTable();
  }
  const validNames = new Set(rubric.map(r => r.name));
  const tickets = buildTicketsForClassification();
  const classifications = st.wf.map(i => {
    const w = st.names[i];
    return { workflow: (w === CATCHALL_WORKFLOW || validNames.has(w)) ? w : CATCHALL_WORKFLOW };
  });
  playbookRows = []; // restored classification shifts ticket-index alignment — stale Playbook rows would misalign
  applyClassifications(tickets, classifications);
  coverageInfo = (st.coverageInfo && typeof st.coverageInfo === 'object') ? st.coverageInfo : null;
  reconciliation = Array.isArray(st.reconciliation) ? new Map(st.reconciliation) : null;
  citedEvidenceById = Array.isArray(st.citedEvidence) ? new Map(st.citedEvidence) : null;
  showError('');
  showDashboard();
  return true;
}

// --- Enriched-state file export/import (Phase 6): save an expensive completed
// run to a JSON file and reload it later without re-spending API budget. The
// snapshot maps workflows per-ticket by index, so import requires the identical
// grouped dataset loaded first (same fingerprint). ---
const ENRICHED_FILE_KIND = 'ent1998-enriched';

// Local timestamp to the minute for filenames, e.g. "2026-07-29_1046".
function fileTimestamp(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

function exportEnrichedState() {
  if (!dashboardRows.length) { showError('Run a classification first — there is no completed run to save.'); return; }
  const now = new Date();
  const payload = { _kind: ENRICHED_FILE_KIND, _version: 1, savedAt: now.toISOString(), ...buildEnrichedState() };
  const json = JSON.stringify(payload);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `${datasetFilePrefix()}run_${fileTimestamp(now)}.json`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  showToast('Run saved — reload it later with "Load Run (JSON)" on the same dataset.');
}

function importEnrichedState(file) {
  if (!file) return;
  if (!ticketRecords.length) { showError('Upload and group the same time-entry export first, then load the saved run.'); return; }
  const reader = new FileReader();
  reader.onload = () => {
    let st;
    try { st = JSON.parse(reader.result); } catch (e) { showError('That file is not valid JSON — pick a run file saved from this tool.'); return; }
    if (!st || typeof st !== 'object' || (st._kind && st._kind !== ENRICHED_FILE_KIND)) {
      showError('That JSON is not an ENT-1998 saved run.'); return;
    }
    if (st.fingerprint !== datasetFingerprint()) {
      showError('This saved run is for a different dataset. Load and group the exact time-entry export it was saved from, then try again.'); return;
    }
    if (!applyEnrichedState(st)) { showError('Could not apply the saved run — its shape does not match the current dataset.'); return; }
    saveEnrichedState(); // mirror the imported run into localStorage so a plain reload keeps it
    showToast('Saved run loaded — no API budget spent.');
  };
  reader.onerror = () => showError('Could not read that file.');
  reader.readAsText(file);
}

// In-browser classification (live, progressive). Best for small slices — the browser
// caps concurrent connections, so the full dataset should use the Batch API path below.
async function runClassification() {
  cancelRequested = false;
  showError('');
  if (!rubric.length) { showError('Add at least one workflow to the rubric.'); return; }
  if (!parsedRows.length) { showError('Upload a time-entry export first.'); return; }
  if (!ticketRecords.length && !runGrouping()) return;

  const tickets = buildTicketsForClassification();
  const validNames = new Set(rubric.map(r => r.name));
  const chunks = chunkTickets(tickets);
  const classifications = new Array(tickets.length).fill(null);
  let completed = 0, chunkIdx = 0, hadErrors = 0, firstErrorMessage = null;

  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('classifyBtn').disabled = true;
  const updateProgress = () => {
    const pct = Math.round((completed / chunks.length) * 100);
    document.getElementById('progressBarInner').style.width = pct + '%';
    document.getElementById('progressText').textContent =
      `Classifying batch ${completed} of ${chunks.length} (${tickets.length.toLocaleString()} tickets total)...`;
  };
  updateProgress();

  async function worker() {
    while (chunkIdx < chunks.length) {
      if (cancelRequested) return;
      const myIdx = chunkIdx++;
      const offset = myIdx * BATCH_SIZE;
      try {
        const result = await classifyBatch(chunks[myIdx], validNames);
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
  document.getElementById('classifyBtn').disabled = false;

  if (cancelRequested) {
    document.getElementById('progressWrap').style.display = 'none';
    showError('Classification cancelled.');
    return;
  }
  if (hadErrors === chunks.length) {
    document.getElementById('progressWrap').style.display = 'none';
    showError(`Every batch failed — nothing was classified. First error: ${firstErrorMessage}. Check the server console (running server.py) and try again.`);
    return;
  }
  if (hadErrors) showError(`${hadErrors} of ${chunks.length} batches failed and were left Unclassified. First error: ${firstErrorMessage}`);

  await finalizeClassification(tickets, classifications);
  document.getElementById('progressWrap').style.display = 'none';
}

// --- Batch API classification (async, ~50% cheaper, survives a closed tab) ---

const BATCH_LS_KEY = 'ent1998_pending_batch';
function datasetFingerprint() { return `${ticketRecords.length}:${parsedRows.length}:${headers.length}`; }
function saveBatchState(batchIds) {
  try { localStorage.setItem(BATCH_LS_KEY, JSON.stringify({ batch_ids: batchIds, fingerprint: datasetFingerprint() })); } catch (e) {}
}
function loadBatchState() { try { return JSON.parse(localStorage.getItem(BATCH_LS_KEY) || 'null'); } catch (e) { return null; } }
function clearBatchState() { try { localStorage.removeItem(BATCH_LS_KEY); } catch (e) {} }

// Split classification requests into groups small enough for the Batch API's 256MB limit.
// The system prompt is repeated per request in the batch JSONL, so each request's expanded
// weight includes it; we greedily fill a group until the next request would exceed the budget.
function partitionRequests(requests, systemText, schema) {
  const perReq = systemText.length + JSON.stringify(schema).length + 600; // system repeated + schema + envelope
  const groups = [];
  let cur = [], curBytes = 0;
  for (const r of requests) {
    const rBytes = perReq + (r.user ? r.user.length : 0) + (r.custom_id ? r.custom_id.length : 0);
    if (cur.length && curBytes + rBytes > MAX_BATCH_BYTES) { groups.push(cur); cur = []; curBytes = 0; }
    cur.push(r); curBytes += rBytes;
  }
  if (cur.length) groups.push(cur);
  return groups;
}

async function runBatchClassification() {
  cancelRequested = false;
  showError('');
  if (!rubric.length) { showError('Add at least one workflow to the rubric.'); return; }
  if (!parsedRows.length) { showError('Upload a time-entry export first.'); return; }
  if (!ticketRecords.length && !runGrouping()) return;

  const tickets = buildTicketsForClassification();
  const chunks = chunkTickets(tickets);
  const system = buildSystemPrompt();
  // custom_id encodes the GLOBAL chunk index, so results map back correctly no matter
  // which sub-batch they come from.
  const requests = chunks.map((chunk, i) => ({ custom_id: `c${i}`, user: buildUserPrompt(chunk) }));
  const groups = partitionRequests(requests, system, RESPONSE_SCHEMA);

  document.getElementById('classifyBtn').disabled = true;
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('progressBarInner').style.width = '0%';

  const batchIds = [];
  try {
    for (let gi = 0; gi < groups.length; gi++) {
      document.getElementById('progressText').textContent =
        `Submitting batch ${gi + 1} of ${groups.length} (${groups[gi].length.toLocaleString()} requests)...`;
      const resp = await fetch(`${API_BASE}/api/batch/create`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: MODEL, max_tokens: 4096, system, schema: RESPONSE_SCHEMA, requests: groups[gi] }),
      });
      const body = await resp.json();
      if (!body.ok) throw new Error(body.error || 'Batch create failed');
      batchIds.push(body.batch_id);
    }
  } catch (err) {
    // Best-effort cancel of any batches already created, so a mid-way failure doesn't leave orphaned spend.
    for (const id of batchIds) {
      try { await fetch(`${API_BASE}/api/batch/cancel`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ batch_id: id }) }); } catch (e) {}
    }
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('classifyBtn').disabled = false;
    showError('Could not start the batch: ' + describeError(err));
    return;
  }
  saveBatchState(batchIds);
  await pollBatchesToCompletion(batchIds, tickets, chunks);
}

// A batch status/results error that will never succeed on retry (bad/stale/garbage id),
// vs a transient network blip. Used to break the poll loop instead of retrying forever.
function isPermanentBatchError(err) {
  return typeof err === 'string' && /invalid|msgbatch|not[ _]?found|\b400\b/i.test(err);
}

// Poll every batch of a run to completion, then merge results back by custom_id and finalize.
// Reused by the resume path after a page reload.
async function pollBatchesToCompletion(batchIds, tickets, chunks) {
  const validNames = new Set(rubric.map(r => r.name));
  const total = chunks.length;
  const bar = document.getElementById('progressBarInner');
  const text = document.getElementById('progressText');
  const cancelBtn = document.getElementById('cancelBtn');
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('classifyBtn').disabled = true;
  activeBatchIds = batchIds.slice();   // enables the Cancel-batch handler
  cancelBtn.textContent = 'Cancel batch';
  const endPoll = () => {
    activeBatchIds = [];
    cancelBtn.textContent = 'Cancel';
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('classifyBtn').disabled = false;
  };

  while (true) {
    if (cancelRequested) { endPoll(); return; } // the Cancel-batch handler owns the cancel call + message
    let done = 0, allEnded = true, badStatus = null, anyError = false, fatalError = null;
    for (const id of batchIds) {
      let status;
      try {
        const resp = await fetch(`${API_BASE}/api/batch/status?id=${encodeURIComponent(id)}`);
        status = await resp.json();
      } catch (err) { status = { ok: false, error: describeError(err) }; }
      if (status.ok) {
        const c = status.request_counts || {};
        done += (c.succeeded || 0) + (c.errored || 0) + (c.canceled || 0) + (c.expired || 0);
        if (status.processing_status !== 'ended') allEnded = false;
        if (['canceled', 'canceling', 'expired'].includes(status.processing_status)) badStatus = status.processing_status;
      } else {
        allEnded = false; anyError = true;
        if (isPermanentBatchError(status.error)) fatalError = status.error;
      }
    }
    const pct = total > 0 ? Math.round(done / total * 100) : 0;
    bar.style.width = pct + '%';
    // A permanent error (invalid/stale id) will never recover — stop and clear the state
    // instead of looping "retrying..." forever.
    if (fatalError) { endPoll(); clearBatchState(); showError(`Stopped watching the batch — invalid or stale batch reference (${fatalError}). Re-run classification.`); return; }
    if (badStatus) { endPoll(); clearBatchState(); showError(`Batch ${badStatus}.`); return; }
    text.textContent = anyError
      ? `Batch status check failed — retrying...`
      : `Batch classifying: ${done.toLocaleString()} / ${total.toLocaleString()} requests (${pct}%) across ${batchIds.length} batch${batchIds.length === 1 ? '' : 'es'}. Safe to close the tab.`;
    if (allEnded) break;
    await sleep(5000);
  }

  text.textContent = 'Batches complete — downloading results...';
  const classifications = new Array(tickets.length).fill(null);
  try {
    for (const id of batchIds) {
      const resp = await fetch(`${API_BASE}/api/batch/results?id=${encodeURIComponent(id)}`);
      const body = await resp.json();
      if (!body.ok) throw new Error(body.error || 'Batch results failed');
      if (body.usage) console.log('Batch cache usage:', id, body.usage);
      for (const r of body.results || []) {
        const m = /^c(\d+)$/.exec(r.custom_id || '');
        if (!m) continue;
        const chunkIdx = +m[1];
        const chunk = chunks[chunkIdx];
        if (!chunk || !r.ok) continue;
        const offset = chunkIdx * BATCH_SIZE;
        let parsed;
        try { parsed = JSON.parse(r.text); } catch (e) { continue; }
        for (const item of parsed.classifications || []) {
          if (item.index == null || item.index >= chunk.length) continue;
          let wf = item.workflow;
          if (wf !== 'NONE' && !validNames.has(wf)) wf = 'NONE';
          classifications[offset + item.index] = { workflow: wf, confidence: item.confidence };
        }
      }
    }
  } catch (err) {
    endPoll();
    showError('Could not download batch results: ' + describeError(err));
    return;
  }

  clearBatchState();
  activeBatchIds = [];                        // run done — disable the cancel path
  cancelBtn.textContent = 'Cancel';
  await finalizeClassification(tickets, classifications); // keeps progressWrap visible for the theme step
  document.getElementById('progressWrap').style.display = 'none';
  document.getElementById('classifyBtn').disabled = false;
}

// If a batch is pending for the currently-grouped dataset, resume watching it.
function maybeResumeBatch() {
  const st = loadBatchState();
  if (!st || st.fingerprint !== datasetFingerprint()) return false;
  const ids = (st.batch_ids || (st.batch_id ? [st.batch_id] : [])) // legacy single-id fallback
    .filter(id => typeof id === 'string' && id.startsWith('msgbatch_'));
  if (!ids.length) { clearBatchState(); return false; } // stale/garbage entry — don't auto-resume it
  showError('');
  const tickets = buildTicketsForClassification();
  pollBatchesToCompletion(ids, tickets, chunkTickets(tickets));
  return true;
}

document.getElementById('classifyBtn').addEventListener('click', () => {
  if (document.getElementById('useBatch').checked) runBatchClassification();
  else runClassification();
});
document.getElementById('cancelBtn').addEventListener('click', async () => {
  cancelRequested = true; // stops the in-browser worker loop and the batch poll loop
  const ids = activeBatchIds;
  if (!ids || !ids.length) return; // in-browser mode: runClassification shows "Classification cancelled."
  // Batch mode: actually cancel every batch of the run on Anthropic (not just stop watching).
  activeBatchIds = [];
  document.getElementById('cancelBtn').textContent = 'Cancel';
  document.getElementById('progressText').textContent = `Cancelling ${ids.length} batch${ids.length === 1 ? '' : 'es'}...`;
  try {
    for (const id of ids) {
      const resp = await fetch(`${API_BASE}/api/batch/cancel`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ batch_id: id }),
      });
      const body = await resp.json();
      if (!body.ok) throw new Error(body.error || 'cancel failed');
    }
    showError('Batch cancelled.');
  } catch (err) {
    showError(`Could not cancel the batch: ${describeError(err)}`);
  }
  clearBatchState();
  document.getElementById('progressWrap').style.display = 'none';
  document.getElementById('classifyBtn').disabled = false;
});

document.getElementById('startOverBtn').addEventListener('click', () => {
  document.getElementById('dashboardSection').style.display = 'none';
  document.getElementById('setupSection').style.display = 'block';
});

// --- Dashboard (aggregation, filtering, rendering) ---

const companyFilter = document.getElementById('companyFilter');
const tierFilterEl = document.getElementById('tierFilter');
const reconFilterEl = document.getElementById('reconFilter');

// Keep the reconciliation-status filter one row per aggregated workflow.
// `all` passes everything; `none` matches workflows with no verdict yet;
// otherwise match the verdict's status string exactly.
function reconFilterPredicate(workflow) {
  const rf = reconFilterEl.value;
  if (rf === 'all') return true;
  const v = reconciliation && reconciliation.get(workflow);
  return rf === 'none' ? !v : !!v && v.status === rf;
}
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

// Rubric coverage-gap warning banner (§2.3), shown above the workflow dashboard.
function renderCoverageBanner() {
  const el = document.getElementById('coverageBanner');
  if (!coverageInfo || coverageInfo.pct <= COVERAGE_GAP_THRESHOLD) { el.style.display = 'none'; return; }
  const pct = (coverageInfo.pct * 100).toFixed(1);
  const noNotesNote = coverageInfo.noNotes
    ? ` ${coverageInfo.noNotes.toLocaleString()} ticket(s) had no notes ("no evidence").` : '';
  let themesHtml = '';
  if (coverageInfo.themes && coverageInfo.themes.length) {
    themesHtml = '<div style="margin-top:8px;"><strong>Recurring themes among unclassified tickets (candidate new workflows):</strong><ul style="margin:6px 0 0 18px;">' +
      coverageInfo.themes.map(t => {
        const ex = (t.example_ticket_numbers || []).slice(0, 3).map(xmlEscape).join(', ');
        return `<li><strong>${xmlEscape(t.theme)}</strong> — ${xmlEscape(t.description)}${ex ? ` <span style="color:var(--text-dim);">(e.g. ${ex})</span>` : ''}</li>`;
      }).join('') + '</ul></div>';
  } else if (coverageInfo.themeError) {
    themesHtml = `<div style="margin-top:8px; color:var(--text-dim);">Theme summary unavailable: ${xmlEscape(coverageInfo.themeError)}</div>`;
  }
  el.className = 'error-banner';
  el.style.display = 'block';
  el.innerHTML = `⚠️ <strong>Rubric coverage gap:</strong> ${coverageInfo.unclassified.toLocaleString()} of ${coverageInfo.total.toLocaleString()} tickets (${pct}%) matched no workflow — above the ${(COVERAGE_GAP_THRESHOLD * 100).toFixed(0)}% threshold.${noNotesNote}${themesHtml}`;
}

function render() {
  renderCoverageBanner();
  const filteredRaw = getFilteredRows();
  let agg = aggregateRows(filteredRaw);

  const tier = tierFilterEl.value;
  if (tier !== 'all') agg = agg.filter(r => r.tier === tier);

  // The reconciliation filter only makes sense once verdicts exist; hide it
  // (and neutralize a stale selection) until reconciliation has run.
  const hasRecon = reconciliation && reconciliation.size > 0;
  document.getElementById('reconFilterWrap').style.display = hasRecon ? '' : 'none';
  if (!hasRecon) reconFilterEl.value = 'all';
  if (reconFilterEl.value !== 'all') agg = agg.filter(r => reconFilterPredicate(r.workflow));

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

  // Reconciliation-status summary — a strip of cards counting workflows by
  // status across the currently-filtered view. Shown only once reconciliation
  // has run. Colors mirror the badge palette (aligned=green ... changed=red).
  const reconSummaryEl = document.getElementById('reconSummaryStats');
  if (hasRecon) {
    const rc = { 'Aligned': 0, 'Minor drift': 0, 'Significantly changed': 0, 'Insufficient evidence': 0 };
    let notReconciled = 0;
    for (const r of agg) {
      const v = reconciliation.get(r.workflow);
      if (v && rc[v.status] != null) rc[v.status]++;
      else if (!v) notReconciled++;
    }
    const reconWfPct = n => agg.length > 0 ? (n / agg.length * 100) : 0;
    const card = (n, label, cls) => `<div class="stat ${cls}"><div class="num">${n.toLocaleString()}</div><div class="lbl">${label}</div></div>`;
    reconSummaryEl.innerHTML =
      card(agg.length, 'Workflows Reconciled', '') +
      card(rc['Aligned'], `Aligned (${reconWfPct(rc['Aligned']).toFixed(0)}%)`, 'tier-stat tier-stat-low') +
      card(rc['Minor drift'], `Minor Drift (${reconWfPct(rc['Minor drift']).toFixed(0)}%)`, 'tier-stat tier-stat-medium') +
      card(rc['Significantly changed'], `Significantly Changed (${reconWfPct(rc['Significantly changed']).toFixed(0)}%)`, 'tier-stat tier-stat-high') +
      card(rc['Insufficient evidence'], 'Insufficient Evidence', 'tier-stat tier-stat-neutral') +
      (notReconciled ? card(notReconciled, 'Not Reconciled', 'tier-stat tier-stat-neutral') : '');
    reconSummaryEl.style.display = '';
  } else {
    reconSummaryEl.style.display = 'none';
    reconSummaryEl.innerHTML = '';
  }

  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  tableBody.innerHTML = agg.map(r => `
    <tr>
      <td class="cat-name"><div class="cat-cell">${xmlEscape(r.category)}</div></td>
      <td class="wf-name"><div class="wf-cell">${xmlEscape(r.workflow)}${descByName.get(r.workflow) ? `<div class="wf-desc">${xmlEscape(descByName.get(r.workflow))}</div>` : ''}</div></td>
      <td class="num-cell">${r.tickets.toLocaleString()}</td>
      <td class="num-cell">${r.ticketPct.toFixed(1)}%</td>
      <td class="num-cell">${r.hours.toLocaleString(undefined, {maximumFractionDigits: 1})}</td>
      <td class="num-cell">${r.hoursPct.toFixed(1)}%</td>
      <td class="num-cell">${r.aht.toFixed(1)}</td>
      <td class="num-cell">${(r.frr * 100).toFixed(0)}%</td>
      <td class="num-cell">${r.touches.toFixed(2)}</td>
      <td><span class="tier tier-${r.tier}">${r.tier}</span></td>
      ${renderReconCells(r.workflow)}
    </tr>
  `).join('');

  window.__currentAgg = agg;
}

// Reconciliation cells for one workflow row (Feature A §2.2). All values are
// LLM-generated and are escaped via xmlEscape() before insertion.
function renderReconCells(workflowName) {
  const empty = '<td class="recon-empty">—</td>';
  const v = reconciliation && reconciliation.get(workflowName);
  if (!v) return empty + empty + empty + empty;

  const cls = { 'Aligned': 'aligned', 'Minor drift': 'minor', 'Significantly changed': 'changed', 'Insufficient evidence': 'insufficient' }[v.status] || 'insufficient';
  const sub = v.withEvidence != null ? `<div class="recon-sub">${v.withEvidence} w/ notes</div>` : '';
  const statusCell = `<td class="recon-td"><span class="recon-badge recon-${cls}">${xmlEscape(v.status)}</span>${sub}</td>`;
  const diffCell = `<td class="recon-td"><div class="recon-cell">${xmlEscape(v.observed_differences || '')}</div></td>`;

  let rmCell;
  if (v.status === 'Insufficient evidence' || (!v.roadblock && !v.main_action)) {
    rmCell = empty;
  } else {
    const cites = (v.evidence_ticket_ids || []).map(uid => {
      const ev = citedEvidenceById && citedEvidenceById.get(uid);
      if (ev && ev.nl) {
        return `<details class="recon-ev"><summary>${xmlEscape(uid)}</summary>` +
          `<div class="recon-en">${xmlEscape(ev.en || '')}</div>` +
          `<div class="recon-nl">NL: ${xmlEscape(ev.nl)}</div></details>`;
      }
      return `<span class="recon-cite">${xmlEscape(uid)}</span>`;
    }).join('');
    rmCell = `<td class="recon-td"><div class="recon-cell">` +
      `<div><strong>Roadblock:</strong> ${xmlEscape(v.roadblock || '—')}</div>` +
      `<div><strong>Main action:</strong> ${xmlEscape(v.main_action || '—')}</div>` +
      (cites ? `<div class="recon-cites">${cites}</div>` : '') + `</div></td>`;
  }
  const updCell = `<td class="recon-td"><div class="recon-cell">${v.suggested_description_update ? xmlEscape(v.suggested_description_update) : ''}</div></td>`;
  return statusCell + diffCell + rmCell + updCell;
}

document.querySelectorAll('table.dash thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    if (!key) return; // reconciliation columns are not sortable
    if (sortKey === key) sortDir *= -1; else { sortKey = key; sortDir = -1; }
    render();
  });
});
companyFilter.addEventListener('change', render);
tierFilterEl.addEventListener('change', render);
reconFilterEl.addEventListener('change', render);
document.getElementById('runReconBtn').addEventListener('click', runReconciliation);

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
  const isPb = view === 'playbook';
  document.getElementById('workflowView').style.display = (isNb || isPb) ? 'none' : 'block';
  document.getElementById('nonBillableView').style.display = isNb ? 'block' : 'none';
  document.getElementById('playbookView').style.display = isPb ? 'block' : 'none';
  // Workflow-specific export buttons only make sense on the workflow view (NB/Playbook export their own way).
  document.getElementById('exportCsvBtn').style.display = (isNb || isPb) ? 'none' : '';
  document.getElementById('exportBtn').style.display = (isNb || isPb) ? 'none' : '';
  document.getElementById('viewWorkflowBtn').className = 'btn small' + (isNb || isPb ? ' secondary' : '');
  document.getElementById('viewNonBillableBtn').className = 'btn small' + (isNb ? '' : ' secondary');
  document.getElementById('viewPlaybookBtn').className = 'btn small' + (isPb ? '' : ' secondary');
  if (isNb) renderNonBillable();
  if (isPb) renderPlaybook();
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
document.getElementById('viewPlaybookBtn').addEventListener('click', () => {
  if (!dashboardRows.length) { showError('Classify tickets first to build the Playbook view.'); return; }
  showError('');
  setDashboardView('playbook');
});

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

// --- Playbook view: Company x Workflow rows with description, common steps,
// responsible engineer tier, and the existing volume/quality metrics. ---

const TIER_SCHEMA = {
  type: "object",
  properties: {
    mappings: {
      type: "array",
      items: {
        type: "object",
        properties: {
          queue_label: { type: "string" },
          tier: { type: "string", enum: TIER_LEVELS },
        },
        required: ["queue_label", "tier"],
        additionalProperties: false,
      },
    },
  },
  required: ["mappings"],
  additionalProperties: false,
};

function buildTierSystem() {
  return `You are normalizing an MSP's raw support-ticket queue/team labels (mostly Dutch, inconsistent per company) into a standard support-tier taxonomy. For each distinct label given, pick the single best-matching bucket from: ${TIER_LEVELS.join(', ')}.

Guidance:
- "1e lijn" / "1ste lijn" / "Servicedesk 1e lijn" / "Level I" / first-line service-desk labels -> Tier 1 (Front-line).
- "2e lijn" / "2de lijn" -> Tier 2 (Escalation).
- "3e lijn" / "System Engineering" -> Tier 3 (Specialist Engineering).
- "SOC" / "Security" / "Monitoring" / "Vulnerability" / "Back-up" alerting queues -> Security / SOC.
- "Beheer" / "Beheerafspraak" / "Administratie" / "Planning" / "Intake/dispatching" -> Management / Admin.
- "Sales" / "Inside Sales" / "Post Sale" -> Sales / Account.
- Numbered color codes (e.g. "2. Groen", "3. Blauw", "4. Rood") are a company-specific severity/tier ladder, not a literal team name — if the sequence clearly implies ascending tiers, map low numbers to Tier 1 and high numbers to Tier 3; otherwise use Other / Unclear.
- Anything you cannot confidently place -> Other / Unclear. Do not guess beyond what the label plausibly means.

Echo each queue_label EXACTLY as given. Return exactly one mapping per input label.`;
}

async function normalizeQueueTiers(labels) {
  if (!labels.length) return new Map();
  const user = `Distinct queue/team labels (${labels.length}):\n${labels.map(l => `- ${l}`).join('\n')}`;
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL_SUMMARY, max_tokens: 4096, system: buildTierSystem(), user, schema: TIER_SCHEMA }),
  });
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Tier normalization failed');
  const parsed = JSON.parse(body.text);
  const known = new Set(labels);
  const map = new Map();
  for (const m of parsed.mappings || []) {
    if (known.has(m.queue_label) && TIER_LEVELS.includes(m.tier)) map.set(m.queue_label, m.tier);
  }
  return map;
}

function ticketTier(rec) {
  if (!rec || !rec.queue_label || !tierByQueueLabel) return 'Unknown';
  return tierByQueueLabel.get(rec.queue_label) || 'Unknown';
}

// Group classified tickets by (company, workflow). Mirrors aggregateRows() (which groups
// by workflow only) but adds the dominant Responsible Engineer Tier per group and excludes
// the catch-all bucket (no rubric description/steps make sense for "Unclassified / Other").
// Also returns, per group, the ticketRecords indices needed to sample notes for steps synthesis.
function aggregatePlaybookGroups() {
  const byKey = new Map(); // "company␟workflow" -> group
  for (let i = 0; i < dashboardRows.length; i++) {
    const r = dashboardRows[i];
    if (r.workflow === CATCHALL_WORKFLOW) continue;
    const key = r.company + '␟' + r.workflow;
    let g = byKey.get(key);
    if (!g) {
      g = { company: r.company, workflow: r.workflow, category: r.category, tickets: 0, hours: 0, touches: 0, first_touch: 0, tierCounts: {}, recordIdx: [] };
      byKey.set(key, g);
    }
    g.tickets += r.ticketCount; g.hours += r.hours; g.touches += r.touches; g.first_touch += r.first_touch;
    const tier = ticketTier(ticketRecords[i]);
    g.tierCounts[tier] = (g.tierCounts[tier] || 0) + 1;
    g.recordIdx.push(i);
  }
  const out = [];
  for (const g of byKey.values()) {
    const aht = g.tickets > 0 ? (g.hours / g.tickets) * 60 : 0;
    const frr = g.tickets > 0 ? g.first_touch / g.tickets : 0;
    const touchesPerTicket = g.tickets > 0 ? g.touches / g.tickets : 0;
    let tier = 'Unknown', tierN = -1;
    for (const [t, n] of Object.entries(g.tierCounts)) if (n > tierN) { tier = t; tierN = n; }
    out.push({ company: g.company, workflow: g.workflow, category: g.category, tickets: g.tickets, hours: g.hours, aht, frr, touches: touchesPerTicket, tier, recordIdx: g.recordIdx });
  }
  return out;
}

const STEPS_SCHEMA = {
  type: "object",
  properties: {
    common_steps: { type: "array", items: { type: "string" }, minItems: 3, maxItems: 5 },
  },
  required: ["common_steps"],
  additionalProperties: false,
};

function buildStepsSystem() {
  return `You are summarizing how engineers ACTUALLY resolve helpdesk tickets in one workflow, for one company, based on raw time-entry notes (mostly Dutch — reason over them directly, do NOT pre-translate). Identify the common sequence of steps engineers take across these tickets and return between 3 and 5 concise, ordered, English steps (e.g. "Verify user identity", "Reset password in Entra ID", "Confirm resolution with customer"). Base this ONLY on patterns visible across MULTIPLE tickets in the input — never invent a step that isn't evidenced in the notes. If the tickets show genuinely different processes, describe the single most common path.`;
}

function buildStepsUser(company, wf, sampleNotes) {
  const bundles = sampleNotes.map((n, i) => `Ticket ${i + 1}:\n${n.replace(/\s+/g, ' ').slice(0, 800)}`).join('\n\n');
  return `Company: ${company}\nWorkflow: "${wf.name}"\nWorkflow description: "${wf.description}"\n\nSample ticket notes (${sampleNotes.length} of them):\n\n${bundles}`;
}

async function synthesizeSteps(company, wf, sampleNotes) {
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, max_tokens: 1024, system: buildStepsSystem(), user: buildStepsUser(company, wf, sampleNotes), schema: STEPS_SCHEMA }),
  });
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Steps synthesis failed');
  return (JSON.parse(body.text).common_steps) || [];
}

// Orchestrate the Playbook build: normalize queue-label tiers (once, if ticket metadata was
// attached), aggregate by (company, workflow), then synthesize common steps per group with
// enough evidence (concurrency-limited worker pool, mirrors runReconciliation's shape).
async function runBuildPlaybook() {
  if (!dashboardRows.length) { showError('Classify tickets first, then build the Playbook.'); return; }
  showError('');
  const statusEl = document.getElementById('playbookStatus');
  const btn = document.getElementById('buildPlaybookBtn');
  btn.disabled = true;
  const setStatus = (msg, done) => { statusEl.textContent = msg; statusEl.style.color = done ? 'var(--low)' : 'var(--text-dim)'; };

  if (!tierByQueueLabel) {
    const labels = Array.from(new Set(ticketRecords.map(r => r.queue_label).filter(Boolean)));
    if (labels.length) {
      setStatus(`Normalizing ${labels.length} queue label(s) into support tiers...`);
      try { tierByQueueLabel = await normalizeQueueTiers(labels); }
      catch (err) { showError('Tier normalization failed (tiers will show "Unknown"): ' + describeError(err)); tierByQueueLabel = new Map(); }
    } else {
      tierByQueueLabel = new Map(); // no ticket metadata attached — every group shows "Unknown"
    }
  }

  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  const groups = aggregatePlaybookGroups();
  const tasks = groups
    .map((g, gi) => ({ g, gi, notes: g.recordIdx.map(i => ticketRecords[i]).filter(r => r.has_notes) }))
    .filter(t => t.notes.length >= PLAYBOOK_MIN_EVIDENCE);
  for (const g of groups) { g.steps = []; g.stepsStatus = 'Insufficient evidence'; }

  let ti = 0, done = 0;
  const upSteps = () => setStatus(`Synthesizing common steps: ${done}/${tasks.length} groups (${groups.length - tasks.length} skipped — insufficient evidence)...`);
  upSteps();
  async function worker() {
    while (ti < tasks.length) {
      const t = tasks[ti++];
      const wf = { name: t.g.workflow, description: descByName.get(t.g.workflow) || '' };
      const sample = t.notes.slice(0, PLAYBOOK_SAMPLE).map(r => r.notes);
      try {
        t.g.steps = await synthesizeSteps(t.g.company, wf, sample);
        t.g.stepsStatus = 'Generated';
      } catch (err) {
        t.g.stepsStatus = 'Failed: ' + describeError(err);
      }
      done++; upSteps();
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, worker));

  playbookRows = groups.map(g => ({ ...g, description: descByName.get(g.workflow) || '' }));
  btn.disabled = false;
  setStatus(`Playbook built — ${playbookRows.length} company × workflow row(s), ${tasks.length} with synthesized steps.`, true);
  renderPlaybook();
}

let playbookSortKey = 'tickets';
let playbookSortDir = -1;

function populatePlaybookCompanyFilter() {
  const el = document.getElementById('playbookCompanyFilter');
  const prev = el.value;
  const companies = Array.from(new Set(playbookRows.map(r => r.company))).sort();
  el.innerHTML = '<option value="all">All Companies</option>' + companies.map(c => `<option value="${escapeAttr(c)}">${xmlEscape(c)}</option>`).join('');
  el.value = companies.includes(prev) ? prev : 'all';
}

const PLAYBOOK_TIER_CLASS = {
  'Tier 1 (Front-line)': 'tier-low', 'Tier 2 (Escalation)': 'tier-medium', 'Tier 3 (Specialist Engineering)': 'tier-high',
};

function renderPlaybook() {
  populatePlaybookCompanyFilter();
  const company = document.getElementById('playbookCompanyFilter').value;
  let rows = company === 'all' ? playbookRows : playbookRows.filter(r => r.company === company);
  rows = rows.slice().sort((a, b) => {
    const va = a[playbookSortKey], vb = b[playbookSortKey];
    if (typeof va === 'string') return va.localeCompare(vb) * playbookSortDir;
    return (va - vb) * playbookSortDir;
  });
  document.getElementById('playbookTableBody').innerHTML = rows.map(r => {
    const tierCls = PLAYBOOK_TIER_CLASS[r.tier] || '';
    const stepsHtml = r.steps && r.steps.length
      ? '<ol class="steps-list">' + r.steps.map(s => `<li>${xmlEscape(s)}</li>`).join('') + '</ol>'
      : `<span class="steps-empty">${xmlEscape(r.stepsStatus || 'Insufficient evidence')}</span>`;
    return `<tr>
      <td>${xmlEscape(r.company)}</td>
      <td class="wf-name"><div class="wf-cell">${xmlEscape(r.workflow)}</div></td>
      <td><div class="wf-desc" style="max-width:260px;">${xmlEscape(r.description || '')}</div></td>
      <td><div class="steps-cell">${stepsHtml}</div></td>
      <td>${tierCls ? `<span class="tier ${tierCls}">${xmlEscape(r.tier)}</span>` : `<span class="tier tier-neutral">${xmlEscape(r.tier)}</span>`}</td>
      <td class="num-cell">${r.tickets.toLocaleString()}</td>
      <td class="num-cell">${r.hours.toLocaleString(undefined, {maximumFractionDigits: 1})}</td>
      <td class="num-cell">${(r.frr * 100).toFixed(0)}%</td>
      <td class="num-cell">${r.touches.toFixed(2)}</td>
      <td class="num-cell">${r.aht.toFixed(1)}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="10" style="text-align:center; color:var(--text-dim);">No Playbook rows yet — click "Build / Refresh Playbook".</td></tr>`;
}

document.querySelectorAll('#playbookTable thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.pbkey;
    if (!key) return;
    if (playbookSortKey === key) playbookSortDir *= -1; else { playbookSortKey = key; playbookSortDir = -1; }
    renderPlaybook();
  });
});
document.getElementById('playbookCompanyFilter').addEventListener('change', renderPlaybook);
document.getElementById('buildPlaybookBtn').addEventListener('click', runBuildPlaybook);

// --- Export: clipboard TSV ---

function buildTSV() {
  const headersRow = ['Category', 'Workflow', 'Tickets', 'Total Hours', 'AHT (min)', 'FRR', 'Touches/Ticket', 'Complexity',
    'Reconciliation Status', 'Observed Differences', 'Roadblock', 'Main Action', 'Cited Tickets', 'Suggested Update'];
  const lines = [headersRow.join('\t')];
  const clean = s => String(s == null ? '' : s).replace(/[\t\r\n]+/g, ' '); // keep TSV one-row-per-record
  for (const r of window.__currentAgg || []) {
    const v = reconciliation && reconciliation.get(r.workflow);
    const rc = v
      ? [v.status, v.observed_differences, v.roadblock, v.main_action, citedTicketsList(v).join(' | '), v.suggested_description_update]
      : ['', '', '', '', '', ''];
    lines.push([r.category, r.workflow, r.tickets, r.hours.toFixed(1), r.aht.toFixed(1), (r.frr * 100).toFixed(0) + '%', r.touches.toFixed(2), r.tier, ...rc].map(clean).join('\t'));
  }
  return lines.join('\n');
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

document.getElementById('saveRunBtn').addEventListener('click', exportEnrichedState);
document.getElementById('loadRunBtn').addEventListener('click', () => document.getElementById('loadRunFile').click());
document.getElementById('loadRunFile').addEventListener('change', (e) => {
  const f = e.target.files && e.target.files[0];
  importEnrichedState(f);
  e.target.value = ''; // allow re-selecting the same file
});

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
  if (reconFilterEl.value !== 'all') agg = agg.filter(r => reconFilterPredicate(r.workflow));
  return sortAggRows(agg);
}

const _xStr = v => `<Cell><Data ss:Type="String">${xmlEscape(v == null ? '' : v)}</Data></Cell>`;
const _xNum = v => `<Cell><Data ss:Type="Number">${Number(v || 0)}</Data></Cell>`;
const _xHdr = v => `<Cell ss:StyleID="Header"><Data ss:Type="String">${xmlEscape(v)}</Data></Cell>`;

// Each cited ticket as "english, dutch (uid)" — falls back to the uid alone when
// no evidence was captured. Shared by the .xls and TSV exports.
function citedTicketsList(v) {
  return (v.evidence_ticket_ids || []).map(uid => {
    const e = citedEvidenceById && citedEvidenceById.get(uid);
    const en = (e && e.en) ? e.en : '';
    const nl = (e && e.nl) ? e.nl : '';
    return (en || nl) ? `${en}, ${nl} (${uid})` : uid;
  });
}

// Reconciliation cells for the Workflow Dashboard sheet (empty until reconciliation runs).
function reconSheetCells(workflow) {
  const v = reconciliation && reconciliation.get(workflow);
  if (!v) return _xStr('') + _xStr('') + _xStr('') + _xStr('') + _xStr('') + _xStr('');
  return _xStr(v.status) + _xStr(v.observed_differences) + _xStr(v.roadblock) + _xStr(v.main_action) +
    _xStr(citedTicketsList(v).join('\n')) + _xStr(v.suggested_description_update);
}

function sheetXML(sheetName, rows) {
  const descByName = new Map(rubric.map(r => [r.name, r.description]));
  const hdrs = ['Category', 'Workflow', 'Description', 'Tickets', 'Total Hours', 'AHT (min)', 'FRR', 'Touches/Ticket', 'Complexity',
    'Reconciliation Status', 'Observed Differences', 'Roadblock', 'Main Action', 'Cited Tickets', 'Suggested Update'];
  const headerRow = '<Row>' + hdrs.map(_xHdr).join('') + '</Row>';
  const dataRows = rows.map(r => {
    const cells = [
      _xStr(r.category), _xStr(r.workflow), _xStr(descByName.get(r.workflow) || ''),
      _xNum(r.tickets), `<Cell><Data ss:Type="Number">${r.hours.toFixed(1)}</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.aht.toFixed(1)}</Data></Cell>`,
      `<Cell><Data ss:Type="String">${(r.frr * 100).toFixed(0)}%</Data></Cell>`,
      `<Cell><Data ss:Type="Number">${r.touches.toFixed(2)}</Data></Cell>`,
      _xStr(r.tier),
      reconSheetCells(r.workflow),
    ].join('');
    return `<Row>${cells}</Row>`;
  }).join('');
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${headerRow}${dataRows}</Table></Worksheet>`;
}

// Non-Billable Pivot: Company (Label) × Workflow matrix of non-billable hours, with totals.
function nonBillablePivotSheetXML(sheetName) {
  const inScope = ticketRecords.filter(r => r.nonbillable_hours > 0);
  const wfSet = new Set(), byCompany = new Map();
  for (const r of inScope) {
    const wf = r.workflow || '(not classified)';
    wfSet.add(wf);
    if (!byCompany.has(r.company)) byCompany.set(r.company, new Map());
    const m = byCompany.get(r.company);
    m.set(wf, (m.get(wf) || 0) + r.nonbillable_hours);
  }
  const workflows = [...wfSet].sort();
  const companies = [...byCompany.keys()].sort();
  const header = '<Row>' + _xHdr('Company (Label)') + workflows.map(_xHdr).join('') + _xHdr('Total') + '</Row>';
  const body = companies.map(c => {
    const m = byCompany.get(c);
    let rowTotal = 0;
    const cells = workflows.map(w => { const v = m.get(w) || 0; rowTotal += v; return `<Cell><Data ss:Type="Number">${v.toFixed(2)}</Data></Cell>`; }).join('');
    return `<Row>${_xStr(c)}${cells}<Cell><Data ss:Type="Number">${rowTotal.toFixed(2)}</Data></Cell></Row>`;
  }).join('');
  const colTotals = workflows.map(w => { let t = 0; for (const m of byCompany.values()) t += (m.get(w) || 0); return t; });
  const grand = colTotals.reduce((a, b) => a + b, 0);
  const totalRow = `<Row>${_xHdr('Total')}${colTotals.map(t => `<Cell><Data ss:Type="Number">${t.toFixed(2)}</Data></Cell>`).join('')}<Cell><Data ss:Type="Number">${grand.toFixed(2)}</Data></Cell></Row>`;
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${header}${body}${totalRow}</Table></Worksheet>`;
}

// Non-Billable Detail: one flat row per in-scope ticket.
function nonBillableDetailSheetXML(sheetName) {
  const inScope = ticketRecords.filter(r => r.nonbillable_hours > 0);
  const hdrs = ['Company (Label)', 'Workflow', 'Ticket', 'Total Hours', 'Non-Billable Hours', 'NB Flag', 'Contract Types', 'Roles'];
  const header = '<Row>' + hdrs.map(_xHdr).join('') + '</Row>';
  const body = inScope.map(r =>
    `<Row>${_xStr(r.company)}${_xStr(r.workflow || '(not classified)')}${_xStr(r.ticket_id)}` +
    `<Cell><Data ss:Type="Number">${(r.hours || 0).toFixed(2)}</Data></Cell>` +
    `<Cell><Data ss:Type="Number">${(r.nonbillable_hours || 0).toFixed(2)}</Data></Cell>` +
    `${_xStr(r.nonbillable_flag)}${_xStr((r.contract_types || []).join(', '))}${_xStr((r.roles || []).join(', '))}</Row>`
  ).join('');
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${header}${body}</Table></Worksheet>`;
}

// Ticket Detail sheet: one flat row per ticket (unfiltered) — ticket id, company,
// workflow, hours, touches. Workflow falls back to "(not classified)" pre-classification.
function ticketDetailSheetXML(sheetName) {
  const hdrs = ['Ticket ID', 'Company (Label)', 'Workflow', 'Total Hours', 'Touches'];
  const header = '<Row>' + hdrs.map(_xHdr).join('') + '</Row>';
  const body = ticketRecords.map(r =>
    `<Row>${_xStr(r.ticket_id)}${_xStr(r.company)}${_xStr(r.workflow || '(not classified)')}` +
    `<Cell><Data ss:Type="Number">${(r.hours || 0).toFixed(2)}</Data></Cell>` +
    `<Cell><Data ss:Type="Number">${r.touches || 0}</Data></Cell></Row>`
  ).join('');
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${header}${body}</Table></Worksheet>`;
}

// Playbook sheet: one row per Company x Workflow, mirrors the Playbook view columns.
function playbookSheetXML(sheetName) {
  const hdrs = ['Company (Label)', 'Workflow', 'Description', 'Common Steps', 'Responsible Tier', 'Tickets', 'Total Hours', 'FRR', 'Touches/Ticket', 'AHT (min)'];
  const header = '<Row>' + hdrs.map(_xHdr).join('') + '</Row>';
  const body = playbookRows.map(r => {
    const stepsText = (r.steps && r.steps.length) ? r.steps.map((s, i) => `${i + 1}. ${s}`).join('\n') : (r.stepsStatus || 'Insufficient evidence');
    return `<Row>${_xStr(r.company)}${_xStr(r.workflow)}${_xStr(r.description || '')}${_xStr(stepsText)}${_xStr(r.tier)}` +
      `<Cell><Data ss:Type="Number">${r.tickets}</Data></Cell>` +
      `<Cell><Data ss:Type="Number">${r.hours.toFixed(1)}</Data></Cell>` +
      `<Cell><Data ss:Type="String">${(r.frr * 100).toFixed(0)}%</Data></Cell>` +
      `<Cell><Data ss:Type="Number">${r.touches.toFixed(2)}</Data></Cell>` +
      `<Cell><Data ss:Type="Number">${r.aht.toFixed(1)}</Data></Cell></Row>`;
  }).join('');
  return `<Worksheet ss:Name="${xmlEscape(sanitizeSheetName(sheetName))}"><Table>${header}${body}</Table></Worksheet>`;
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
  let sheets = sheetXML(uniqueName('Workflow Dashboard'), rowsForCompany(null));
  sheets += ticketDetailSheetXML(uniqueName('Ticket Detail'));
  if (ticketRecords.some(r => r.nonbillable_hours > 0)) {
    sheets += nonBillablePivotSheetXML(uniqueName('Non-Billable Pivot'));
    sheets += nonBillableDetailSheetXML(uniqueName('Non-Billable Detail'));
  }
  if (playbookRows.length) sheets += playbookSheetXML(uniqueName('Playbook'));
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

// "sample_time_entries.csv" -> "sample_time_entries_" (sanitized for a filename).
function datasetFilePrefix() {
  if (!datasetName) return '';
  const base = datasetName.replace(/\.[^.]+$/, '').replace(/[^0-9A-Za-z._-]+/g, '_').replace(/^_+|_+$/g, '');
  return base ? base + '_' : '';
}

document.getElementById('exportCsvBtn').addEventListener('click', async () => {
  const xml = buildWorkbookXML();
  const fileLabel = `${datasetFilePrefix()}workflows_by_company.xls`;
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

  if (downloadAttempted) { showToast(`Workbook download started (${fileLabel}) — Workflow Dashboard (+ reconciliation), Ticket Detail, Non-Billable Pivot & Detail${playbookRows.length ? ', Playbook' : ''}, and a tab per company. If Excel warns about the file format, choose "Yes, open it".`); return; }

  const flat = buildTSV();
  try { await navigator.clipboard.writeText(flat); showToast('Download blocked here — copied a flattened summary to clipboard instead'); return; } catch (e) {}
  const modalTa = document.getElementById('modalTextarea');
  modalTa.value = flat;
  document.getElementById('modalBackdrop').classList.add('show');
  modalTa.focus(); modalTa.select();
});
