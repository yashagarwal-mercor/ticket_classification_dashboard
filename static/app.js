const MODEL = "claude-haiku-4-5";
const MODEL_SUMMARY = "claude-sonnet-5"; // one-shot quality tasks (theme summary, reconciliation)
const BATCH_SIZE = 12;      // notes bundles are long (full Dutch email threads) — smaller batches than the old 30
const CONCURRENCY = 5;
const CATCHALL_WORKFLOW = "Unclassified / Other";
const COVERAGE_GAP_THRESHOLD = 0.10; // >10% unclassified triggers a rubric coverage-gap warning (§2.3)
const THEME_SAMPLE_SIZE = 40;        // # of unclassified notes sampled for theme summarization
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
let coverageInfo = null;   // {total, unclassified, pct, noNotes, themes:[{theme,description,examples}]} — §2.3 coverage gap
let activeBatchId = null;  // set while a Batch API run is being polled; enables the Cancel-batch button
let reconciliation = null;      // Map(workflowName -> verdict) from Feature A reconciliation (Phase 5)
let citedEvidenceById = null;   // Map(uid -> {en, nl}) evidence for cited tickets (rendered + persisted)

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
    hint.textContent = `A classification batch (${pending.batch_id}) is in progress — Group the same file to resume watching it.`;
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

function buildExtractionSystem(wf) {
  return `You are analyzing helpdesk/support tickets that were all classified as the workflow "${wf.name}".
Workflow description (the rubric's stated claim about this process):
"${wf.description}"

Each ticket comes with its engineer time-entry notes (mostly Dutch, sometimes English). Reason over the Dutch directly — do NOT translate first. For each ticket, extract this fixed schema:
- main_action: a short English verb phrase for the dominant manual action performed (e.g. "reset password in Entra", "assigned AD group permissions", "escalated to Engineering").
- bottleneck: the main roadblock/wait — one of: none, customer wait, vendor/third-party, internal handoff, approval, technical blocker.
- bottleneck_detail: a short English phrase with specifics (empty string when bottleneck is "none").
- customer_contact: true if the engineer communicated with the customer / end-user on this ticket.
- customer_contact_channel: one of none, email, phone, remote session, on-site, unknown.
- teams_involved: the roles/teams that touched the ticket, including escalations named in the notes (e.g. "doorgezet naar Engineering" -> "Engineering"). Use short English labels.
- rubric_gap: anything the ticket actually involved that the workflow description above does NOT mention; empty string if the ticket fits the description.
- evidence_nl: the single most telling short quote from the notes, verbatim in its original language.
- evidence_en: the English translation of that quote.

Output ALL fields in English EXCEPT evidence_nl (keep it verbatim). Echo ticket_id EXACTLY as given. Return exactly one object per input ticket.`;
}

function buildExtractionUser(chunk) {
  return "Tickets:\n\n" + chunk.map(t =>
    `ticket_id: ${t.uid}\nnotes: ${(t.notes || '').replace(/\s+/g, ' ').slice(0, 1500)}`
  ).join('\n\n');
}

// Extract one chunk of same-workflow tickets. chunk items: {uid, notes}. wf: {name, description}.
// Returns Map(uid -> extraction), dropping any echoed id that wasn't sent (anti-hallucination).
async function extractBatch(chunk, wf) {
  const resp = await fetch(`${API_BASE}/api/classify`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, max_tokens: 4096, system: buildExtractionSystem(wf), user: buildExtractionUser(chunk), schema: EXTRACTION_SCHEMA }),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const body = await resp.json();
  if (!body.ok) throw new Error(body.error || 'Extraction failed');
  const parsed = JSON.parse(body.text);
  const sentIds = new Set(chunk.map(t => t.uid));
  const result = new Map();
  for (const e of parsed.extractions || []) {
    if (e && typeof e.ticket_id === 'string' && sentIds.has(e.ticket_id)) result.set(e.ticket_id, e);
  }
  return result;
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

  // Stage 1 — extraction (per-workflow chunks, worker pool). exById lets us surface
  // per-ticket evidence (English + original Dutch) for the tickets the reconcile cites.
  const tasks = [];
  for (const [wfName, tix] of byWorkflow) {
    const wf = { name: wfName, description: descByName.get(wfName) || '' };
    for (let i = 0; i < tix.length; i += EXTRACT_BATCH_SIZE) tasks.push({ chunk: tix.slice(i, i + EXTRACT_BATCH_SIZE), wf });
  }
  const extractionsByWf = new Map();
  const exById = new Map();
  let ti = 0, exDone = 0;
  const upEx = () => setStatus(`Extracting ticket details: ${exDone}/${tasks.length} batches...`);
  upEx();
  async function exWorker() {
    while (ti < tasks.length) {
      if (cancelRequested) return;
      const t = tasks[ti++];
      try {
        const m = await extractBatch(t.chunk, t.wf);
        if (!extractionsByWf.has(t.wf.name)) extractionsByWf.set(t.wf.name, []);
        for (const [uid, e] of m) { extractionsByWf.get(t.wf.name).push(e); exById.set(uid, e); }
      } catch (err) { console.error('extract batch failed', err); }
      exDone++; upEx();
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, exWorker));
  if (cancelRequested) { setStatus('Reconciliation cancelled.'); btn.disabled = false; return; }

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
// the concatenated notes (mostly Dutch) prefixed with label/contract tags.
function buildTicketsForClassification() {
  return ticketRecords.map(rec => {
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

function saveEnrichedState() {
  try {
    const names = [];
    const idx = new Map();
    const wf = dashboardRows.map(r => {
      if (!idx.has(r.workflow)) { idx.set(r.workflow, names.length); names.push(r.workflow); }
      return idx.get(r.workflow);
    });
    localStorage.setItem(ENRICHED_LS_KEY, JSON.stringify({
      fingerprint: datasetFingerprint(), names, wf, rubric, coverageInfo,
      reconciliation: reconciliation ? [...reconciliation.entries()] : null,
      citedEvidence: citedEvidenceById ? [...citedEvidenceById.entries()] : null,
    }));
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
  if (!Array.isArray(st.names) || !Array.isArray(st.wf) || st.wf.length !== ticketRecords.length) return false;

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
  applyClassifications(tickets, classifications);
  coverageInfo = (st.coverageInfo && typeof st.coverageInfo === 'object') ? st.coverageInfo : null;
  reconciliation = Array.isArray(st.reconciliation) ? new Map(st.reconciliation) : null;
  citedEvidenceById = Array.isArray(st.citedEvidence) ? new Map(st.citedEvidence) : null;
  showError('');
  showDashboard();
  return true;
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
function saveBatchState(batchId) {
  try { localStorage.setItem(BATCH_LS_KEY, JSON.stringify({ batch_id: batchId, fingerprint: datasetFingerprint() })); } catch (e) {}
}
function loadBatchState() { try { return JSON.parse(localStorage.getItem(BATCH_LS_KEY) || 'null'); } catch (e) { return null; } }
function clearBatchState() { try { localStorage.removeItem(BATCH_LS_KEY); } catch (e) {} }

async function runBatchClassification() {
  cancelRequested = false;
  showError('');
  if (!rubric.length) { showError('Add at least one workflow to the rubric.'); return; }
  if (!parsedRows.length) { showError('Upload a time-entry export first.'); return; }
  if (!ticketRecords.length && !runGrouping()) return;

  const tickets = buildTicketsForClassification();
  const chunks = chunkTickets(tickets);
  const requests = chunks.map((chunk, i) => ({ custom_id: `c${i}`, user: buildUserPrompt(chunk) }));

  document.getElementById('classifyBtn').disabled = true;
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('progressBarInner').style.width = '0%';
  document.getElementById('progressText').textContent =
    `Submitting ${requests.length.toLocaleString()} requests to the Batch API...`;

  let batchId;
  try {
    const resp = await fetch(`${API_BASE}/api/batch/create`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: MODEL, max_tokens: 4096, system: buildSystemPrompt(), schema: RESPONSE_SCHEMA, requests }),
    });
    const body = await resp.json();
    if (!body.ok) throw new Error(body.error || 'Batch create failed');
    batchId = body.batch_id;
  } catch (err) {
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('classifyBtn').disabled = false;
    showError('Could not start the batch: ' + describeError(err));
    return;
  }
  saveBatchState(batchId);
  await pollBatchToCompletion(batchId, tickets, chunks);
}

// Poll a batch to completion, then map results back by custom_id and finalize.
// Reused by the resume path after a page reload.
async function pollBatchToCompletion(batchId, tickets, chunks) {
  const validNames = new Set(rubric.map(r => r.name));
  const total = chunks.length;
  const bar = document.getElementById('progressBarInner');
  const text = document.getElementById('progressText');
  const cancelBtn = document.getElementById('cancelBtn');
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('classifyBtn').disabled = true;
  activeBatchId = batchId;              // enables the Cancel-batch handler
  cancelBtn.textContent = 'Cancel batch';
  const endPoll = () => {
    activeBatchId = null;
    cancelBtn.textContent = 'Cancel';
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('classifyBtn').disabled = false;
  };

  while (true) {
    if (cancelRequested) { endPoll(); return; } // the Cancel-batch handler owns the cancel call + message
    let status;
    try {
      const resp = await fetch(`${API_BASE}/api/batch/status?id=${encodeURIComponent(batchId)}`);
      status = await resp.json();
    } catch (err) { status = { ok: false, error: describeError(err) }; }

    if (status.ok) {
      const c = status.request_counts || {};
      const done = (c.succeeded || 0) + (c.errored || 0) + (c.canceled || 0) + (c.expired || 0);
      const pct = total > 0 ? Math.round(done / total * 100) : 0;
      bar.style.width = pct + '%';
      text.textContent = `Batch classifying: ${done.toLocaleString()} / ${total.toLocaleString()} requests (${pct}%) — ${status.processing_status}. Safe to close the tab.`;
      if (status.processing_status === 'ended') break;
      if (['canceled', 'canceling', 'expired'].includes(status.processing_status)) {
        endPoll(); clearBatchState();
        showError(`Batch ${status.processing_status}.`);
        return;
      }
    } else {
      text.textContent = `Batch status check failed (${status.error || 'unknown'}) — retrying...`;
    }
    await sleep(5000);
  }

  text.textContent = 'Batch complete — downloading results...';
  let body;
  try {
    const resp = await fetch(`${API_BASE}/api/batch/results?id=${encodeURIComponent(batchId)}`);
    body = await resp.json();
    if (!body.ok) throw new Error(body.error || 'Batch results failed');
  } catch (err) {
    endPoll();
    showError('Could not download batch results: ' + describeError(err));
    return;
  }
  if (body.usage) console.log('Batch cache usage:', body.usage);

  const classifications = new Array(tickets.length).fill(null);
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

  clearBatchState();
  activeBatchId = null;                       // batch done — disable the cancel path
  cancelBtn.textContent = 'Cancel';
  await finalizeClassification(tickets, classifications); // keeps progressWrap visible for the theme step
  document.getElementById('progressWrap').style.display = 'none';
  document.getElementById('classifyBtn').disabled = false;
}

// If a batch is pending for the currently-grouped dataset, resume watching it.
function maybeResumeBatch() {
  const st = loadBatchState();
  if (!st || st.fingerprint !== datasetFingerprint()) return false;
  showError('');
  const tickets = buildTicketsForClassification();
  pollBatchToCompletion(st.batch_id, tickets, chunkTickets(tickets));
  return true;
}

document.getElementById('classifyBtn').addEventListener('click', () => {
  if (document.getElementById('useBatch').checked) runBatchClassification();
  else runClassification();
});
document.getElementById('cancelBtn').addEventListener('click', async () => {
  cancelRequested = true; // stops the in-browser worker loop and the batch poll loop
  const id = activeBatchId;
  if (!id) return; // in-browser mode: runClassification shows "Classification cancelled."
  // Batch mode: actually cancel the batch on Anthropic (not just stop watching).
  activeBatchId = null;
  document.getElementById('cancelBtn').textContent = 'Cancel';
  document.getElementById('progressText').textContent = 'Cancelling batch...';
  try {
    const resp = await fetch(`${API_BASE}/api/batch/cancel`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ batch_id: id }),
    });
    const body = await resp.json();
    if (!body.ok) throw new Error(body.error || 'cancel failed');
    showError('Batch cancelled.');
  } catch (err) {
    showError(`Could not cancel the batch (id ${id}): ${describeError(err)}`);
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
