"""
Local backend for the ticket workflow classifier.

Serves the static frontend (static/) and two Anthropic-backed endpoints:

  POST /api/infer     -- given column headers + a sample of rows, infer the
                          column mapping (which column is company, hours,
                          touches, ticket-count, etc). Uses claude-opus-4-8
                          (quality matters more than cost here -- it's one
                          call, not thousands). The workflow rubric is NOT
                          generated here -- the user supplies it themselves.
  POST /api/classify   -- classify one batch of rows against the rubric.
                          Uses whatever model the frontend requests
                          (defaults to claude-sonnet-5 for bulk classification).

Runs entirely server-side because some Anthropic organizations (ones with
custom data-retention settings) reject direct browser-to-API calls at the
CORS layer -- see the comment in classify_relay.py's predecessor. Routing
every Claude call through this process sidesteps that regardless of org
policy, and also means your API key never has to touch the browser.

Run:
    pip3 install anthropic
    export ANTHROPIC_API_KEY="sk-ant-..."
    python3 server.py
Then open http://localhost:8787
"""

import csv
import json
import os
import re
import tempfile
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import anthropic
import openpyxl
from anthropic.types.message_create_params import MessageCreateParamsNonStreaming
from anthropic.types.messages.batch_create_params import Request

PORT = 8787
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

if not os.environ.get("ANTHROPIC_API_KEY"):
    raise SystemExit("Set ANTHROPIC_API_KEY before running this script.")

client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env

INFER_MODEL = "claude-opus-4-8"

# Ticket descriptions can run to ~14k chars; cap them so the projected lookup we return
# to the browser stays bounded (mirrors META_DESC_CAP in app.js).
META_DESC_CAP = 4000

# Anthropic batch ids look like "msgbatch_...". Require that exact prefix before
# passing a client-supplied id to the SDK: guards against parameter injection AND
# rejects stale/garbage ids locally (with a clear error) instead of proxying a
# doomed request to Anthropic, which is what let a bad id spin the poll loop.
BATCH_ID_RE = re.compile(r"^msgbatch_[A-Za-z0-9_-]{1,119}$")

_MAPPING_FIELDS = [
    "instance_column", "ticketnumber_column", "startdate_column",
    "summarynotes_column", "internalnotes_column", "kpi_label_column",
    "hours_worked_column", "hours_to_bill_column", "nonbillable_column",
    "roles_column", "billingcodes_column", "contracttype_column",
]

INFER_SCHEMA = {
    "type": "object",
    "properties": {
        "column_mapping": {
            "type": "object",
            "properties": {f: {"type": ["string", "null"]} for f in _MAPPING_FIELDS},
            "required": _MAPPING_FIELDS,
            "additionalProperties": False,
        },
    },
    "required": ["column_mapping"],
    "additionalProperties": False,
}

INFER_SYSTEM_PROMPT = """You are setting up an IT/MSP helpdesk analysis tool from a raw TIME-ENTRY export (one row per time entry logged against a ticket; many rows roll up to one ticket). You'll be given the column headers and a sample of rows. Map each field to the single best-matching column header, or null if truly absent. Return the header string EXACTLY as given.

   - instance_column: the customer/company/account/instance this entry belongs to (e.g. "datavisual", "hupra"). Often named all_time_entries[instance].
   - ticketnumber_column: the ticket identifier that groups entries into one ticket (e.g. all_time_entries[ticketnumber], a value like "T20260112.0024"). Prefer the human ticket number over an internal row id.
   - startdate_column: the start date/time of the entry, used to order notes chronologically (e.g. all_time_entries[startdatetime]).
   - summarynotes_column: the primary free-text engineer note (mostly Dutch). Often all_time_entries[summarynotes].
   - internalnotes_column: a secondary/internal free-text note, if present (often all_time_entries[internalnotes]); null if absent.
   - kpi_label_column: a level-1 label/KPI category (e.g. all_time_entries[KPI_detail_level01], values like "Klant tickets" / "Interne uren"); null if absent.
   - hours_worked_column: hours worked per entry (e.g. [Sumhoursworked]).
   - hours_to_bill_column: billable hours per entry (e.g. [Sumhourstobill]); null if absent.
   - nonbillable_column: a TRUE/FALSE (or yes/no) flag marking the entry non-billable (e.g. all_time_entries[isnonbillable]); null if absent.
   - roles_column: the engineer role/team (e.g. all_roles[name], values like "Eerstelijns Service Engineer"); null if absent.
   - billingcodes_column: the billing code (e.g. all_billingcodes[name], values like "Support remote"); null if absent.
   - contracttype_column: the contract type (e.g. all_time_entries[contracttype], values like "Strippenkaart", "Servicecontract"); null if absent.
   Map instance_column, ticketnumber_column, startdate_column, summarynotes_column, and hours_worked_column whenever any plausible column exists -- these are required downstream."""


def _cached_system(system_text):
    """Wrap a system prompt string as a cache_control'd text block so the
    repeated rubric prefix is served from cache (~0.1x) across requests.
    Caching only kicks in above the model's minimum prefix (4096 tokens on
    Haiku) -- the frontend pads the rubric prompt to clear that floor."""
    return [{"type": "text", "text": system_text, "cache_control": {"type": "ephemeral"}}]


# --- Tickets-export parsing (streamed server-side to keep the 83MB/315MB-XML file
# off the browser, which OOMs trying to materialize the whole workbook) ---

def _s(v):
    """Coerce any cell value to a stripped string ('' for blank)."""
    return "" if v is None else str(v).strip()


def _valid_cols(cols):
    """The candidate-header map must be {field: [header, ...]} of strings."""
    if not isinstance(cols, dict):
        return False
    for k, v in cols.items():
        if not isinstance(k, str) or not isinstance(v, list):
            return False
        if not all(isinstance(x, str) for x in v):
            return False
    return True


def _open_xlsx(path):
    """Return (headers, row_iterator, close_fn) for an xlsx, read-only/streamed.
    We pass a file object (not the path) so openpyxl doesn't reject the temp file's
    non-.xlsx suffix — it validates by extension only when given a path string."""
    fh = open(path, "rb")
    wb = openpyxl.load_workbook(fh, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    try:
        headers = [_s(x) for x in next(it)]
    except StopIteration:
        headers = []

    def close():
        wb.close()
        fh.close()

    return headers, it, close


def _open_csv(path):
    """Return (headers, row_iterator, close_fn) for a CSV, streamed line by line."""
    fh = open(path, newline="", encoding="utf-8-sig")
    reader = csv.reader(fh)
    try:
        headers = [_s(x) for x in next(reader)]
    except StopIteration:
        headers = []
    return headers, reader, fh.close


def _project_tickets(headers, rows_iter, cols):
    """Stream rows, keeping only the candidate columns, into a compact
    "<instance>/<ticketnumber>" -> {header: value} lookup. Returns {} if the
    compound-key columns aren't present (nothing to join on)."""
    index = {h: i for i, h in enumerate(headers)}
    present = {}  # field -> the first candidate header that exists in the file
    for field, cands in cols.items():
        for h in cands:
            if h in index:
                present[field] = h
                break
    inst_h, tn_h = present.get("instance"), present.get("ticketnumber")
    if not inst_h or not tn_h:
        return {"headers": headers, "byKey": {}, "count": 0}
    inst_i, tn_i = index[inst_h], index[tn_h]
    desc_h = present.get("description")
    store = [(h, index[h]) for h in present.values()]  # (header, column index)
    by_key = {}
    for row in rows_iter:
        n = len(row)
        inst = _s(row[inst_i]) if inst_i < n else ""
        tn = _s(row[tn_i]) if tn_i < n else ""
        if not inst and not tn:
            continue
        rec = {}
        for h, ci in store:
            v = _s(row[ci]) if ci < n else ""
            if h == desc_h and len(v) > META_DESC_CAP:
                v = v[:META_DESC_CAP]
            rec[h] = v
        by_key[inst + "/" + tn] = rec
    return {"headers": headers, "byKey": by_key, "count": len(by_key)}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send_json(self, status, obj):
        payload = json.dumps(obj).encode()
        self.send_response(status)
        self._cors_headers()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self._send_json(200, {"ok": True})
            return
        if parsed.path in ("/api/batch/status", "/api/batch/results"):
            batch_id = (parse_qs(parsed.query).get("id") or [""])[0]
            if not BATCH_ID_RE.match(batch_id):
                self._send_json(200, {"ok": False, "error": "invalid batch id"})
                return
            if parsed.path == "/api/batch/status":
                self._handle_batch_status(batch_id)
            else:
                self._handle_batch_results(batch_id)
            return
        super().do_GET()

    def do_POST(self):
        if self.path == "/api/infer":
            self._handle_infer()
        elif self.path == "/api/classify":
            self._handle_classify()
        elif self.path == "/api/batch/create":
            self._handle_batch_create()
        elif self.path == "/api/batch/cancel":
            self._handle_batch_cancel()
        elif urlparse(self.path).path == "/api/tickets/parse":
            self._handle_tickets_parse()
        else:
            self._send_json(404, {"ok": False, "error": "not found"})

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length))

    def _handle_tickets_parse(self):
        """Parse an uploaded Tickets export server-side and return a compact
        projected lookup. The raw bytes are streamed to a temp file (never held
        twice in memory); openpyxl read_only streams the sheet at low memory."""
        tmp_path = None
        try:
            parsed = urlparse(self.path)
            raw_cols = (parse_qs(parsed.query).get("cols") or ["{}"])[0]
            cols = json.loads(raw_cols)
            if not _valid_cols(cols):
                self._send_json(200, {"ok": False, "error": "invalid cols parameter"})
                return
            length = int(self.headers.get("Content-Length", 0))
            if length <= 0:
                self._send_json(200, {"ok": False, "error": "empty upload"})
                return
            with tempfile.NamedTemporaryFile(delete=False, suffix=".upload") as tf:
                tmp_path = tf.name
                remaining = length
                while remaining > 0:
                    chunk = self.rfile.read(min(remaining, 1 << 20))
                    if not chunk:
                        break
                    tf.write(chunk)
                    remaining -= len(chunk)
            with open(tmp_path, "rb") as fh:
                magic = fh.read(2)
            headers, rows_iter, close_fn = (
                _open_xlsx(tmp_path) if magic == b"PK" else _open_csv(tmp_path)
            )
            try:
                result = _project_tickets(headers, rows_iter, cols)
            finally:
                close_fn()
            self._send_json(200, {"ok": True, **result})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    def _handle_infer(self):
        try:
            body = self._read_body()
            headers = body["headers"]
            sample_rows = body["sampleRows"]
            user_content = (
                f"Columns: {json.dumps(headers)}\n\n"
                f"Sample rows ({len(sample_rows)} of them):\n{json.dumps(sample_rows, indent=1)}"
            )
            response = client.messages.create(
                model=INFER_MODEL,
                max_tokens=8192,
                system=INFER_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_content}],
                output_config={"format": {"type": "json_schema", "schema": INFER_SCHEMA}},
            )
            text_block = next(b for b in response.content if b.type == "text")
            self._send_json(200, {"ok": True, "text": text_block.text})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def _handle_classify(self):
        try:
            body = self._read_body()
            response = client.messages.create(
                model=body["model"],
                max_tokens=body.get("max_tokens", 4096),
                system=_cached_system(body["system"]),  # cache_control on the repeated rubric prefix
                messages=[{"role": "user", "content": body["user"]}],
                output_config={"format": {"type": "json_schema", "schema": body["schema"]}},
            )
            text_block = next(b for b in response.content if b.type == "text")
            self._send_json(200, {"ok": True, "text": text_block.text})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    # --- Batch API relay (proxies the Anthropic Message Batches API by id) ---

    def _handle_batch_create(self):
        try:
            body = self._read_body()
            model = body["model"]
            max_tokens = body.get("max_tokens", 4096)
            system = _cached_system(body["system"])  # shared, cached prefix across all requests
            schema = body["schema"]
            requests = [
                Request(
                    custom_id=r["custom_id"],
                    params=MessageCreateParamsNonStreaming(
                        model=model,
                        max_tokens=max_tokens,
                        system=system,
                        messages=[{"role": "user", "content": r["user"]}],
                        output_config={"format": {"type": "json_schema", "schema": schema}},
                    ),
                )
                for r in body["requests"]
            ]
            batch = client.messages.batches.create(requests=requests)
            self._send_json(200, {"ok": True, "batch_id": batch.id, "processing_status": batch.processing_status})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def _handle_batch_cancel(self):
        try:
            body = self._read_body()
            batch_id = body.get("batch_id", "")
            if not BATCH_ID_RE.match(batch_id):
                self._send_json(200, {"ok": False, "error": "invalid batch id"})
                return
            b = client.messages.batches.cancel(batch_id)
            self._send_json(200, {"ok": True, "processing_status": b.processing_status})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def _handle_batch_status(self, batch_id):
        try:
            b = client.messages.batches.retrieve(batch_id)
            c = b.request_counts
            self._send_json(200, {
                "ok": True,
                "processing_status": b.processing_status,
                "request_counts": {
                    "processing": c.processing, "succeeded": c.succeeded,
                    "errored": c.errored, "canceled": c.canceled, "expired": c.expired,
                },
            })
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def _handle_batch_results(self, batch_id):
        try:
            b = client.messages.batches.retrieve(batch_id)
            if b.processing_status != "ended":
                self._send_json(200, {"ok": False, "error": "batch not ended", "processing_status": b.processing_status})
                return
            results, cache_read, cache_creation = [], 0, 0
            for r in client.messages.batches.results(batch_id):
                rtype = r.result.type
                if rtype == "succeeded":
                    msg = r.result.message
                    text = next((blk.text for blk in msg.content if blk.type == "text"), "")
                    cache_read += getattr(msg.usage, "cache_read_input_tokens", 0) or 0
                    cache_creation += getattr(msg.usage, "cache_creation_input_tokens", 0) or 0
                    results.append({"custom_id": r.custom_id, "ok": True, "text": text})
                else:
                    err = rtype
                    try:
                        err = r.result.error.type
                    except Exception:
                        pass
                    results.append({"custom_id": r.custom_id, "ok": False, "error": str(err)})
            self._send_json(200, {
                "ok": True,
                "results": results,
                "usage": {"cache_read_input_tokens": cache_read, "cache_creation_input_tokens": cache_creation},
            })
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def log_message(self, format, *args):
        pass  # keep the console quiet -- errors already surface in the browser


if __name__ == "__main__":
    print(f"Serving on http://localhost:{PORT} -- open that URL in your browser.")
    ThreadingHTTPServer(("localhost", PORT), Handler).serve_forever()
