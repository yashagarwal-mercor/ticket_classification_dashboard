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
                          (defaults to claude-haiku-4-5 for bulk classification).

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

import json
import os
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

import anthropic

PORT = 8787
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

if not os.environ.get("ANTHROPIC_API_KEY"):
    raise SystemExit("Set ANTHROPIC_API_KEY before running this script.")

client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env

INFER_MODEL = "claude-opus-4-8"

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
        if self.path == "/api/health":
            self._send_json(200, {"ok": True})
            return
        super().do_GET()

    def do_POST(self):
        if self.path == "/api/infer":
            self._handle_infer()
        elif self.path == "/api/classify":
            self._handle_classify()
        else:
            self._send_json(404, {"ok": False, "error": "not found"})

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length))

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
                system=body["system"],
                messages=[{"role": "user", "content": body["user"]}],
                output_config={"format": {"type": "json_schema", "schema": body["schema"]}},
            )
            text_block = next(b for b in response.content if b.type == "text")
            self._send_json(200, {"ok": True, "text": text_block.text})
        except Exception as e:
            self._send_json(200, {"ok": False, "error": str(e)})

    def log_message(self, format, *args):
        pass  # keep the console quiet -- errors already surface in the browser


if __name__ == "__main__":
    print(f"Serving on http://localhost:{PORT} -- open that URL in your browser.")
    ThreadingHTTPServer(("localhost", PORT), Handler).serve_forever()
