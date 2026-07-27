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

INFER_SCHEMA = {
    "type": "object",
    "properties": {
        "column_mapping": {
            "type": "object",
            "properties": {
                "company_column": {"type": ["string", "null"]},
                "text_columns": {"type": "array", "items": {"type": "string"}},
                "hours_column": {"type": ["string", "null"]},
                "touches_column": {"type": ["string", "null"]},
                "first_resolution_column": {"type": ["string", "null"]},
                "ticket_count_column": {"type": ["string", "null"]},
            },
            "required": [
                "company_column", "text_columns", "hours_column",
                "touches_column", "first_resolution_column", "ticket_count_column",
            ],
            "additionalProperties": False,
        },
    },
    "required": ["column_mapping"],
    "additionalProperties": False,
}

INFER_SYSTEM_PROMPT = """You are setting up an IT/MSP helpdesk ticket classification tool from a raw spreadsheet export. You'll be given the column headers and a sample of rows.

Identify the COLUMN MAPPING:
   - company_column: a column identifying the customer/company/account this row belongs to, if one exists (else null).
   - text_columns: the column(s) that carry the actual classification signal for what the ticket is about. If there's a single free-text description/title/summary column, use just that one. If there is NO free-text column but there are structured tag columns instead (e.g. IssueType + SubIssueType, or Category + Subcategory), list ALL of them in the order they should be combined -- these will be concatenated per row to form the classification input.
   - hours_column: total hours spent, if present.
   - touches_column: number of touches/interactions/time-entries, if present.
   - first_resolution_column: a column indicating first-touch resolution, if present -- this might be a boolean/yes-no flag (one row = one ticket) OR a count (one row represents many tickets, e.g. "CountFirstResolution" alongside a ticket-count column). Either shape is fine, just identify the column.
   - ticket_count_column: if this data is PRE-AGGREGATED (i.e. each row represents a group of N tickets sharing the same tags/company/period, rather than one row = one ticket), identify the column holding that count (e.g. "TicketCount"). If each row is one individual ticket, set this to null.
   Only map hours_column/touches_column to null if truly absent -- these are usually present under some name (SumHours, TotalHours, Hours, Duration, SumTouches, Touches, InteractionCount, etc)."""


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
