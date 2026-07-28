# Ticket Workflow Classifier & Dashboard

Upload a spreadsheet of helpdesk/SRE tickets, define the workflows you want to
classify them into, and get a dashboard of ticket volume, hours, first-touch
resolution rate, and complexity tier per workflow — filterable by company and
exportable to Excel.

Classification is done by an LLM (Claude), called once per batch of tickets.
Column mapping (which column is "hours", which is "company", etc.) can
optionally be inferred by AI as well, but the workflow rubric itself is
always written by you — the tool does not invent workflows on your behalf.

## Setup

1. Create and activate a virtual environment:

   ```
   python3 -m venv .venv
   source .venv/bin/activate
   ```

   (On Windows: `.venv\Scripts\activate`)

2. Install the Anthropic Python SDK into the virtual environment:

   ```
   pip install anthropic
   ```

3. Export your Anthropic API key:

   ```
   export ANTHROPIC_API_KEY="sk-ant-..."
   ```

4. Run the server (with the virtual environment still active):

   ```
   python3 server.py
   ```

5. Open [http://localhost:8787](http://localhost:8787) in your browser.

The Python process serves the static frontend and proxies the two Claude API
calls (column-mapping inference and per-batch classification) — your API key
stays server-side and never touches the browser.

## Spreadsheet format

- Accepts `.xlsx`, `.xls`, or `.csv`.
- **The first row must be the column headers.** Every column after that is
  treated as one row of data, keyed by that header.
- Each row can represent either a single ticket, or a pre-aggregated group of
  tickets (see "Ticket Count" below) — both are supported.

## Using the tool

### 1. Upload your spreadsheet

Drop in the file. The tool shows how many rows and columns it detected.

### 2. Set complexity tier thresholds

A workflow's complexity tier is derived from its aggregate First Resolution
Rate (FRR), Average Handle Time (AHT), and touches-per-ticket:

- **Low**: FRR ≥ Low-FRR **and** AHT ≤ Low-AHT **and** Touches/Ticket ≤
  Low-Touches
- **Medium**: FRR ≥ Med-FRR **and** AHT ≤ Med-AHT
- **High**: everything else

Defaults are Low-FRR 70%, Low-AHT 30 min, Low-Touches 1.7, Med-FRR 50%,
Med-AHT 60 min — all editable in the UI, both before and after classifying
(the dashboard re-tiers live as you change them).

### 3. (Optional) Infer column mapping with AI

Click **"Infer Column Mapping with AI"** to have Claude read your headers and
a sample of rows and guess which column is which:

- **Company** (optional) — the customer/account a row belongs to
- **Hours** and **Touches** (required) — total time spent and number of
  touches/interactions
- **First-Touch-Resolved** (optional) — a yes/no flag or a count of
  first-touch resolutions
- **Ticket Count** (optional) — only needed if your spreadsheet is
  pre-aggregated, i.e. one row already represents a group of N tickets
  (e.g. one row per tag combination per month) rather than one row per
  ticket. Leave blank if every row is a single ticket.
- **Classification text column(s)** — the free-text or tag column(s) that
  describe what the ticket is about; if you pick more than one, they're
  concatenated per row.

This step is optional and always reviewable/editable — you can also just map
every column by hand. **Double check the mapping before classifying**,
especially Ticket Count: mapping it to the wrong column is the most common
cause of a dashboard total that doesn't match your actual ticket count.

### 4. Define your workflow rubric

This is on you — the tool does not generate a rubric for you. Enter it either
by pasting lines into the text box in the format:

```
Name | Category | Description
```

(Category is optional; if omitted it defaults to "General") and clicking
**"Parse pasted text into table"**, or by building it row-by-row directly in
the table. A good rubric is specific rather than generic — e.g. "Password
Reset" and "MFA Setup / Reset" as separate entries rather than one catch-all
"Identity & Access" entry — since the description text is what the
classifier matches tickets against.

### 5. Classify & view the dashboard

Click **"Classify Tickets & Build Dashboard"**. Tickets are batched and sent
to Claude concurrently; each ticket is matched to the single best-fitting
workflow name from your rubric, or left "Unclassified / Other" if nothing
fits well. The dashboard then shows, per workflow: ticket count, % of total,
total hours, % of total hours, AHT, FRR, touches/ticket, and complexity tier.
Filter by company or tier, sort by clicking any column header, and export to
clipboard/Excel (either a single flattened sheet, or one tab per company).
