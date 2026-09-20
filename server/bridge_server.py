"""
Local Bridge Server for Desktop Excel AI Assistant Add-in.
Serves taskpane web assets over HTTP and brokers AI requests
to OpenCode models via `opencode serve` (auto-started on port 4096).
"""

from __future__ import annotations

import os
import sys
import json
import re
import time
import base64
import threading
import subprocess
import urllib.request
import urllib.error
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
SERVER_DIR = Path(__file__).resolve().parent
HOST = "localhost"
PORT = 3000

# Bug 43: Restrict CORS to the add-in's own origin only.
ALLOWED_ORIGIN = "http://localhost:3000"

SYSTEM_PROMPT = """You are an expert Excel AI Assistant running directly inside Microsoft Excel.
The user is working in an active workbook.

When the user asks to inspect, calculate, update, format, clean, or analyze data:
1. Explain what you are doing in plain, friendly language.
2. When you want to modify or update the sheet, provide an ```excel-action code block containing a JSON object with your proposed changes.

EXCEL ACTION SPECIFICATION:
```excel-action
{
  "summary": "Brief 1-sentence description of changes",
  "actions": [
    {
      "type": "set_values",
      "range": "B2:B5",
      "values": [[10], [20], [30], [40]]
    },
    {
      "type": "set_formulas",
      "range": "C2:C5",
      "formulas": [["=A2*B2"], ["=A3*B3"], ["=A4*B4"], ["=A5*B5"]]
    },
    {
      "type": "fill_formula",
      "range": "C2:C20",
      "formula": "=A2*B2"
    },
    {
      "type": "format_range",
      "range": "A1:C1",
      "bold": true,
      "fill": "#107C41",
      "fontColor": "#FFFFFF",
      "fontSize": 12,
      "wrapText": true,
      "horizontalAlignment": "Center"
    },
    {
      "type": "set_number_format",
      "range": "B2:B20",
      "format": "#,##0.00"
    },
    {
      "type": "borders",
      "range": "A1:C5",
      "edges": "all",
      "style": "continuous",
      "weight": "thin",
      "color": "#000000"
    },
    {
      "type": "create_table",
      "range": "A1:C5",
      "tableName": "MyTable",
      "hasHeaders": true
    },
    {
      "type": "sort_range",
      "range": "A1:D20",
      "sortBy": [{"column": 1, "ascending": false}],
      "hasHeaders": true
    },
    {
      "type": "find_replace",
      "range": "A1:Z100",
      "find": "N/A",
      "replace": "0",
      "matchCase": false,
      "matchEntireCell": true
    },
    {
      "type": "conditional_format",
      "range": "B2:B20",
      "rule": "cellValue",
      "operator": "greater",
      "formula1": "100",
      "fill": "#C6EFCE",
      "fontColor": "#006100"
    },
    {
      "type": "merge_range",
      "range": "A1:D1"
    },
    {
      "type": "clear_range",
      "range": "Z1:Z100",
      "what": "contents"
    },
    {
      "type": "add_comment",
      "cell": "B2",
      "text": "Verify this figure with Finance."
    },
    {
      "type": "insert_rows",
      "row": 5,
      "count": 2
    },
    {
      "type": "delete_columns",
      "column": "C",
      "count": 1
    },
    {
      "type": "freeze_panes",
      "range": "B2"
    },
    {
      "type": "hide_sheet",
      "name": "RawData"
    },
    {
      "type": "create_pivot_table",
      "range": "A1:D9",
      "sourceRange": "Sheet1!A1:D9",
      "destSheet": "SalesPivot",
      "destCell": "A1",
      "tableName": "SalesPivot1",
      "rows": ["Product"],
      "columns": ["Region"],
      "values": [{"field": "Amount", "aggregation": "sum"}],
      "filters": ["Salesperson"]
    },
    {
      "type": "create_chart",
      "sheet": "Sheet1",
      "sourceRange": "Sheet1!A1:C13",
      "chartType": "column",
      "title": "Revenue vs Expenses by Month",
      "position": "E2:K20",
      "seriesBy": "columns",
      "dataLabels": false
    }
  ]
}
```

Rules for actions:
- Range addresses must be valid Excel A1-style notations (e.g. "A1", "B2:D10").
- "values" and "formulas" MUST be 2D arrays corresponding to rows and columns.
- Formulas must start with '=' and use uppercase formula names (SUM, AVERAGE, IF, VLOOKUP, XLOOKUP).
- fill_formula writes the formula to the top-left cell and AutoFills it across
  the range, so relative references adjust automatically.
- To write to a SPECIFIC sheet (not active), add "sheet": "SheetName" to any
  range-level action. Example: {"type": "set_values", "sheet": "Summary", "range": "A1", "values": [["Total"]]}
- Structure edits: insert_rows/delete_rows need 1-based "row" + "count";
  insert_columns/delete_columns take "column" as 0-based index OR A1 letters
  ("C") + "count". delete_sheet / hide_sheet / unhide_sheet need "name".
- freeze_panes pins TOP rows / LEFT columns (they stay visible while
  scrolling). ALWAYS use row/column COUNTS, never cell addresses:
  "freeze 1st row" -> {"type": "freeze_panes", "rows": 1};
  "freeze first 2 rows" -> {"type": "freeze_panes", "rows": 2};
  "freeze 1st column" -> {"type": "freeze_panes", "columns": 1};
  "freeze 1st row AND 1st column" -> {"type": "freeze_panes",
  "rows": 1, "columns": 1}. unfreeze_panes reverts it.
- Row/column visibility (reversible pairs, no backup needed):
  {"type": "hide_rows", "row": 1, "count": 2} /
  {"type": "unhide_rows", "row": 1, "count": 2} (1-based row + count);
  {"type": "hide_columns", "column": "C", "count": 1} /
  {"type": "unhide_columns", "column": "C"} ("column" is 0-based index
  OR A1 letters). If the user says rows/columns "disappeared", check
  hiding first and unhide them.
- sort_range: "sortBy" is a list of {"column", "ascending"}. "column" is the
  0-based offset INSIDE the range, or an absolute A1 letter ("B").
- find_replace: omit "range" to search the whole used range of the sheet.
  "matchEntireCell": true replaces whole-cell matches only.
- Number formats: "0.00", "#,##0", "0%", "dd-mmm-yyyy", "\"$\"#,##0.00".
- borders "edges": all|outer|inner|top|bottom|left|right; "style":
  continuous|dash|dotted|double; "weight": hairline|thin|medium|thick.
- conditional_format "rule": cellValue (needs "operator": between|equal|
  notequal|greater|less|greaterorequal|lessorequal + "formula1"/"formula2"),
  colorScale (optional minColor/midColor/maxColor), dataBar (optional
  "fill"), or "clear" to remove rules. Use clear_conditional_format to
  wipe all rules from a range.
- format_range also supports "italic", "wrapText", "horizontalAlignment"
  (Left|Center|Right|Justify) and "verticalAlignment" (Top|Center|Bottom).
- merge_range merges a range into one cell (Excel keeps the TOP-LEFT value
  and clears the rest - only merge title/header rows, never data). Pair it
  with format_range horizontalAlignment Center + verticalAlignment Center
  for "merge and center". {"type": "merge_range", "range": "A1:D1",
  "unmerge": true} splits it back.
- clear_range "what": all|contents|formats. autofit accepts "rows": true
  to also autofit row heights.
- For create_pivot_table: "rows", "columns", "values", "filters" must match exact column header names from the source data. "sourceRange" should include the sheet name (e.g. "Sheet1!A1:D9").
- For create_chart: "chartType" is bar|column|stackedbar|stackedcolumn|
  line|linemarkers|pie|doughnut|area (default column). "sourceRange"
  should include the sheet name. "position" is the anchor cells for the
  chart object (e.g. "E2:K20") on an empty area - check context values so
  the chart does not cover data. "seriesBy" is columns (default, one
  series per data column) or rows. Optional: "title", "dataLabels" true,
  "legend": false, "legendPosition" top|bottom|left|right,
  "valueAxisTitle", "categoryAxisTitle". This creates a standard chart
  object (static snapshot, also works on pivot output ranges). Always pass
  "seriesNames" (e.g. ["Revenue", "Expenses"]) so hover tooltips show real
  names instead of "Series 1". Existing chart names are listed in context
  - delete by EXACT name, never guess:
  {"type": "delete_chart", "sheet": "Sheet1", "name": "Chart 1"}, or
  {"type": "delete_chart", "sheet": "Sheet1", "all": true} to clear them.
- For AXES and GRIDLINES use format_chart_axis (chart by EXACT context
  name, omit "chart" only if the sheet has exactly one):
  vertical guides every N on the X axis = {"type": "format_chart_axis",
  "sheet": "Sheet2", "chart": "RGB Trend", "axis": "category",
  "majorUnit": 5, "majorGridlines": true}. Horizontal Y lines =
  "axis": "value" with majorUnit. Optional: "minorUnit",
  "numberFormat" (e.g. "#,##0"), "title", "minorGridlines". You CAN set
  gridline spacing directly - never claim otherwise and never fall back
  to helper-column workarounds or manual steps for it.
- IMPORTANT: You CAN create bar/line/pie chart objects with create_chart.
  Do NOT tell the user to insert charts manually - build them.
- To create a new blank sheet use: {"type": "add_sheet", "name": "SheetName", "activate": true}
- To rename a sheet use: {"type": "rename_sheet", "from": "OldName", "to": "NewName"}
- To roll a sheet back to its hidden backup use:
  {"type": "restore_backup", "sheet": "SheetName"}; to remove the backup
  copy use {"type": "delete_backup", "sheet": "SheetName"}.
- FILE operations: you CANNOT Save, Save As, or rename the open workbook
  file itself (Excel owns it; no such API exists - never claim otherwise,
  never give manual Save-As steps as if you did it). You CAN save a FULL
  .xlsx COPY under any name via {"type": "save_copy_as",
  "filename": "colourcodeai"} - the local bridge writes it and the result
  states the exact path; relay that path to the user. If the user only
  wants the current file saved, tell them to press Ctrl+S.
- Tables known to exist are listed in context - prefer writing new data
  next to them, not inside them, unless asked.
- You will be given context for ALL sheets in the workbook (capped at 500
  rows x 100 cols per sheet; "truncated": true means data was cut - say so
  and work with the visible part, or ask the user to select the range they
  care about, since the SELECTION values are always sent in full detail).
- Only emit the excel-action block if actual spreadsheet modifications are needed. If just answering a question, no action block is necessary.
- IMPORTANT: You CAN create real PivotTables using create_pivot_table. Do NOT tell the user to create PivotTables manually.
- IMPORTANT: You CAN create new sheets using add_sheet. Do NOT tell the user to manually create sheets.
- IMPORTANT: You CAN sort, find/replace, freeze panes, hide sheets, and add
  conditional formatting directly. Do NOT give manual step-by-step
  instructions for anything the actions above can do.
- SAFETY / FULL BACKUPS (the engine guarantees this - explain it briefly on
  risky edits): BEFORE the batch runs, EVERY sheet it touches is fully
  backed up to a hidden AI_Backup__<Sheet> copy (values, formulas,
  formats, borders, conditional formats) PLUS its table list
  (name/address/style) and chart definitions. Previous backup replaced.
- Undo reverts step-by-step (newest first). Structural undos and the
  Restore button ROLL THE WHOLE SHEET BACK - values, formats, tables,
  charts - and the sheet itself is never deleted during rollback, so
  cross-sheet references stay intact. PivotTables and cell comments are
  NOT rebuilt: say Ctrl+S first before pivot-heavy surgery.
- You can also emit {"type": "restore_backup", "sheet": "X"} /
  {"type": "delete_backup", "sheet": "X"} yourself.
"""

ANSI_ESCAPE = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

def clean_cli_output(text: str) -> str:
    """Strip ANSI escape codes and progress-spinner lines from CLI output."""
    text = ANSI_ESCAPE.sub('', text)
    lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith('\r') or stripped.startswith('\x08'):
            continue
        if re.match(r'^[>|]\s*(build|thinking|running|loading)', stripped, re.IGNORECASE):
            continue
        lines.append(line)
    return '\n'.join(lines).strip()

MAX_ACTIONS = 50  # Bug 53: cap extracted action count

def extract_excel_actions(response_text: str):
    """Extracts structured excel-action JSON blocks from model response.
    Bug 53: Caps at MAX_ACTIONS to prevent runaway batches.
    """
    pattern = r"```excel-action\s*([\s\S]*?)\s*```"
    matches = re.findall(pattern, response_text, re.IGNORECASE)
    actions = []
    for match in matches:
        try:
            parsed = json.loads(match.strip())
            actions.extend(parsed.get("actions", []))
        except Exception as e:
            print(f"Failed to parse action block: {e}", flush=True)
    return actions[:MAX_ACTIONS]


def build_excel_context_block(context: dict | None) -> str:
    """Formats the full multi-sheet Excel context for inclusion in a prompt."""
    if not context:
        return ""
    block = "CURRENT EXCEL CONTEXT:\n"
    block += f"Active Sheet: {context.get('activeSheet') or context.get('sheetName', 'Sheet1')}\n"
    if context.get("readWarnings"):
        block += f"WARNING: {context.get('readWarnings')}\n"
    if context.get("address"):
        block += f"Active Sheet Range: {context.get('address')}\n"
    if context.get("totalRows") or context.get("totalCols"):
        block += f"Active Sheet Size: {context.get('totalRows')} rows x {context.get('totalCols')} cols\n"
    if context.get("truncated"):
        block += "NOTE: Active sheet data below is TRUNCATED (500 rows x 100 cols cap).\n"
    if context.get("values"):
        block += f"Active Sheet Values: {json.dumps(context.get('values'))}\n"
    if context.get("formulas"):
        block += f"Active Sheet Formulas: {json.dumps(context.get('formulas'))}\n"
    if context.get("tables"):
        block += f"Active Sheet Tables: {json.dumps(context.get('tables'))}\n"
    if context.get("chartCount"):
        block += f"Active Sheet Charts: {context.get('chartCount')}"
        if context.get("charts"):
            block += f" ({', '.join(context.get('charts'))})"
        block += "\n"

    sel = context.get("selection") or {}
    if sel.get("address"):
        block += f"\nUSER SELECTION: {sel.get('address')}\n"
        if sel.get("truncated"):
            block += "(selection values truncated)\n"
        if sel.get("values"):
            block += f"Selection Values: {json.dumps(sel.get('values'))}\n"

    if context.get("namedRanges"):
        block += f"\nDefined Names: {json.dumps(context.get('namedRanges'))}\n"

    other_sheets = [s for s in (context.get("sheets") or [])
                    if s.get("name") != (context.get("activeSheet") or context.get("sheetName"))]
    if other_sheets:
        block += "\nOTHER SHEETS IN WORKBOOK:\n"
        for s in other_sheets:
            if s.get("backup"):
                block += f"\nSheet: {s['name']} (hidden AI backup copy - ignore)\n"
                continue
            block += f"\nSheet: {s['name']}"
            if s.get("visible") is False:
                block += " (HIDDEN)"
            if s.get("totalRows") or s.get("totalCols"):
                block += f" [{s.get('totalRows')} rows x {s.get('totalCols')} cols"
                if s.get("truncated"):
                    block += ", TRUNCATED"
                block += "]"
            block += "\n"
            if s.get("tables"):
                block += f"Tables: {json.dumps(s['tables'])}\n"
            if s.get("commentCount"):
                block += f"Comments: {s['commentCount']}\n"
            if s.get("chartCount"):
                block += f"Charts: {s['chartCount']}"
                if s.get("charts"):
                    block += f" ({', '.join(s['charts'])})"
                block += "\n"
            if s.get("values"):
                block += f"Values: {json.dumps(s['values'])}\n"
    return block + "\n"


OPENCODE_SERVER = "http://127.0.0.1:4096"
OPENCODE_DEFAULT_MODEL = "opencode/muse-spark-1.3-contributor-free"

# Bug 1/2: Per-conversation session IDs instead of one global session.
# Key: conversationId (UUID sent by frontend), Value: OpenCode session ID.
_opencode_sessions: dict[str, str] = {}
# Bug 3: Per-conversation lock to serialize concurrent requests on the same session.
_session_locks: dict[str, threading.Lock] = {}
_sessions_mutex = threading.Lock()  # guards the dicts above

# Bug 41: Per-filename lock to prevent concurrent save-copy races.
_save_locks: dict[str, threading.Lock] = {}
_save_locks_mutex = threading.Lock()


def _get_session_lock(conv_id: str) -> threading.Lock:
    """Return (creating if needed) the per-conversation request lock."""
    with _sessions_mutex:
        if conv_id not in _session_locks:
            _session_locks[conv_id] = threading.Lock()
        return _session_locks[conv_id]


def _get_save_lock(filename: str) -> threading.Lock:
    """Return (creating if needed) the per-filename write lock."""
    with _save_locks_mutex:
        if filename not in _save_locks:
            _save_locks[filename] = threading.Lock()
        return _save_locks[filename]


def _find_opencode_exe() -> str:
    """Resolve the real opencode binary (prefers .exe over npm .cmd/.ps1 shims)."""
    import shutil
    candidates = [
        Path(os.path.expandvars(r"%AppData%\npm\node_modules\opencode-ai\bin\opencode.exe")),
        Path.home() / "AppData" / "Roaming" / "npm" / "node_modules" / "opencode-ai" / "bin" / "opencode.exe",
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    for name in ["opencode.exe", "opencode.cmd", "opencode"]:
        found = shutil.which(name)
        if found:
            return found
    return "opencode"


def _opencode_port_open() -> bool:
    import socket
    s = socket.socket()
    s.settimeout(1.5)
    try:
        s.connect(("127.0.0.1", 4096))
        return True
    except OSError:
        return False
    finally:
        s.close()


def ensure_opencode_serve(timeout_s: int = 25) -> None:
    """Start `opencode serve --port 4096` detached if nothing is listening."""
    if _opencode_port_open():
        return
    exe = _find_opencode_exe()
    print(f"[OpenCode] No serve on 127.0.0.1:4096 - starting `{exe} serve --port 4096` ...", flush=True)
    kwargs: dict = {
        "cwd": str(BASE_DIR),
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
        "stdin": subprocess.DEVNULL,
    }
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen([exe, "serve", "--port", "4096"], **kwargs)
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(1)
        if _opencode_port_open():
            print("[OpenCode] serve is up on 127.0.0.1:4096.", flush=True)
            return
    raise RuntimeError("opencode serve did not start on 127.0.0.1:4096")


def _get_or_create_session(conv_id: str, model_id: str = OPENCODE_DEFAULT_MODEL) -> str:
    """Return the existing OpenCode session for conv_id, or create a new one.
    Bug 1/2: Each conversationId gets its own isolated OpenCode session.
    Caller must already hold _get_session_lock(conv_id).
    """
    with _sessions_mutex:
        if conv_id in _opencode_sessions:
            return _opencode_sessions[conv_id]

    ensure_opencode_serve()
    try:
        payload = json.dumps({"modelID": model_id}).encode()
        req = urllib.request.Request(
            f"{OPENCODE_SERVER}/session",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            sess = json.loads(resp.read())
            new_sess_id = sess["id"]
            with _sessions_mutex:
                _opencode_sessions[conv_id] = new_sess_id
            print(f"[OpenCode] Created session {new_sess_id} for conversation {conv_id[:8]}", flush=True)
            return new_sess_id
    except Exception as e:
        raise RuntimeError(f"Failed to create OpenCode session: {e}")


def fetch_opencode_models():
    """Model list with reasoning-effort variants + context limits via serve."""
    try:
        ensure_opencode_serve()
        req = urllib.request.Request(f"{OPENCODE_SERVER}/config/providers", method="GET")
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
        provs = data if isinstance(data, list) else data.get("providers", [])
        if isinstance(provs, dict):
            provs = list(provs.values())
        models = []
        for p in provs:
            if not isinstance(p, dict):
                continue
            pid = p.get("id", "opencode")
            ms = p.get("models") or {}
            items = ms.values() if isinstance(ms, dict) else (ms or [])
            for m in items:
                if not isinstance(m, dict):
                    continue
                mid = m.get("id")
                if not mid:
                    continue
                full = str(mid) if "/" in str(mid) else f"{pid}/{mid}"
                variants = sorted((m.get("variants") or {}).keys())
                limit = m.get("limit") or {}
                models.append({
                    "id": f"opencode:{full}",
                    "name": m.get("name") or str(mid),
                    "provider": "opencode",
                    "variants": variants,
                    "context": limit.get("context"),
                })
        if models:
            return models
    except Exception as e:
        print(f"Failed to fetch opencode models via serve: {e}", flush=True)

    models = []
    try:
        res = subprocess.run([_find_opencode_exe(), "models"], capture_output=True, text=True, timeout=15)
        for line in res.stdout.split('\n'):
            line = line.strip()
            if line.startswith("opencode/"):
                models.append({
                    "id": f"opencode:{line}",
                    "name": line.split('/')[-1],
                    "provider": "opencode",
                    "variants": [],
                    "context": None,
                })
    except Exception as e:
        print(f"Failed to fetch opencode models via CLI: {e}", flush=True)

    if not models:
        models = [{
            "id": f"opencode:{OPENCODE_DEFAULT_MODEL}",
            "name": "muse-spark-1.3-contributor-free",
            "provider": "opencode",
            "variants": ["minimal", "low", "medium", "high", "xhigh"],
            "context": None,
        }]
    return models


def opencode_send(text: str, conv_id: str, model: str | None = None,
                  variant: str | None = None, timeout: int = 180) -> str:
    """Send one message on the session for conv_id; return assistant text."""
    model_id = model or OPENCODE_DEFAULT_MODEL
    sess_id = _get_or_create_session(conv_id, model_id)
    body: dict = {
        "parts": [{"type": "text", "text": text}],
        "modelID": model_id,
    }
    if variant:
        body["variant"] = variant
    payload = json.dumps(body).encode()
    req = urllib.request.Request(
        f"{OPENCODE_SERVER}/session/{sess_id}/message",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        result = json.loads(resp.read())
    parts = result.get("parts", []) if isinstance(result, dict) else []
    texts = [p.get("text", "") for p in parts if p.get("type") == "text"]
    return clean_cli_output("\n".join(texts)) or clean_cli_output(str(result))


def run_opencode(prompt: str, context: dict | None, conv_id: str,
                 model: str | None = None, variant: str | None = None) -> str:
    """Executes prompt via the persistent OpenCode server REST API.
    Bug 3: Holds the per-conversation lock for the duration of the call.
    Bug 19: Raises on failure so caller returns HTTP 500.
    """
    full_text = SYSTEM_PROMPT + "\n\n"
    full_text += build_excel_context_block(context)
    full_text += f"USER REQUEST: {prompt}"
    lock = _get_session_lock(conv_id)
    with lock:
        return opencode_send(full_text, conv_id=conv_id, model=model, variant=variant)


# Bug 18: Whole-word regex + question-phrase exclusion. Bug 52: deduplicated.
_MODIFY_KEYWORDS = [
    "add", "insert", "put", "write", "fill", "calculate",
    "sum", "total", "format", "bold", "color", "colour", "update",
    "create", "make", "set", "apply", "append", "delete",
    "remove", "sort", "multiply", "divide", "subtract",
    "replace", "find", "freeze", "hide", "border",
    "highlight", "comment", "clear", "autofit",
    "restore", "backup", "revert",
    "chart", "graph", "bar", "line", "pie",
    "merge", "center", "centre",
    "unhide", "hidden", "disappeared",
    "save", "download", "copy", "rename", "file",
    "gridline", "axis", "units",
]
_QUESTION_PATTERNS = re.compile(
    r'\b(what\s+is|what\s+are|explain|how\s+does|how\s+do|why\s+is|why\s+does'
    r'|tell\s+me\s+about|describe)\b',
    re.IGNORECASE
)
_MODIFY_RE = re.compile(
    r'\b(' + '|'.join(re.escape(kw) for kw in _MODIFY_KEYWORDS) + r')\b',
    re.IGNORECASE
)


class BridgeHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        first = args[0] if args else ""
        if isinstance(first, str) and "/api/" in first:
            super().log_message(format, *args)

    def end_headers(self):
        # Bug 43: Restrict CORS to localhost:3000 only.
        self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        url_path = self.path.split("?")[0]

        if url_path == "/api/models":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            models = fetch_opencode_models()
            self.wfile.write(json.dumps({"models": models}).encode("utf-8"))
            return

        # Serve static assets
        if url_path in ["/", "/taskpane.html"]:
            file_path = BASE_DIR / "src" / "taskpane.html"
            content_type = "text/html"
        elif url_path == "/manifest.xml":
            file_path = BASE_DIR / "manifest.xml"
            content_type = "text/xml"
        elif url_path == "/taskpane.css":
            file_path = BASE_DIR / "src" / "taskpane.css"
            content_type = "text/css"
        elif url_path == "/taskpane.js":
            file_path = BASE_DIR / "src" / "taskpane.js"
            content_type = "application/javascript"
        elif url_path == "/excel_actions.js":
            file_path = BASE_DIR / "src" / "excel_actions.js"
            content_type = "application/javascript"
        elif url_path.startswith("/assets/"):
            # Bug 42: Canonicalize and restrict to assets directory.
            asset_rel = url_path[len("/assets/"):]
            asset_rel = asset_rel.replace("\\", "/")
            # Remove path traversal components
            safe_parts = [p for p in asset_rel.split("/") if p and p != ".." and p != "."]
            if not safe_parts:
                self.send_error(400, "Bad asset path")
                return
            assets_root = (BASE_DIR / "assets").resolve()
            file_path = (assets_root / "/".join(safe_parts)).resolve()
            try:
                file_path.relative_to(assets_root)
            except ValueError:
                self.send_error(403, "Path not allowed")
                return
            asset_name = safe_parts[-1]
            content_type = "image/png" if asset_name.endswith(".png") else "application/octet-stream"
        else:
            self.send_error(404, "File Not Found")
            return

        if file_path.exists():
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            self.end_headers()
            with open(file_path, "rb") as f:
                self.wfile.write(f.read())
        else:
            self.send_error(404, f"File {file_path.name} not found")

    def do_POST(self):
        url_path = self.path.split("?")[0]

        if url_path == "/api/save-copy":
            # Bug 39: Check Content-Length BEFORE reading the body.
            content_len = int(self.headers.get("Content-Length", 0))
            if content_len > 200 * 1024 * 1024:
                self.send_error(400, "File too large (200 MB cap)")
                return
            post_body = self.rfile.read(content_len)
            try:
                data = json.loads(post_body.decode("utf-8"))
            except Exception as e:
                self.send_error(400, f"Invalid JSON payload: {e}")
                return
            raw_name = str(data.get("filename", "") or "workbook-copy")
            # Bug 47: Preserve original extension (.xlsm/.xlsb); only append .xlsx if absent.
            safe_base = re.sub(r'[\\/:*?"<>|]', "", raw_name).strip()[:80] or "workbook-copy"
            ext_match = re.search(r'\.(xlsx|xlsm|xlsb|xls)$', safe_base, re.IGNORECASE)
            safe = safe_base if ext_match else (safe_base + ".xlsx")
            saved_dir = BASE_DIR / "saved"
            saved_dir.mkdir(exist_ok=True)
            target_base = saved_dir / safe
            # Bug 41: Per-filename lock to prevent concurrent write races.
            file_lock = _get_save_lock(safe)
            with file_lock:
                target = target_base
                if target.exists():
                    stem = Path(safe).stem
                    ext = Path(safe).suffix
                    i = 1
                    while (saved_dir / f"{stem}-{i}{ext}").exists():
                        i += 1
                    target = saved_dir / f"{stem}-{i}{ext}"
                try:
                    raw = base64.b64decode(data.get("data", ""))
                except Exception:
                    self.send_error(400, "Invalid base64 file data")
                    return
                if raw[:2] != b"PK":
                    self.send_error(400, "Upload is not an .xlsx (zip) file")
                    return
                if len(raw) > 200 * 1024 * 1024:
                    self.send_error(400, "File too large (200 MB cap)")
                    return
                target.write_bytes(raw)
            print(f"[SaveCopy] wrote {target} ({len(raw)} bytes)", flush=True)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({
                "ok": True, "filename": target.name, "path": str(target)
            }).encode("utf-8"))
            return

        if url_path == "/api/chat":
            # Bug 39: Check Content-Length before reading.
            content_len = int(self.headers.get("Content-Length", 0))
            if content_len > 200 * 1024 * 1024:
                self.send_error(400, "Request too large (200 MB cap)")
                return
            post_body = self.rfile.read(content_len)
            try:
                data = json.loads(post_body.decode("utf-8"))
            except Exception as e:
                self.send_error(400, f"Invalid JSON payload: {e}")
                return

            provider_raw = data.get("provider", "opencode")
            sub_model = None
            if ":" in provider_raw:
                _, sub_model = provider_raw.split(":", 1)
            elif provider_raw and "/" in provider_raw:
                sub_model = provider_raw

            prompt = data.get("prompt", "")
            context = data.get("context", {})
            variant = data.get("variant") or None
            # Bug 1/2: Use per-conversation ID sent by the frontend.
            conv_id = str(data.get("conversationId") or "default")

            print(f"Received query (conv: {conv_id[:8]}..., model: {sub_model or OPENCODE_DEFAULT_MODEL}, "
                  f"variant: {variant}) | prompt length: {len(prompt)}", flush=True)

            # Bug 19: Raise on failure -> HTTP 500.
            try:
                response_text = run_opencode(prompt, context, conv_id=conv_id,
                                             model=sub_model, variant=variant)
            except Exception as e:
                error_msg = str(e)
                print(f"[OpenCode] Error: {error_msg}", flush=True)
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": f"AI backend error: {error_msg}"}).encode("utf-8"))
                return

            actions = extract_excel_actions(response_text)

            # Bug 18: Whole-word matching + question-phrase exclusion.
            # Bug 52: Deduplicated via _MODIFY_RE.
            prompt_lower = prompt.lower()
            is_pure_question = bool(_QUESTION_PATTERNS.search(prompt_lower))
            wants_modification = (
                not is_pure_question and
                bool(_MODIFY_RE.search(prompt_lower))
            )

            if not actions and wants_modification and context and context.get("address"):
                extraction_prompt = (
                    f"Based on the following AI assistant answer, produce ONLY a JSON excel-action block "
                    f"to apply changes to Excel.\n\nAI ANSWER:\n{response_text}\n\n"
                    f"EXCEL CONTEXT:\nSheet: {context.get('sheetName','Sheet1')}\n"
                    f"Range: {context.get('address')}\n"
                    f"Values: {json.dumps(context.get('values', []))}\n\n"
                    f"Output ONLY a ```excel-action block with the required changes. No other text."
                )
                try:
                    # Bug 17: Use the user-chosen model, not the default.
                    # Bug 3: The per-conversation lock is already held by run_opencode above,
                    # so we call opencode_send directly here (still within the same conv).
                    lock = _get_session_lock(conv_id)
                    with lock:
                        action_text = opencode_send(
                            extraction_prompt,
                            conv_id=conv_id,
                            model=sub_model or OPENCODE_DEFAULT_MODEL,
                            variant=variant,
                            timeout=90
                        )
                    actions = extract_excel_actions(action_text)
                    if actions:
                        print(f"[Auto-Extract] Found {len(actions)} action(s) from second pass", flush=True)
                except Exception as e:
                    print(f"[Auto-Extract] Failed: {e}", flush=True)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({
                "response": response_text,
                "actions": actions
            }).encode("utf-8"))
            return

        self.send_error(404, "API route not found")


def main():
    server_address = (HOST, PORT)
    httpd = ThreadingHTTPServer(server_address, BridgeHandler)
    # Bug 44/45: Server is plain HTTP only. HTTPS/SSL code removed.
    # Excel Desktop WebView2 trusts HTTP on localhost without certificates.
    print("=" * 60, flush=True)
    print(f" Excel AI Assistant Bridge Server Running", flush=True)
    print(f" URL: http://{HOST}:{PORT}/taskpane.html", flush=True)
    print(f" Provider: OpenCode (auto-starts `opencode serve --port 4096`)", flush=True)
    print(f" Session isolation: per-conversation UUID", flush=True)
    print("=" * 60, flush=True)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down bridge server...", flush=True)
        httpd.server_close()

if __name__ == "__main__":
    main()
