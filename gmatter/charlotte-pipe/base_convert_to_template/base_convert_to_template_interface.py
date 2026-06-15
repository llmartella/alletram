"""
SAP Sales Data Processor — Local Web UI
Run with:  python base_convert_to_template_interface.py
Then open: http://127.0.0.1:8080
"""

import json
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request

from base_convert_to_template import (
    FIELD_LABELS,
    INPUT_DIR,
    OUTPUT_HEADERS,
    REQUIRED_FIELDS,
    _normalise_header,
    check_missing_required_fields,
    headers_signature,
    process_file_with_map,
    save_mapping,
    scan_input_files,
)

app = Flask(__name__)

# In-memory state for the current run
_scan_results: list = []


# ============================================================
# HTML TEMPLATE
# ============================================================

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SAP Sales Processor</title>
<style>
  /* ── Tokens ── */
  :root {
    --ink:        #1a1f2e;
    --ink-mid:    #4a5168;
    --ink-light:  #8890a4;
    --surface:    #f5f6f9;
    --card:       #ffffff;
    --accent:     #2563eb;
    --accent-dim: #dbeafe;
    --success:    #16a34a;
    --success-dim:#dcfce7;
    --warn:       #d97706;
    --warn-dim:   #fef3c7;
    --danger:     #dc2626;
    --danger-dim: #fee2e2;
    --border:     #e2e5ed;
    --radius:     8px;
    --mono: "SF Mono", "Fira Code", "Consolas", monospace;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    background: var(--surface);
    color: var(--ink);
    min-height: 100vh;
  }

  /* ── Layout ── */
  header {
    background: var(--card);
    border-bottom: 1px solid var(--border);
    padding: 0 32px;
    height: 56px;
    display: flex;
    align-items: center;
    gap: 12px;
    position: sticky;
    top: 0;
    z-index: 10;
  }
  header svg { flex-shrink: 0; }
  header h1 { font-size: 15px; font-weight: 600; letter-spacing: -.01em; }
  header span { font-size: 13px; color: var(--ink-light); }

  main { max-width: 860px; margin: 0 auto; padding: 32px 24px 64px; }

  /* ── Scan bar ── */
  .scan-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 24px;
    gap: 12px;
  }
  .scan-bar h2 { font-size: 18px; font-weight: 600; }

  /* ── Buttons ── */
  button {
    font: inherit;
    border: none;
    border-radius: var(--radius);
    cursor: pointer;
    transition: opacity .15s, background .15s;
  }
  button:disabled { opacity: .45; cursor: not-allowed; }
  .btn-primary {
    background: var(--accent);
    color: #fff;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 500;
  }
  .btn-primary:hover:not(:disabled) { opacity: .88; }
  .btn-ghost {
    background: transparent;
    color: var(--accent);
    padding: 8px 14px;
    font-size: 13px;
    border: 1px solid var(--border);
  }
  .btn-ghost:hover:not(:disabled) { background: var(--accent-dim); }
  .btn-success {
    background: var(--success);
    color: #fff;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 500;
  }
  .btn-success:hover:not(:disabled) { opacity: .88; }

  /* ── File cards ── */
  .file-list { display: flex; flex-direction: column; gap: 10px; }

  .file-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
  }
  .file-card-header {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 16px;
  }
  .file-name {
    flex: 1;
    font-size: 13px;
    font-weight: 500;
    font-family: var(--mono);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .badge {
    font-size: 11px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 99px;
    white-space: nowrap;
    flex-shrink: 0;
    letter-spacing: .02em;
    text-transform: uppercase;
  }
  .badge-ready    { background: var(--success-dim); color: var(--success); }
  .badge-mapping  { background: var(--warn-dim);    color: var(--warn); }
  .badge-done     { background: var(--accent-dim);  color: var(--accent); }
  .badge-error    { background: var(--danger-dim);  color: var(--danger); }
  .badge-saved    { background: #f0fdf4; color: #15803d; }

  .row-count {
    font-size: 12px;
    color: var(--ink-light);
    white-space: nowrap;
  }

  /* ── Mapping panel ── */
  .mapping-panel {
    border-top: 1px solid var(--border);
    padding: 16px 16px 20px;
    background: #fafbfc;
  }
  .mapping-panel h3 {
    font-size: 13px;
    font-weight: 600;
    color: var(--ink-mid);
    margin-bottom: 4px;
  }
  .mapping-panel p {
    font-size: 12px;
    color: var(--ink-light);
    margin-bottom: 14px;
    line-height: 1.5;
  }

  .field-grid {
    display: grid;
    grid-template-columns: 180px 1fr;
    gap: 8px 12px;
    align-items: center;
    margin-bottom: 16px;
  }
  .field-label {
    font-size: 13px;
    font-weight: 500;
  }
  .field-label .required-star { color: var(--danger); margin-left: 2px; }

  select {
    font: inherit;
    font-size: 13px;
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 6px 10px;
    background: #fff;
    color: var(--ink);
    width: 100%;
    cursor: pointer;
  }
  select:focus { outline: 2px solid var(--accent); outline-offset: 1px; }

  /* ── Save prompt ── */
  .save-prompt {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 14px;
    background: var(--accent-dim);
    border-radius: 6px;
    margin-top: 12px;
    font-size: 13px;
  }
  .save-prompt label { display: flex; align-items: center; gap: 6px; cursor: pointer; }
  .save-prompt input[type=checkbox] { width: 14px; height: 14px; cursor: pointer; }

  .mapping-actions {
    display: flex;
    gap: 10px;
    margin-top: 16px;
  }

  /* ── Summary bar ── */
  .summary {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 16px 20px;
    margin-bottom: 24px;
    display: flex;
    gap: 24px;
    align-items: center;
    flex-wrap: wrap;
  }
  .summary-stat { display: flex; flex-direction: column; gap: 2px; }
  .summary-stat .num { font-size: 22px; font-weight: 700; }
  .summary-stat .lbl { font-size: 11px; color: var(--ink-light); text-transform: uppercase; letter-spacing: .06em; }
  .stat-ready   .num { color: var(--success); }
  .stat-pending .num { color: var(--warn); }
  .stat-done    .num { color: var(--accent); }

  /* ── Empty / loading ── */
  .empty {
    text-align: center;
    padding: 48px 24px;
    color: var(--ink-light);
    font-size: 14px;
    line-height: 1.8;
  }
  .spinner {
    display: inline-block;
    width: 18px; height: 18px;
    border: 2px solid var(--border);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin .7s linear infinite;
    vertical-align: middle;
    margin-right: 6px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── Toast ── */
  #toast {
    position: fixed;
    bottom: 24px; right: 24px;
    background: var(--ink);
    color: #fff;
    padding: 10px 18px;
    border-radius: var(--radius);
    font-size: 13px;
    opacity: 0;
    transition: opacity .2s;
    pointer-events: none;
    max-width: 320px;
  }
  #toast.show { opacity: 1; }
  #toast.toast-error { background: var(--danger); }
</style>
</head>
<body>

<header>
  <svg width="22" height="22" viewBox="0 0 22 22" fill="none">
    <rect width="22" height="22" rx="5" fill="#2563eb"/>
    <path d="M5 7h12M5 11h8M5 15h10" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/>
  </svg>
  <h1>SAP Sales Processor</h1>
  <span id="dir-label"></span>
</header>

<main>
  <div class="scan-bar">
    <h2>Input Files</h2>
    <button class="btn-primary" id="scan-btn" onclick="scanFiles()">
      Scan Folder
    </button>
  </div>

  <div id="summary" class="summary" style="display:none"></div>
  <div id="file-list" class="file-list"></div>
  <div id="loading" class="empty" style="display:none">
    <span class="spinner"></span> Scanning files…
  </div>
  <div id="empty" class="empty" style="display:none">
    No Excel files found in the Input folder.<br>
    Add .xlsx or .xlsm files and click <strong>Scan Folder</strong>.
  </div>
</main>

<div id="toast"></div>

<script>
// ── State ──────────────────────────────────────────────────
let files = [];   // array of file info objects from server

// ── API helpers ───────────────────────────────────────────
async function api(url, opts = {}) {
  const r = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...opts,
  });
  return r.json();
}

// ── Toast ─────────────────────────────────────────────────
function toast(msg, isError = false) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "show" + (isError ? " toast-error" : "");
  setTimeout(() => el.className = "", 3000);
}

function toggleMapping(filename) {
  const panel = document.getElementById(`review-${filename}`);
  const btn = panel.previousElementSibling.previousElementSibling;
  if (!panel) return;
  const isOpen = panel.style.display !== "none";
  panel.style.display = isOpen ? "none" : "block";
  // Find the button and update its label
  const card = document.getElementById("card-" + CSS.escape(filename));
  const reviewBtn = card.querySelector(".btn-ghost");
  if (reviewBtn) reviewBtn.textContent = isOpen ? "Review Mapping" : "Hide Mapping";
}

// ── Scan ──────────────────────────────────────────────────
async function scanFiles() {
  document.getElementById("loading").style.display = "block";
  document.getElementById("file-list").innerHTML = "";
  document.getElementById("summary").style.display = "none";
  document.getElementById("empty").style.display = "none";
  document.getElementById("scan-btn").disabled = true;

  try {
    const data = await api("/api/scan");
    files = data.files;
    document.getElementById("dir-label").textContent = data.input_dir;
    renderAll();

    // Auto-process any files that are already ready
    const ready = files.filter(f => f.status === "ready");
    for (const f of ready) {
      await processFile(f.filename);
    }
  } catch (e) {
    toast("Scan failed: " + e.message, true);
  } finally {
    document.getElementById("loading").style.display = "none";
    document.getElementById("scan-btn").disabled = false;
  }
}

// ── Render all ────────────────────────────────────────────
function renderAll() {
  const list = document.getElementById("file-list");
  list.innerHTML = "";

  if (!files.length) {
    document.getElementById("empty").style.display = "block";
    return;
  }

  updateSummary();

  files.forEach(f => {
    const card = buildCard(f);
    list.appendChild(card);
  });
}

function updateSummary() {
  const pending = files.filter(f => f.status === "needs_mapping").length;
  const done    = files.filter(f => f.status === "done").length;
  const total   = files.length;

  const el = document.getElementById("summary");
  el.style.display = "flex";
  el.innerHTML = `
    <div class="summary-stat stat-done"><span class="num">${total}</span><span class="lbl">Total Files</span></div>
    <div class="summary-stat stat-done"><span class="num">${done}</span><span class="lbl">Processed</span></div>
    <div class="summary-stat stat-pending"><span class="num">${pending}</span><span class="lbl">Need Mapping</span></div>
  `;
}

// ── Build a single card ───────────────────────────────────
function buildCard(f) {
  const card = document.createElement("div");
  card.className = "file-card";
  card.id = "card-" + CSS.escape(f.filename);

  const badgeClass = {
    ready: "badge-ready",
    needs_mapping: "badge-mapping",
    done: "badge-done",
    error: "badge-error",
  }[f.status] || "badge-mapping";

  const badgeText = {
    ready: f.from_saved_mapping ? "saved mapping" : "auto-mapped",
    needs_mapping: "needs mapping",
    done: "done",
    error: "error",
  }[f.status] || f.status;

  const rowCount = f.row_count != null
    ? `<span class="row-count">${f.row_count} rows</span>` : "";

  // Files that are done or ready can be expanded to review/edit mapping
  const canExpand = f.status === "done" || f.status === "ready";
  const expandBtn = canExpand
    ? `<button class="btn-ghost" style="font-size:12px;padding:4px 10px;"
         onclick="toggleMapping('${f.filename}')">Review Mapping</button>`
    : "";

  card.innerHTML = `
    <div class="file-card-header">
      <span class="file-name" title="${f.filename}">${f.filename}</span>
      ${rowCount}
      ${expandBtn}
      <span class="badge ${badgeClass}">${badgeText}</span>
    </div>
    ${f.status === "needs_mapping" ? buildMappingPanel(f) : ""}
    ${f.status === "error" ? `<div class="mapping-panel"><p style="color:var(--danger)">${f.error || "Unknown error"}</p></div>` : ""}
    <div id="review-${f.filename}" style="display:none">${buildMappingPanel(f, true)}</div>
  `;
  return card;
}

// ── Build mapping panel ───────────────────────────────────
function buildMappingPanel(f, isReview = false) {
  const allFields = {{ field_labels | tojson }};
  const required  = {{ required_fields | tojson }};
  const headers   = f.headers;

  const blankOption = `<option value="">— not in this file —</option>`;
  const selectOne   = `<option value="" disabled selected>— Select One —</option>`;

  const rows = Object.entries(allFields).map(([field, label]) => {
    const isReq = required.includes(field);
    const star = isReq ? `<span class="required-star">*</span>` : "";
    const needsMapping = f.missing_fields.includes(field);
    const existingIdx = f.col_map[field] != null ? f.col_map[field] : "";

    return `
      <span class="field-label">${label}${star}</span>
      <select id="sel-${isReview ? 'review-' : ''}${f.filename}-${field}" data-field="${field}"
              style="${needsMapping ? 'border-color:var(--warn)' : ''}">
        ${isReq && !isReview ? selectOne : blankOption}
        ${headers.map((h, i) =>
          h ? `<option value="${i}" ${existingIdx === i ? "selected" : ""}>${h}</option>` : ""
        ).filter(Boolean).join("")}
      </select>`;
  }).join("");

  const buttonId = isReview ? `review-${f.filename}` : f.filename;
  const buttonLabel = isReview ? "Update Mapping &amp; Reprocess" : "Apply &amp; Process";

  return `
    <div class="mapping-panel">
      <h3>${isReview ? "Review / Edit Column Mapping" : "Column Mapping Required"}</h3>
      <p>${isReview
        ? "Mapping applied to this file. You can adjust any field and reprocess."
        : "These columns could not be matched automatically. Fields marked <span style='color:var(--danger)'>*</span> are required."
      }</p>
      <div class="field-grid">${rows}</div>
      <div class="save-prompt">
        <label>
          <input type="checkbox" id="save-${isReview ? 'review-' : ''}${f.filename}" checked>
          Remember this mapping for files with the same column layout
        </label>
      </div>
      <div class="mapping-actions">
        <button class="btn-success" onclick="applyMapping('${f.filename}', ${isReview})">
          ${buttonLabel}
        </button>
      </div>
    </div>`;
}

// ── Apply manual mapping ──────────────────────────────────
async function applyMapping(filename, isReview = false) {
  const f = files.find(x => x.filename === filename);
  if (!f) return;

  const allFields = {{ field_labels | tojson }};
  const required  = {{ required_fields | tojson }};
  const prefix = isReview ? `review-` : ``;

  const userMap = {};
  for (const field of Object.keys(allFields)) {
    const sel = document.getElementById(`sel-${prefix}${filename}-${field}`);
    if (sel && sel.value !== "") {
      userMap[field] = parseInt(sel.value, 10);
    }
  }

  const missing = required.filter(r => userMap[r] == null);
  if (missing.length) {
    const labels = missing.map(m => allFields[m]).join(", ");
    toast(`Please map required fields: ${labels}`, true);
    return;
  }

  const saveIt = document.getElementById(`save-${prefix}${filename}`).checked;

  const headerMap = {};
  for (const [field, idx] of Object.entries(userMap)) {
    headerMap[field] = f.headers[idx];
  }

  try {
    const result = await api("/api/process", {
      method: "POST",
      body: JSON.stringify({
        filename,
        col_map: userMap,
        header_map: headerMap,
        save_mapping: saveIt,
        headers_signature: f.headers_signature,
      }),
    });

    if (result.error) {
      toast(result.error, true);
      return;
    }

    f.status = "done";
    f.row_count = result.row_count;
    f.col_map = userMap;
    f.from_saved_mapping = false;

    rerenderCard(f);
    updateSummary();
    toast(`✓ ${result.row_count} rows written → ${result.output_filename}`);
  } catch (e) {
    toast("Processing failed: " + e.message, true);
  }
}

// ── Process a ready file ──────────────────────────────────
async function processFile(filename) {
  const f = files.find(x => x.filename === filename);
  if (!f) return;

  try {
    const result = await api("/api/process", {
      method: "POST",
      body: JSON.stringify({ filename, col_map: f.col_map }),
    });

    if (result.error) {
      f.status = "error";
      f.error = result.error;
    } else {
      f.status = "done";
      f.row_count = result.row_count;
    }

    rerenderCard(f);
    updateSummary();
  } catch (e) {
    f.status = "error";
    f.error = e.message;
    rerenderCard(f);
  }
}

// ── Re-render a single card in place ─────────────────────
function rerenderCard(f) {
  const list = document.getElementById("file-list");
  const old = document.getElementById("card-" + CSS.escape(f.filename));
  const newCard = buildCard(f);
  if (old) {
    list.replaceChild(newCard, old);
  } else {
    list.appendChild(newCard);
  }
}
</script>
</body>
</html>
"""


# ============================================================
# ROUTES
# ============================================================

@app.route("/")
def index():
    return render_template_string(
        HTML,
        field_labels=FIELD_LABELS,
        required_fields=REQUIRED_FIELDS,
    )


@app.route("/api/scan")
def api_scan():
    results = scan_input_files()

    global _scan_results
    _scan_results = results

    slim = []
    for r in results:
        slim.append({
            "filename":           r["filename"],
            "status":             r["status"],
            "headers":            r["headers"],
            "col_map":            {k: v for k, v in r.get("col_map", {}).items()
                                   if not k.startswith("_")},
            "missing_fields":     r.get("missing_fields", []),
            "error":              r.get("error", ""),
            "from_saved_mapping": r.get("from_saved_mapping", False),
            "headers_signature":  r.get("headers_signature", ""),
            "row_count":          None,
        })

    return jsonify({"files": slim, "input_dir": str(INPUT_DIR)})


@app.route("/api/process", methods=["POST"])
def api_process():
    data = request.get_json()
    filename     = data.get("filename")
    user_col_map = data.get("col_map", {})
    header_map   = data.get("header_map", {})
    do_save      = data.get("save_mapping", False)
    sig          = data.get("headers_signature", "")

    file_info = next((r for r in _scan_results if r["filename"] == filename), None)
    if not file_info:
        return jsonify({"error": f"File not found in scan results: {filename}"})

    merged = dict(file_info.get("col_map") or {})
    for k, v in user_col_map.items():
        merged[k] = int(v)
    file_info["col_map"] = merged

    try:
        row_count, out_name = process_file_with_map(file_info)
    except Exception as e:
        return jsonify({"error": str(e)})

    if do_save and sig and header_map:
        save_mapping(sig, header_map)

    return jsonify({"row_count": row_count, "output_filename": out_name})


# ============================================================
# LAUNCH
# ============================================================

def open_browser():
    webbrowser.open("http://127.0.0.1:8080")


if __name__ == "__main__":
    threading.Timer(1.2, open_browser).start()
    print("Starting SAP Sales Processor...")
    print("Opening http://127.0.0.1:8080")
    app.run(debug=False, port=8080, host="127.0.0.1")