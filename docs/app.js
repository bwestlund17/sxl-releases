"use strict";

const $ = (id) => document.getElementById(id);
const state = {
  token: null,
  runId: null,
  polling: false,
  // grid: { fileName, fileId?, runId?, sheets: [{name, cells, rows, cols}], active }
  workbook: null,
  selected: "A1",
  pendingFileId: null,
  // staged web-grid edits, keyed by sheet name: Map<address, {value?|formula?}>
  pending: {}
};

// Resolve the API base so the same app works at "/" (local-mode server) and
// under a nested prefix (hosted Supabase Edge Function, served at
// ".../functions/v1/sxl-platform/"). Override order: window.SXL_PLATFORM_URL →
// <meta name="sxl-api-base"> → ?api= → the directory the page is served from.
const API_BASE = (() => {
  const meta = document.querySelector('meta[name="sxl-api-base"]');
  const override = (typeof window !== "undefined" && window.SXL_PLATFORM_URL)
    || (meta && meta.content)
    || new URLSearchParams(location.search).get("api")
    || "";
  if (override) return String(override).replace(/\/+$/, "");
  return location.pathname.replace(/\/[^/]*$/, "");
})();
const api = (path) => `${API_BASE}${String(path).startsWith("/") ? path : `/${path}`}`;

// Supabase Auth config (injected as meta tags on the hosted page). Absent in
// local mode, where the app keeps using the anonymous local token.
const AUTH_URL = ((document.querySelector('meta[name="sxl-auth-url"]') || {}).content || "").replace(/\/+$/, "");
const AUTH_ANON = (document.querySelector('meta[name="sxl-auth-anon-key"]') || {}).content || "";
const AUTH_ENABLED = Boolean(AUTH_URL && AUTH_ANON);
const OAUTH_PROVIDERS = ["google", "github", "azure", "apple", "gitlab", "bitbucket", "discord", "linkedin_oidc"];
const PROVIDER_LABELS = { azure: "Microsoft", linkedin_oidc: "LinkedIn" };

function setStatus(message) {
  $("status").textContent = message;
}

function setAuthStatus(message) {
  $("authStatus").textContent = message;
}

function authFetchRaw(path, options = {}) {
  const headers = Object.assign({ apikey: AUTH_ANON }, options.headers || {});
  return fetch(`${AUTH_URL}${path}`, Object.assign({}, options, { headers }));
}

function readHashSession() {
  const hash = location.hash.startsWith("#") ? location.hash.slice(1) : location.hash;
  if (!hash) return null;
  const params = new URLSearchParams(hash);
  const accessToken = params.get("access_token");
  const error = params.get("error_description") || params.get("error") || "";
  if (!accessToken && !error) return null;
  return { accessToken, error };
}

async function exchangeSession(accessToken) {
  const response = await fetch(api("/login"), {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ accessToken })
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.error || `HTTP ${response.status}`);
  state.token = body.token;
  localStorage.setItem("sxl.platform.token", body.token);
  if (body.username) localStorage.setItem("sxl.platform.username", body.username);
  return body;
}

async function handleAuthRedirect() {
  const session = readHashSession();
  if (!session) return false;
  history.replaceState(null, "", location.pathname + location.search);
  if (session.error) {
    setAuthStatus(`Sign-in failed: ${session.error}`);
    return false;
  }
  try {
    const body = await exchangeSession(session.accessToken);
    setAuthStatus(`Signed in as ${body.username || "verified user"}.`);
    $("authSignOut").hidden = false;
    return true;
  } catch (error) {
    setAuthStatus(`Sign-in failed: ${error.message || error}`);
    return false;
  }
}

function renderProviderButtons(external = {}) {
  const target = $("authProviders");
  target.innerHTML = "";
  const enabled = OAUTH_PROVIDERS.filter((name) => external[name]);
  for (const provider of enabled) {
    const button = document.createElement("button");
    button.textContent = `Continue with ${PROVIDER_LABELS[provider] || provider[0].toUpperCase() + provider.slice(1)}`;
    button.className = "secondary";
    button.style.marginRight = "8px";
    button.addEventListener("click", () => {
      const redirectTo = location.origin + location.pathname;
      location.href = `${AUTH_URL}/auth/v1/authorize?provider=${encodeURIComponent(provider)}&redirect_to=${encodeURIComponent(redirectTo)}`;
    });
    target.appendChild(button);
  }
  if (!enabled.length) {
    const note = document.createElement("span");
    note.className = "muted";
    note.textContent = "Email is the enabled sign-in method; OAuth providers appear here once enabled in Supabase.";
    target.appendChild(note);
  }
}

async function setupAuth() {
  $("authCard").hidden = false;
  const storedUser = localStorage.getItem("sxl.platform.username");
  if (storedUser) {
    setAuthStatus(`Signed in as ${storedUser}.`);
    $("authSignOut").hidden = false;
  }
  try {
    const response = await authFetchRaw("/auth/v1/settings");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const settings = await response.json();
    renderProviderButtons(settings.external || {});
  } catch (error) {
    setAuthStatus(`Auth settings unavailable (${error.message || error}).`);
  }
  $("authEmailBtn").addEventListener("click", async () => {
    const email = $("authEmail").value.trim();
    if (!email) {
      setAuthStatus("Enter your email first.");
      return;
    }
    $("authEmailBtn").disabled = true;
    try {
      const redirectTo = location.origin + location.pathname;
      const response = await authFetchRaw(`/auth/v1/otp?redirect_to=${encodeURIComponent(redirectTo)}`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ email, create_user: true })
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      setAuthStatus(`Sign-in link sent to ${email}. Use the link, or enter the 6-digit code below.`);
    } catch (error) {
      setAuthStatus(`Could not send link: ${error.message || error}`);
    } finally {
      $("authEmailBtn").disabled = false;
    }
  });
  // Six-digit email OTP: verifies without a redirect, so it works even before
  // the Pages URL is added to Supabase's redirect allowlist.
  $("authOtpBtn").addEventListener("click", async () => {
    const email = $("authEmail").value.trim();
    const token = $("authOtp").value.replace(/\D/g, "");
    if (!email || !token) {
      setAuthStatus("Enter your email and the 6-digit code first.");
      return;
    }
    $("authOtpBtn").disabled = true;
    try {
      const response = await authFetchRaw("/auth/v1/verify", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ type: "email", email, token })
      });
      const session = await response.json().catch(() => ({}));
      if (!response.ok || !session.access_token) {
        throw new Error(session.error_description || session.msg || `HTTP ${response.status}`);
      }
      const body = await exchangeSession(session.access_token);
      setAuthStatus(`Signed in as ${body.username || email}.`);
      $("authSignOut").hidden = false;
    } catch (error) {
      setAuthStatus(`Could not verify code: ${error.message || error}`);
    } finally {
      $("authOtpBtn").disabled = false;
    }
  });
  $("authSignOut").addEventListener("click", () => {
    state.token = null;
    localStorage.removeItem("sxl.platform.token");
    localStorage.removeItem("sxl.platform.username");
    location.reload();
  });
}

async function login() {
  const response = await fetch(api("/login"), {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ username: "local-web" })
  });
  if (!response.ok) throw new Error(`login failed: HTTP ${response.status}`);
  const body = await response.json();
  state.token = body.token;
  localStorage.setItem("sxl.platform.token", body.token);
  return body.token;
}

async function ensureToken() {
  if (state.token) return state.token;
  const stored = localStorage.getItem("sxl.platform.token");
  if (stored) {
    state.token = stored;
    return stored;
  }
  return login();
}

async function authFetch(path, options = {}) {
  const token = await ensureToken();
  const headers = Object.assign({}, options.headers || {}, { authorization: `Bearer ${token}` });
  let response = await fetch(api(path), Object.assign({}, options, { headers }));
  if (response.status === 401) {
    await login();
    response = await fetch(api(path), Object.assign({}, options, {
      headers: Object.assign({}, options.headers || {}, { authorization: `Bearer ${state.token}` })
    }));
  }
  return response;
}

async function uploadFile(file) {
  const form = new FormData();
  form.append("file", file);
  const response = await authFetch("/api/spreadsheets/upload", { method: "POST", body: form });
  if (!response.ok) throw new Error(`upload failed: HTTP ${response.status}`);
  return response.json();
}

// ---- Grid ------------------------------------------------------------------
function columnToLetters(index) {
  let letters = "";
  while (index > 0) {
    const rem = (index - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    index = Math.floor((index - 1) / 26);
  }
  return letters;
}

function emptyWorkbook(name) {
  return {
    fileName: name,
    fileId: null,
    runId: null,
    sheets: [{ name: "Sheet1", cells: {}, rows: 1, cols: 1 }],
    active: 0
  };
}

function setWorkbook(workbook) {
  state.workbook = workbook;
  state.selected = "A1";
  state.pending = {};
  updatePendingBar();
  $("workbookTitle").textContent = workbook.fileName || "untitled";
  renderSheetTabs();
  renderGrid();
}

function activePending() {
  const wb = state.workbook;
  if (!wb) return new Map();
  const name = wb.sheets[wb.active].name;
  if (!state.pending[name]) state.pending[name] = new Map();
  return state.pending[name];
}

function stageEdit(address, edit) {
  const pending = activePending();
  pending.set(address, edit);
  updatePendingBar();
  renderGrid();
}

function discardEdits() {
  const wb = state.workbook;
  if (!wb) return;
  delete state.pending[wb.sheets[wb.active].name];
  updatePendingBar();
  renderGrid();
}

function updatePendingBar() {
  const wb = state.workbook;
  if (!wb) { $("pendingWrap").hidden = true; return; }
  const activeName = wb.sheets[wb.active].name;
  const activePending = state.pending[activeName];
  const count = activePending ? activePending.size : 0;
  const elsewhere = Object.keys(state.pending)
    .filter((name) => name !== activeName && state.pending[name] && state.pending[name].size > 0)
    .reduce((sum, name) => sum + state.pending[name].size, 0);
  $("pendingWrap").hidden = count === 0 && elsewhere === 0;
  $("pendingCount").textContent =
    `${count} staged edit${count === 1 ? "" : "s"}` + (elsewhere > 0 ? ` (+${elsewhere} on other sheets)` : "");
}

function renderSheetTabs() {
  const target = $("sheetTabs");
  target.innerHTML = "";
  const wb = state.workbook;
  if (!wb) return;
  wb.sheets.slice(0, 15).forEach((sheet, index) => {
    const tab = document.createElement("button");
    tab.className = `sheet-tab${index === wb.active ? " active" : ""}`;
    tab.textContent = sheet.name;
    tab.addEventListener("click", () => {
      wb.active = index;
      state.selected = "A1";
      renderSheetTabs();
      renderGrid();
    });
    target.appendChild(tab);
  });
}

const GRID_MAX_ROWS = 200;
const GRID_MAX_COLS = 40;

function renderGrid() {
  commitActiveEditor();
  const table = $("grid");
  table.innerHTML = "";
  const wb = state.workbook;
  if (!wb) return;
  const sheet = wb.sheets[wb.active] || wb.sheets[0];
  const pending = state.pending[sheet.name];
  // Always render a comfortable viewport, up to the preview caps.
  const rows = Math.max(40, Math.min(sheet.rows || 1, GRID_MAX_ROWS));
  const cols = Math.max(12, Math.min(sheet.cols || 1, GRID_MAX_COLS));
  state.rendered = { rows, cols };

  // Formula preview resolver: staged edits override stored cells; memoized,
  // cycle-guarded, whitelist of typical Excel functions (see engine below).
  const formulaMemo = new Map();
  const formulaVisiting = new Set();
  const resolveFormulaCell = (address) => {
    if (formulaMemo.has(address)) return formulaMemo.get(address);
    if (formulaVisiting.has(address)) return formulaError(FORMULA_ERRORS.circ);
    formulaVisiting.add(address);
    let value = 0;
    try {
      const stagedEdit = pending && pending.get(address);
      const stored = (sheet.cells || {})[address];
      if (stagedEdit) {
        value = stagedEdit.formula ? computeStagedFormula(stagedEdit.formula) : stagedEdit.value;
      } else if (stored) {
        if (stored.v !== undefined && stored.v !== null) value = stored.v;
        else if (stored.f) value = computeStagedFormula(stored.f);
        else value = 0;
      }
    } catch (error) {
      value = isFormulaError(error) ? error : formulaError(FORMULA_ERRORS.value);
    }
    formulaVisiting.delete(address);
    formulaMemo.set(address, value);
    return value;
  };
  const computeStagedFormula = (formula) =>
    evalFormulaNode(parseFormulaTokens(tokenizeFormula(formula.replace(/^=/, ""))), resolveFormulaCell);

  const head = document.createElement("tr");
  const corner = document.createElement("th");
  corner.className = "corner rowhead";
  head.appendChild(corner);
  for (let c = 1; c <= cols; c++) {
    const th = document.createElement("th");
    th.textContent = columnToLetters(c);
    head.appendChild(th);
  }
  table.appendChild(head);

  for (let r = 1; r <= rows; r++) {
    const tr = document.createElement("tr");
    const rowhead = document.createElement("th");
    rowhead.className = "rowhead";
    rowhead.textContent = r;
    tr.appendChild(rowhead);
    for (let c = 1; c <= cols; c++) {
      const address = `${columnToLetters(c)}${r}`;
      const td = document.createElement("td");
      td.dataset.addr = address;
      const pendingEdit = pending && pending.get(address);
      const cell = pendingEdit
        ? (pendingEdit.formula
            ? { f: pendingEdit.formula, ...(pendingEdit.format ? { nf: pendingEdit.format } : {}) }
            : { v: pendingEdit.value, ...(pendingEdit.format ? { nf: pendingEdit.format } : {}) })
        : sheet.cells[address];
      if (cell) {
        if (cell.v !== undefined && cell.v !== null) {
          td.textContent = typeof cell.nf === "string" && typeof cell.v === "number"
            ? renderExcelNumber(cell.nf, cell.v)
            : String(cell.v);
          if (typeof cell.v === "number") td.classList.add("num");
        } else if (cell.f) {
          // Formula preview: compute typical formulas client-side; unsupported
          // ones (#NAME?) render as literal text instead of a fake value.
          const computed = computeStagedFormula(cell.f);
          if (isFormulaError(computed) && computed.__err === FORMULA_ERRORS.name) {
            td.textContent = cell.f;
            td.classList.add("fonly");
          } else if (isFormulaError(computed)) {
            td.textContent = computed.__err;
            td.classList.add("fonly");
          } else if (typeof computed === "number" && typeof cell.nf === "string") {
            td.textContent = renderExcelNumber(cell.nf, computed);
            td.classList.add("num");
          } else {
            td.textContent = formulaDisplay(computed);
            if (typeof computed === "number") td.classList.add("num");
          }
        }
      }
      if (pendingEdit) td.classList.add("pending");
      if (address === state.selected) td.classList.add("sel");
      td.addEventListener("click", () => {
        // Excel-like: clicking a cell moves keyboard focus to the grid so
        // arrows/type-to-edit work immediately.
        $("gridScroll").focus({ preventScroll: true });
        selectCell(address, effectiveCell(sheet, address));
      });
      td.addEventListener("dblclick", () => beginCellEdit(address));
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  const shown = Object.keys(sheet.cells || {}).length;
  $("gridStatus").textContent = `${sheet.name}: ${shown} cell${shown === 1 ? "" : "s"} loaded`
    + (shown >= 20000 ? " (preview capped)" : "");
  selectCell(state.selected, effectiveCell(sheet, state.selected));
}

// Pending edits override the stored cell for display and the formula bar.
function effectiveCell(sheet, address) {
  const pending = state.pending[sheet.name];
  const staged = pending && pending.get(address);
  if (staged) return staged.formula ? { f: staged.formula } : { v: staged.value };
  return (sheet.cells || {})[address];
}

// Excel-style keyboard navigation across the rendered viewport.
function addressParts(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(address || "");
  return match ? { letters: match[1], row: Number(match[2]) } : null;
}
function colToIndex(letters) {
  return letters.split("").reduce((sum, ch) => sum * 26 + (ch.charCodeAt(0) - 64), 0);
}
function moveSelection(dr, dc) {
  commitActiveEditor();
  const current = addressParts(state.selected);
  const rendered = state.rendered || { rows: 40, cols: 12 };
  if (!current) return;
  const col = Math.min(rendered.cols, Math.max(1, colToIndex(current.letters) + dc));
  const row = Math.min(rendered.rows, Math.max(1, current.row + dr));
  const address = `${columnToLetters(col)}${row}`;
  const sheet = state.workbook && state.workbook.sheets[state.workbook.active];
  selectCell(address, sheet ? effectiveCell(sheet, address) : null);
  const td = $("grid").querySelector(`td[data-addr="${address}"]`);
  if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
}

// Cell editing. Three entries: double-click / F2 edit the existing content,
// typing a printable character REPLACES it with that character (Excel).
// Enter or blur commits into the staged edits; Escape cancels.
//
// Commit does NOT rely on blur alone: some embedding contexts never dispatch
// blur/focusout (window without OS focus), so the editor is tracked and
// commitActiveEditor() runs before every navigation/render path.
let activeEditor = null;
function commitActiveEditor() {
  const editor = activeEditor;
  if (editor) editor.commit();
}

function beginCellEdit(address, initial) {
  commitActiveEditor();
  const td = $("grid").querySelector(`td[data-addr="${address}"]`);
  const sheet = state.workbook && state.workbook.sheets[state.workbook.active];
  if (!td || !sheet) return;
  const current = effectiveCell(sheet, address);
  const existing = current ? (current.f || (current.v !== undefined && current.v !== null ? String(current.v) : "")) : "";
  const text = initial !== undefined ? initial : existing;
  td.textContent = "";
  const input = document.createElement("input");
  input.className = "cell-editor";
  input.value = text;
  td.appendChild(input);
  input.focus();
  input.setSelectionRange(text.length, text.length);
  let cancelled = false;
  let done = false;
  const commit = () => {
    const stale = done || cancelled || !activeEditor || activeEditor.input !== input;
    activeEditor = null;
    if (stale) return;
    done = true;
    const value = input.value.trim();
    if (value === existing) {
      renderGrid();
      return;
    }
    stageEdit(address, value.startsWith("=") ? { formula: value } : { value });
    selectCell(address, effectiveCell(sheet, address));
  };
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      commit();
      moveSelection(1, 0);
      $("gridScroll").focus({ preventScroll: true });
    } else if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      cancelled = true;
      activeEditor = null;
      renderGrid();
      $("gridScroll").focus({ preventScroll: true });
    } else if (event.key.startsWith("Arrow") || event.key === "Tab") {
      // Keep the caret in the editor; do not navigate the grid.
      event.stopPropagation();
    }
  });
  // Blur commit still works for real focused windows; the tracked paths
  // above cover contexts where the browser never dispatches blur.
  input.addEventListener("blur", commit);
  activeEditor = { input, commit };
}

// Grid keyboard routing: navigation + type-to-edit when the grid has focus
// and no cell editor input is active.
function onGridKey(event) {
  const tag = (event.target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select") return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  switch (event.key) {
    case "ArrowUp": event.preventDefault(); return moveSelection(-1, 0);
    case "ArrowDown": event.preventDefault(); return moveSelection(1, 0);
    case "ArrowLeft": event.preventDefault(); return moveSelection(0, -1);
    case "ArrowRight": event.preventDefault(); return moveSelection(0, 1);
    case "Enter": event.preventDefault(); return moveSelection(1, 0);
    case "Tab": event.preventDefault(); return moveSelection(0, event.shiftKey ? -1 : 1);
    case "F2": event.preventDefault(); return beginCellEdit(state.selected);
    case "Delete":
    case "Backspace":
      event.preventDefault();
      return stageEdit(state.selected, { value: "" });
    default:
      if (event.key.length === 1) {
        event.preventDefault();
        beginCellEdit(state.selected, event.key);
      }
  }
}

// Formula-bar editing: Enter stages, Esc restores the current value.
function formulaBarCommit() {
  const sheet = state.workbook && state.workbook.sheets[state.workbook.active];
  if (!sheet) return;
  const text = $("formulaBar").value.trim();
  const current = effectiveCell(sheet, state.selected);
  const currentText = current ? (current.f || (current.v !== undefined && current.v !== null ? String(current.v) : "")) : "";
  if (text === currentText) return;
  stageEdit(state.selected, text.startsWith("=") ? { formula: text } : { value: text });
}

function selectCell(address, cell) {
  state.selected = address;
  $("nameBox").textContent = address;
  $("formulaBar").value = cell ? (cell.f || (cell.v !== undefined && cell.v !== null ? String(cell.v) : "")) : "";
  for (const td of $("grid").querySelectorAll("td.sel")) td.classList.remove("sel");
  const target = $("grid").querySelector(`td[data-addr="${address}"]`);
  if (target) target.classList.add("sel");
}
async function openWorkbookFile(file) {
  try {
    setStatus("Uploading file...");
    const uploaded = await uploadFile(file);
    state.pendingFileId = uploaded.fileId;
    await loadWorkbookFromFileId(uploaded.fileId, uploaded.filename || file.name);
    setStatus(`Loaded ${file.name}.`);
  } catch (error) {
    setStatus(`ERROR: ${error.message || error}`);
  }
}

async function loadWorkbookFromFileId(fileId, filename) {
  const response = await authFetch(`/api/spreadsheets/workbook/${encodeURIComponent(fileId)}/sheet-data`);
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    setStatus(body.error || `no grid preview for ${filename} (HTTP ${response.status})`);
    if (!state.workbook) setWorkbook(emptyWorkbook(filename));
    return;
  }
  setWorkbook({
    fileName: filename || body.filename || "workbook.xlsx",
    fileId,
    runId: null,
    sheets: body.sheets,
    active: 0,
    truncated: body.truncated
  });
}

async function loadWorkbookFromRun(runId, artifact) {
  const query = artifact ? `?artifact=${encodeURIComponent(artifact)}` : "";
  const response = await authFetch(`/api/spreadsheets/${encodeURIComponent(runId)}/sheet-data${query}`);
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    setStatus(body.error || `no result workbook for run (HTTP ${response.status})`);
    return false;
  }
  setWorkbook({
    fileName: body.filename || "result.xlsx",
    fileId: null,
    runId,
    sheets: body.sheets,
    active: 0,
    truncated: body.truncated
  });
  return true;
}

async function exportWorkbook() {
  const wb = state.workbook;
  if (!wb) {
    setStatus("Nothing to export yet.");
    return;
  }
  let url;
  if (wb.runId) url = `/api/spreadsheets/${encodeURIComponent(wb.runId)}/download`;
  else if (wb.fileId) url = `/api/spreadsheets/workbook/${encodeURIComponent(wb.fileId)}/download`;
  if (!url) {
    setStatus("This sheet has no backing file to export.");
    return;
  }
  const response = await authFetch(url);
  if (!response.ok) {
    setStatus(`export failed: HTTP ${response.status}`);
    return;
  }
  const blob = await response.blob();
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = wb.fileName || "workbook.xlsx";
  anchor.click();
  URL.revokeObjectURL(anchor.href);
}

// ---- Chat transcript -------------------------------------------------------
function addMessage(kind, text) {
  const entry = document.createElement("div");
  entry.className = `msg ${kind}`;
  entry.textContent = text;
  $("transcript").appendChild(entry);
  $("transcript").scrollTop = $("transcript").scrollHeight;
  return entry;
}

function addDownloadButtons(target, run) {
  const wrap = document.createElement("div");
  wrap.className = "dl";
  const add = (label, url) => {
    const button = document.createElement("button");
    button.className = "secondary";
    button.textContent = label;
    button.addEventListener("click", () => downloadUrl(url));
    wrap.appendChild(button);
  };
  if (run.downloadUrl) add("Download workbook", api(run.downloadUrl));
  for (const artifact of run.artifacts || []) {
    add(`Download ${artifact}`, api(`/api/spreadsheets/${run.runId}/artifacts/${encodeURIComponent(artifact)}`));
  }
  if (wrap.children.length) target.appendChild(wrap);
}

async function downloadUrl(url) {
  const response = await authFetch(url);
  if (!response.ok) {
    setStatus(`download failed: HTTP ${response.status}`);
    return;
  }
  const blob = await response.blob();
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = (url.split("/").pop() || "download").replace(/[?].*$/, "");
  anchor.click();
  URL.revokeObjectURL(anchor.href);
}

async function pollRun(runId, statusMessage) {
  state.polling = true;
  let last = null;
  while (state.polling && state.runId === runId) {
    const response = await authFetch(`/api/spreadsheets/${runId}`);
    if (!response.ok) {
      statusMessage(`status failed: HTTP ${response.status}`);
      state.polling = false;
      return null;
    }
    const run = await response.json();
    if (run.status !== last) {
      last = run.status;
      statusMessage(`Run ${runId.slice(0, 8)}: ${run.status}…`);
    }
    if (["completed", "failed", "cancelled"].includes(run.status)) {
      state.polling = false;
      return run;
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  state.polling = false;
  return null;
}

async function loadRuns() {
  try {
    const response = await authFetch("/api/spreadsheets/runs");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const body = await response.json();
    const list = $("runs");
    list.hidden = false;
    list.innerHTML = "";
    for (const run of (body.runs || []).slice(0, 10)) {
      const item = document.createElement("li");
      const link = document.createElement("a");
      link.textContent = `${run.runId.slice(0, 8)} · ${run.status} · ${String(run.prompt || "").slice(0, 40)}`;
      link.addEventListener("click", async () => {
        const statusResponse = await authFetch(`/api/spreadsheets/${run.runId}`);
        if (statusResponse.ok) {
          const run = await statusResponse.json();
          addMessage("sys", `Loaded run ${run.runId.slice(0, 8)} (${run.status}).`);
          if (run.status === "completed") await loadWorkbookFromRun(run.runId);
        }
      });
      item.appendChild(link);
      list.appendChild(item);
    }
    if (!list.children.length) {
      const item = document.createElement("li");
      item.textContent = "No runs yet.";
      list.appendChild(item);
    }
  } catch (error) {
    setStatus(`Could not load runs: ${error.message || error}`);
  }
}

async function submitRun(promptText, overrides = {}) {
  const prompt = (promptText !== undefined ? promptText : $("prompt").value).trim();
  if (!prompt && !overrides.edits) {
    setStatus("Enter a prompt first.");
    return;
  }
  $("send").disabled = true;
  if (prompt) addMessage("user", prompt);
  if (promptText !== undefined && !overrides.edits) $("prompt").value = "";
  const statusMessage = addMessage("sys", "Submitting…");
  try {
    await ensureToken();
    let initFile = overrides.initFile || null;
    const file = $("file").files && $("file").files[0];
    if (file) {
      statusMessage.textContent = "Uploading attachment…";
      const uploaded = await uploadFile(file);
      initFile = uploaded.fileId;
      $("file").value = "";
    } else if (!initFile && state.pendingFileId) {
      initFile = state.pendingFileId;
    }
    const body = {
      prompt,
      mode: overrides.mode || ($("mode").value === "ask" ? "ask" : "action")
    };
    if (initFile) body.initFile = initFile;
    if (overrides.edits) {
      body.edits = overrides.edits;
      if (overrides.sheet) body.sheet = overrides.sheet;
    }
    if (!overrides.edits && $("model").value) body.model = $("model").value;
    statusMessage.textContent = "Queued…";
    const response = await authFetch("/api/spreadsheets", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(body)
    });
    const submitted = await response.json();
    if (!response.ok) throw new Error(submitted.error || `HTTP ${response.status}`);
    state.runId = submitted.runId;
    const run = await pollRun(submitted.runId, (m) => { statusMessage.textContent = m; });
    if (!run) return null;
    if (run.status === "completed") {
      statusMessage.className = "msg done";
      statusMessage.textContent = run.summary ? String(run.summary).slice(0, 4000) : `Run ${run.runId.slice(0, 8)} completed.`;
      addDownloadButtons(statusMessage, run);
      const loaded = await loadWorkbookFromRun(run.runId);
      if (loaded) addMessage("sys", "Result workbook loaded into the grid.");
    } else {
      statusMessage.className = "msg err";
      statusMessage.textContent = `Run ${run.status}: ${run.error || "no details"}`;
    }
    loadCredits();
    loadRuns();
    return run;
  } catch (error) {
    statusMessage.className = "msg err";
    statusMessage.textContent = `ERROR: ${error.message || error}`;
    return null;
  } finally {
    $("send").disabled = false;
  }
}

// Apply staged grid edits: the audited deterministic --set path. The CLI is
// single-sheet per run, so staged sheets apply as sequential chained runs
// (active sheet first): each result becomes the next run's input via
// from-run promotion, and every run is its own revertible audit record.
async function applyEdits() {
  const wb = state.workbook;
  if (!wb) return;
  const activeName = wb.sheets[wb.active].name;
  const sheetNames = Object.keys(state.pending)
    .filter((name) => state.pending[name] && state.pending[name].size > 0)
    .sort((a, b) => (a === activeName ? -1 : b === activeName ? 1 : 0));
  if (!sheetNames.length) return;
  for (const sheetName of sheetNames) {
    const pending = state.pending[sheetName];
    if (!pending || pending.size === 0) continue;
    let fileId = state.workbook && state.workbook.fileId;
    if (!fileId && state.workbook && state.workbook.runId) {
      setStatus(`Preparing result workbook for ${sheetName}…`);
      try {
        const response = await authFetch("/api/spreadsheets/workbook/from-run", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ runId: state.workbook.runId })
        });
        const body = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(body.error || `HTTP ${response.status}`);
        fileId = body.fileId;
      } catch (error) {
        setStatus(`Could not chain from the result: ${error.message || error}`);
        return;
      }
    }
    if (!fileId) {
      setStatus("Open a workbook first — an empty sheet has nothing to edit.");
      return;
    }
    const edits = [...pending.entries()].map(([address, edit]) => ({
      address,
      ...(edit.formula ? { formula: edit.formula } : { value: edit.value }),
      ...(edit.format ? { format: edit.format } : {}),
      ...(edit.bold !== undefined ? { bold: edit.bold } : {}),
      ...(edit.fillColor ? { fillColor: edit.fillColor } : {}),
      ...(edit.fontColor ? { fontColor: edit.fontColor } : {})
    }));
    const run = await submitRun("", {
      edits,
      sheet: sheetName,
      initFile: fileId,
      mode: "action"
    });
    // On failure: stop the chain and keep this sheet's edits staged.
    if (!run || run.status !== "completed") return;
    // Success: submitRun reloaded the result into state.workbook; drop this
    // sheet's staged edits and let the next iteration chain from the new run.
    delete state.pending[sheetName];
    updatePendingBar();
  }
}

// ---- Credit billing (Stripe Checkout) -------------------------------------
function setBillingStatus(message) {
  $("billingStatus").textContent = message;
}

async function buyPack(packId) {
  setBillingStatus("Starting checkout…");
  try {
    const response = await authFetch("/v1/billing/checkout", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ packId })
    });
    const body = await response.json().catch(() => ({}));
    if (response.status === 503) throw new Error("billing is not enabled on this deployment yet");
    if (response.status === 403) throw new Error("sign in with a verified account (Google/email) to buy credits");
    if (!response.ok || !body.url) throw new Error(body.error || `HTTP ${response.status}`);
    location.href = body.url;
  } catch (error) {
    setBillingStatus(`Checkout unavailable: ${error.message || error}`);
  }
}

async function loadBilling() {
  let body;
  try {
    const response = await fetch(api("/v1/billing/packs"));
    if (!response.ok) return; // local mode / older server has no billing surface
    body = await response.json();
  } catch {
    return;
  }
  $("billingCard").hidden = false;
  const target = $("billingPacks");
  target.innerHTML = "";
  for (const pack of body.packs || []) {
    const button = document.createElement("button");
    button.textContent = `${pack.label} — $${(Number(pack.amountCents) / 100).toFixed(2)}`;
    button.disabled = !body.configured;
    button.style.marginRight = "8px";
    button.addEventListener("click", () => buyPack(pack.id));
    target.appendChild(button);
  }
  const outcome = new URLSearchParams(location.search).get("billing");
  if (outcome === "success") setBillingStatus("Payment received — credits will appear shortly.");
  else if (outcome === "cancel") setBillingStatus("Checkout cancelled.");
  else setBillingStatus(body.configured ? "Credits are added after payment (Stripe)." : "Billing is not enabled on this deployment yet.");
}

async function loadCredits() {
  try {
    const response = await authFetch("/v1/credits/balance");
    if (!response.ok) return;
    const body = await response.json();
    $("creditsChip").textContent = `${body.balance} credits`;
  } catch {
    // chip keeps its placeholder
  }
}

async function loadModels() {
  try {
    const response = await authFetch("/v1/models");
    if (!response.ok) return;
    const body = await response.json();
    const select = $("model");
    select.innerHTML = "";
    for (const model of body.models || []) {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.id;
      select.appendChild(option);
    }
  } catch {
    // run submits without a model field when empty
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  $("openFile").addEventListener("click", () => $("file").click());
  $("file").addEventListener("change", () => {
    const file = $("file").files && $("file").files[0];
    if (file) openWorkbookFile(file);
  });
  $("newFile").addEventListener("click", () => {
    state.pendingFileId = null;
    $("file").value = "";
    setWorkbook(emptyWorkbook("empty-sheet.xlsx"));
    setStatus("New empty sheet. Attach or Open a file to work on real workbooks.");
  });
  $("exportFile").addEventListener("click", exportWorkbook);
  $("send").addEventListener("click", () => submitRun());
  $("prompt").addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      submitRun();
    }
  });
  $("attach").addEventListener("click", () => $("file").click());
  $("accountBtn").addEventListener("click", () => {
    $("accountPanel").hidden = !$("accountPanel").hidden;
  });
  $("runsLink").addEventListener("click", loadRuns);
  $("applyEdits").addEventListener("click", applyEdits);
  $("discardEdits").addEventListener("click", discardEdits);
  $("refreshGrid").addEventListener("click", async () => {
    const wb = state.workbook;
    if (!wb) return;
    if (wb.runId) await loadWorkbookFromRun(wb.runId);
    else if (wb.fileId) await loadWorkbookFromFileId(wb.fileId, wb.fileName);
    else setStatus("Nothing to refresh yet.");
  });
  $("gridScroll").addEventListener("keydown", onGridKey);
  // Number-format presets: stage {value|formula, format} on the selected cell.
  for (const button of document.querySelectorAll(".fmt-btn")) {
    button.addEventListener("click", () => {
      const wb = state.workbook;
      if (!wb) return;
      const sheet = wb.sheets[wb.active];
      const address = state.selected;
      const current = effectiveCell(sheet, address);
      if (!current || (current.v === undefined && !current.f)) {
        setStatus("Select a cell with a value first.");
        return;
      }
      const staged = current.f ? { formula: current.f, format: button.dataset.format } : { value: String(current.v), format: button.dataset.format };
      stageEdit(address, staged);
      setStatus(`Staged ${button.dataset.format} on ${address} — Apply runs it through the audited ledger.`);
    });
  }
  // Character/cell style presets (bold, fill, font) — same staged-edit flow.
  const stageStyle = (patch) => {
    const wb = state.workbook;
    if (!wb) return;
    const sheet = wb.sheets[wb.active];
    const address = state.selected;
    const current = effectiveCell(sheet, address);
    if (!current || (current.v === undefined && !current.f)) {
      setStatus("Select a cell with a value first.");
      return;
    }
    const base = current.f ? { formula: current.f } : { value: String(current.v) };
    stageEdit(address, { ...base, ...patch });
    setStatus(`Staged style on ${address} — Apply runs it through the audited ledger.`);
  };
  for (const button of document.querySelectorAll("#styleButtons .style-btn")) {
    button.addEventListener("click", () => {
      if (button.dataset.bold) {
        // Toggle against the currently staged bold state.
        const sheet = state.workbook.sheets[state.workbook.active];
        const current = effectiveCell(sheet, state.selected);
        const nowBold = Boolean(current && current.bold);
        stageStyle({ bold: !nowBold });
      } else if (button.dataset.fill) {
        stageStyle({ fillColor: button.dataset.fill });
      } else if (button.dataset.font) {
        stageStyle({ fontColor: button.dataset.font });
      }
    });
  }
  $("formulaBar").addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      formulaBarCommit();
      $("formulaBar").blur();
    } else if (event.key === "Escape") {
      const sheet = state.workbook && state.workbook.sheets[state.workbook.active];
      $("formulaBar").value = sheet ? (() => {
        const current = effectiveCell(sheet, state.selected);
        return current ? (current.f || (current.v !== undefined && current.v !== null ? String(current.v) : "")) : "";
      })() : "";
      $("formulaBar").blur();
    }
  });
  for (const chip of document.querySelectorAll(".chip-prompt")) {
    chip.addEventListener("click", () => submitRun(chip.dataset.prompt));
  }
  setWorkbook(emptyWorkbook("empty-sheet.xlsx"));
  if (AUTH_ENABLED) {
    await setupAuth();
    await handleAuthRedirect();
  }
  loadCredits();
  loadModels();
  loadBilling();
  loadRuns();
});

// ---- Formula preview engine (pure, no DOM) --------------------------------
// The grid computes TYPICAL Excel formulas client-side so staged edits show
// real values instead of literal "=..." text (Apply still recomputes in real
// Excel, which stays the source of truth). Whitelisted functions only; an
// unsupported formula renders as its literal text rather than a fake value.

const FORMULA_ERRORS = {
  div0: "#DIV/0!", value: "#VALUE!", name: "#NAME?", ref: "#REF!", circ: "#CIRC!", num: "#NUM!"
};
const FORMULA_MAX_RANGE_CELLS = 10_000;

function formulaError(text) { return { __err: text }; }
function isFormulaError(value) { return value !== null && typeof value === "object" && "__err" in value; }

function formulaToNumber(value) {
  if (isFormulaError(value)) return value;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "string") {
    const num = Number(value);
    return value.trim() !== "" && !Number.isNaN(num) ? num : formulaError(FORMULA_ERRORS.value);
  }
  return formulaError(FORMULA_ERRORS.value);
}

function formulaCompare(a, b) {
  if (typeof a === "number" && typeof b === "number") return a === b ? 0 : (a < b ? -1 : 1);
  const rank = (v) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  if (rank(a) !== rank(b)) return rank(a) < rank(b) ? -1 : 1;
  if (typeof a === "string") return a.toLowerCase() === b.toLowerCase() ? 0 : (a.toLowerCase() < b.toLowerCase() ? -1 : 1);
  return a === b ? 0 : (a < b ? -1 : 1);
}

function tokenizeFormula(text) {
  const tokens = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === " ") { i++; continue; }
    if (/[0-9.]/.test(ch)) {
      const match = /^[0-9]*\.?[0-9]+([eE][+-]?[0-9]+)?/.exec(text.slice(i));
      tokens.push({ type: "num", value: Number(match[0]) });
      i += match[0].length;
    } else if (ch === '"') {
      const end = text.indexOf('"', i + 1);
      if (end === -1) throw formulaError(FORMULA_ERRORS.value);
      tokens.push({ type: "str", value: text.slice(i + 1, end) });
      i = end + 1;
    } else if (/[A-Za-z_$]/.test(ch)) {
      const match = /^[A-Za-z_$][A-Za-z0-9_$.]*/.exec(text.slice(i));
      tokens.push({ type: "ident", value: match[0].replace(/\$/g, "").toUpperCase() });
      i += match[0].length;
    } else if ("+-*/^&%(),:<>=".includes(ch)) {
      if (ch === "<" && text[i + 1] === ">") { tokens.push({ type: "op", value: "<>" }); i += 2; continue; }
      if (ch === "<" && text[i + 1] === "=") { tokens.push({ type: "op", value: "<=" }); i += 2; continue; }
      if (ch === ">" && text[i + 1] === "=") { tokens.push({ type: "op", value: ">=" }); i += 2; continue; }
      tokens.push({ type: ch === ":" || ch === "," || ch === "(" || ch === ")" ? "punct" : "op", value: ch });
      i++;
    } else {
      throw formulaError(FORMULA_ERRORS.value);
    }
  }
  return tokens;
}

// Recursive-descent parse → AST. Whitespace-insensitive; precedence per Excel
// (comparison < concat < +- < */ < ^ < unary < percent).
function parseFormulaTokens(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const take = () => tokens[pos++];
  function parseExpr() { return parseCompare(); }
  function parseCompare() {
    let left = parseConcat();
    while (peek() && peek().type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().value)) {
      const op = take().value;
      left = { type: "binary", op, left, right: parseConcat() };
    }
    return left;
  }
  function parseConcat() {
    let left = parseAdd();
    while (peek() && peek().type === "op" && peek().value === "&") {
      take();
      left = { type: "binary", op: "&", left, right: parseAdd() };
    }
    return left;
  }
  function parseAdd() {
    let left = parseMul();
    while (peek() && peek().type === "op" && (peek().value === "+" || peek().value === "-")) {
      const op = take().value;
      left = { type: "binary", op, left, right: parseMul() };
    }
    return left;
  }
  function parseMul() {
    let left = parsePow();
    while (peek() && peek().type === "op" && (peek().value === "*" || peek().value === "/")) {
      const op = take().value;
      left = { type: "binary", op, left, right: parsePow() };
    }
    return left;
  }
  function parsePow() {
    let left = parseUnary();
    while (peek() && peek().type === "op" && peek().value === "^") {
      take();
      left = { type: "binary", op: "^", left, right: parseUnary() };
    }
    return left;
  }
  function parseUnary() {
    if (peek() && peek().type === "op" && (peek().value === "-" || peek().value === "+")) {
      const op = take().value;
      return { type: "unary", op, value: parseUnary() };
    }
    return parsePostfix();
  }
  function parsePostfix() {
    let value = parsePrimary();
    while (peek() && peek().type === "op" && peek().value === "%") {
      take();
      value = { type: "percent", value };
    }
    return value;
  }
  function parsePrimary() {
    const token = take();
    if (!token) throw formulaError(FORMULA_ERRORS.value);
    if (token.type === "num") return { type: "num", value: token.value };
    if (token.type === "str") return { type: "str", value: token.value };
    if (token.type === "punct" && token.value === "(") {
      const inner = parseExpr();
      const close = take();
      if (!close || close.value !== ")") throw formulaError(FORMULA_ERRORS.value);
      return inner;
    }
    if (token.type === "ident") {
      if (token.value === "TRUE") return { type: "bool", value: true };
      if (token.value === "FALSE") return { type: "bool", value: false };
      if (peek() && peek().type === "punct" && peek().value === "(") {
        take();
        const args = [];
        if (peek() && peek().value === ")") take();
        else {
          for (;;) {
            args.push(parseExpr());
            const next = take();
            if (!next || next.value === ")") break;
            if (next.value !== ",") throw formulaError(FORMULA_ERRORS.value);
          }
        }
        return { type: "call", name: token.value, args };
      }
      // Range: cell ":" cell, or whole-column A:A (rows bounded by the
      // preview row cap at evaluation time).
      if (/^[A-Z]{1,3}$/.test(token.value) && peek() && peek().type === "punct" && peek().value === ":") {
        take();
        const second = take();
        if (!second || second.type !== "ident" || !/^[A-Z]{1,3}$/.test(second.value)) throw formulaError(FORMULA_ERRORS.ref);
        return { type: "range", from: token.value + "1", to: second.value + "1000" };
      }
      if (/^[A-Z]{1,3}[0-9]+$/.test(token.value) && peek() && peek().type === "punct" && peek().value === ":") {
        take();
        const second = take();
        const toAddress = second && second.type === "ident" ? second.value : "";
        if (!/^[A-Z]{1,3}[0-9]+$/.test(toAddress)) throw formulaError(FORMULA_ERRORS.ref);
        return { type: "range", from: token.value, to: toAddress };
      }
      if (/^[A-Z]{1,3}[0-9]+$/.test(token.value)) return { type: "ref", address: token.value };
      throw formulaError(FORMULA_ERRORS.name);
    }
    throw formulaError(FORMULA_ERRORS.value);
  }
  const ast = parseExpr();
  if (pos !== tokens.length) throw formulaError(FORMULA_ERRORS.value);
  return ast;
}

// Every function receives the evaluated args array (ranges pre-expanded to
// flat value arrays). Arg errors are propagated by evalCall before dispatch,
// except COUNT/COUNTA which ignore them.
function num(value) {
  return typeof value === "number" ? value : Number(value) || 0;
}
const FORMULA_FUNCTIONS = {
  SUM: (args) => args.flatMap(flattenNumbers).reduce((s, v) => s + v, 0),
  AVERAGE: (args) => {
    const nums = args.flatMap(flattenNumbers);
    return nums.length ? nums.reduce((s, v) => s + v, 0) / nums.length : formulaError(FORMULA_ERRORS.div0);
  },
  MIN: (args) => {
    const nums = args.flatMap(flattenNumbers);
    return nums.length ? Math.min(...nums) : 0;
  },
  MAX: (args) => {
    const nums = args.flatMap(flattenNumbers);
    return nums.length ? Math.max(...nums) : 0;
  },
  COUNT: (args) => args.flatMap(flattenNumbers).filter((v) => typeof v === "number").length,
  COUNTA: (args) => args.flatMap(flattenValues).filter((v) => v !== null && v !== undefined && v !== "").length,
  PRODUCT: (args) => args.flatMap(flattenNumbers).reduce((p, v) => p * v, 1),
  MEDIAN: (args) => {
    const nums = args.flatMap(flattenNumbers).sort((a, b) => a - b);
    if (!nums.length) return formulaError(FORMULA_ERRORS.num);
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
  },
  AND: (args) => args.every(isTrue),
  OR: (args) => args.some(isTrue),
  NOT: (args) => !isTrue(args[0]),
  ROUND: (args) => {
    const factor = Math.pow(10, num(args[1]) || 0);
    return Math.round((num(args[0]) + Number.EPSILON * Math.sign(num(args[0]) || 1)) * factor) / factor;
  },
  ROUNDUP: (args) => {
    const factor = Math.pow(10, num(args[1]) || 0);
    return num(args[0]) >= 0 ? Math.ceil(num(args[0]) * factor) / factor : Math.floor(num(args[0]) * factor) / factor;
  },
  ROUNDDOWN: (args) => {
    const factor = Math.pow(10, num(args[1]) || 0);
    return num(args[0]) >= 0 ? Math.floor(num(args[0]) * factor) / factor : Math.ceil(num(args[0]) * factor) / factor;
  },
  ABS: (args) => Math.abs(num(args[0])),
  SQRT: (args) => (num(args[0]) < 0 ? formulaError(FORMULA_ERRORS.num) : Math.sqrt(num(args[0]))),
  POWER: (args) => Math.pow(num(args[0]), num(args[1])),
  MOD: (args) => (num(args[1]) === 0 ? formulaError(FORMULA_ERRORS.div0) : num(args[0]) - num(args[1]) * Math.floor(num(args[0]) / num(args[1]))),
  INT: (args) => Math.floor(num(args[0]))
};
function flattenNumbers(value) {
  if (Array.isArray(value)) return value.flatMap(flattenNumbers);
  if (typeof value === "number") return [value];
  if (typeof value === "boolean") return [value ? 1 : 0];
  if (typeof value === "string") { const num = Number(value); return value.trim() !== "" && !Number.isNaN(num) ? [num] : []; }
  return [];
}
function flattenValues(value) {
  return Array.isArray(value) ? value.flatMap(flattenValues) : [value];
}
function isTrue(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") return value.toUpperCase() === "TRUE";
  return false;
}

// Evaluate a formula AST against a cell resolver. Range args expand eagerly
// (bounded); IF evaluates lazily so the untaken branch cannot error.
function evalFormulaNode(node, resolve) {
  try {
    return evalValue(node, resolve);
  } catch (error) {
    if (isFormulaError(error)) return error;
    throw error;
  }
}
function evalValue(node, resolve) {
  switch (node.type) {
    case "num": return node.value;
    case "str": return node.value;
    case "bool": return node.value;
    case "ref": return resolve(node.address);
    case "range": {
      const from = addressParts(node.from);
      const to = addressParts(node.to);
      if (!from || !to) throw formulaError(FORMULA_ERRORS.ref);
      const c1 = colToIndex(from.letters), c2 = colToIndex(to.letters);
      const r1 = from.row, r2 = to.row;
      const colMin = Math.min(c1, c2), colMax = Math.max(c1, c2);
      const rowMin = Math.min(r1, r2), rowMax = Math.max(r1, r2);
      if ((colMax - colMin + 1) * (rowMax - rowMin + 1) > FORMULA_MAX_RANGE_CELLS) throw formulaError(FORMULA_ERRORS.ref);
      const values = [];
      for (let r = rowMin; r <= rowMax; r++) {
        for (let c = colMin; c <= colMax; c++) values.push(resolve(`${columnToLetters(c)}${r}`));
      }
      return values;
    }
    case "unary": {
      const num = formulaToNumber(evalValue(node.value, resolve));
      if (isFormulaError(num)) return num;
      return node.op === "-" ? -num : num;
    }
    case "percent": {
      const num = formulaToNumber(evalValue(node.value, resolve));
      return isFormulaError(num) ? num : num / 100;
    }
    case "binary": return evalBinary(node, resolve);
    case "call": return evalCall(node, resolve);
    default: throw formulaError(FORMULA_ERRORS.value);
  }
}
function evalBinary(node, resolve) {
  if (["=", "<>", "<", ">", "<=", ">="].includes(node.op)) {
    const left = evalValue(node.left, resolve), right = evalValue(node.right, resolve);
    if (isFormulaError(left)) return left;
    if (isFormulaError(right)) return right;
    const cmp = formulaCompare(left, right);
    switch (node.op) {
      case "=": return cmp === 0;
      case "<>": return cmp !== 0;
      case "<": return cmp < 0;
      case ">": return cmp > 0;
      case "<=": return cmp <= 0;
      case ">=": return cmp >= 0;
    }
  }
  // String concatenation happens on display strings, never numeric coercion.
  if (node.op === "&") {
    const left = evalValue(node.left, resolve);
    if (isFormulaError(left)) return left;
    const right = evalValue(node.right, resolve);
    if (isFormulaError(right)) return right;
    return formulaDisplay(left) + formulaDisplay(right);
  }
  const left = formulaToNumber(evalValue(node.left, resolve));
  if (isFormulaError(left)) return left;
  const right = formulaToNumber(evalValue(node.right, resolve));
  if (isFormulaError(right)) return right;
  switch (node.op) {
    case "+": return left + right;
    case "-": return left - right;
    case "*": return left * right;
    case "/": return right === 0 ? formulaError(FORMULA_ERRORS.div0) : left / right;
    case "^": return Math.pow(left, right);
    default: throw formulaError(FORMULA_ERRORS.value);
  }
}
function evalCall(node, resolve) {
  // IF is lazy: the untaken branch must not be able to error.
  if (node.name === "IF") {
    if (node.args.length < 2 || node.args.length > 3) throw formulaError(FORMULA_ERRORS.value);
    const condition = evalValue(node.args[0], resolve);
    if (isFormulaError(condition)) return condition;
    return isTrue(condition) ? evalValue(node.args[1], resolve) : (node.args[2] ? evalValue(node.args[2], resolve) : false);
  }
  const fn = FORMULA_FUNCTIONS[node.name];
  if (!fn) throw formulaError(FORMULA_ERRORS.name);
  const args = node.args.map((arg) => evalValue(arg, resolve));
  for (const arg of args) {
    if (isFormulaError(arg) && node.name !== "COUNT" && node.name !== "COUNTA") return arg;
  }
  return fn(args);
}
function formulaDisplay(value) {
  if (isFormulaError(value)) return value.__err;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value === null || value === undefined) return "";
  return String(value);
}

// ---- Number rendering for Excel format codes -------------------------------
// Renders the subset of Excel number formats the grid supports: percent
// (x100 + %), thousands grouping, fixed decimals, a leading $ currency
// symbol. Anything else (dates, bracketed conditions, text sections) falls
// back to the raw value — real Excel stays the source of truth.
function renderExcelNumber(format, value) {
  const num = typeof value === "number" ? value : Number(value);
  if (typeof value === "boolean" || value === null || value === undefined || !Number.isFinite(num)) {
    return String(value);
  }
  const section = String(format || "General").split(";")[0];
  if (section === "General" || !/^[^"\[\]yYdDhHsS]*$/.test(section)) return String(value);
  const isPercent = section.includes("%");
  const hasCurrency = section.startsWith("$");
  const thousands = section.includes(",");
  const decMatch = section.match(/\.(0+)/);
  const decimals = decMatch ? decMatch[1].length : 0;
  const scaled = num * (isPercent ? 100 : 1);
  let text = Math.abs(scaled).toFixed(decimals);
  if (thousands) {
    const [intPart, decPart] = text.split(".");
    text = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (decPart ? "." + decPart : "");
  }
  if (scaled < 0) text = "-" + text;
  return (hasCurrency ? "$" : "") + text + (isPercent ? "%" : "");
}
