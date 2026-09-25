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
  pendingUpload: null,
  pendingSubmission: null,
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
const PENDING_SUBMISSION_KEY = `sxl-pending-submission:${API_BASE}`;
const PENDING_UPLOAD_KEY = `sxl-pending-upload:${API_BASE}`;
const SELECTED_WORKBOOK_KEY = `sxl-selected-workbook:${API_BASE}`;

// Supabase Auth config (injected as meta tags on the hosted page). Absent in
// local mode, where the app keeps using the anonymous local token.
const AUTH_URL = ((document.querySelector('meta[name="sxl-auth-url"]') || {}).content || "").replace(/\/+$/, "");
const AUTH_ANON = (document.querySelector('meta[name="sxl-auth-anon-key"]') || {}).content || "";
const AUTH_ENABLED = Boolean(AUTH_URL && AUTH_ANON);
const OAUTH_PROVIDERS = ["google", "github", "azure", "apple", "gitlab", "bitbucket", "discord", "linkedin_oidc"];
const PROVIDER_LABELS = { azure: "Microsoft", linkedin_oidc: "LinkedIn" };

function clearAccountDrafts() {
  state.pendingSubmission = null;
  state.pendingUpload = null;
  state.pendingFileId = null;
  state.workbook = null;
  state.pending = {};
  $("recentFilesPanel").hidden = true;
  $("recentFilesBtn").setAttribute("aria-expanded", "false");
  $("recentFilesList").replaceChildren();
  try {
    sessionStorage.removeItem(PENDING_SUBMISSION_KEY);
    sessionStorage.removeItem(PENDING_UPLOAD_KEY);
    sessionStorage.removeItem(SELECTED_WORKBOOK_KEY);
  } catch { /* private browsing may deny storage */ }
  $("prompt").value = "";
  $("file").value = "";
}

try {
  // Hosted retry state belongs to the signed-in account. An unsigned tab must
  // not revive a prior account's prompt before authentication finishes.
  if (!AUTH_ENABLED || localStorage.getItem("sxl.platform.token")) {
    const saved = JSON.parse(sessionStorage.getItem(PENDING_SUBMISSION_KEY) || "null");
    if (saved && typeof saved.key === "string" && typeof saved.bodyJson === "string") {
      state.pendingSubmission = saved;
      const prior = JSON.parse(saved.bodyJson);
      if (typeof prior.prompt === "string" && prior.prompt) $("prompt").value = prior.prompt;
      if (prior.mode === "ask" || prior.mode === "action") $("mode").value = prior.mode;
    }
  }
} catch { /* private browsing may deny storage */ }

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
  const priorToken = localStorage.getItem("sxl.platform.token");
  if (AUTH_ENABLED && priorToken !== body.token) clearAccountDrafts();
  state.token = body.token;
  localStorage.setItem("sxl.platform.token", body.token);
  if (body.username) localStorage.setItem("sxl.platform.username", body.username);
  if (AUTH_ENABLED && priorToken && priorToken !== body.token) location.reload();
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
    clearAccountDrafts();
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
  if (AUTH_ENABLED) {
    $("accountPanel").hidden = false;
    setAuthStatus("Sign in with email or a provider to use the hosted workspace.");
    throw new Error("Sign in to use the hosted workspace.");
  }
  return login();
}

async function authFetch(path, options = {}) {
  const token = await ensureToken();
  const headers = Object.assign({}, options.headers || {}, { authorization: `Bearer ${token}` });
  let response = await fetch(api(path), Object.assign({}, options, { headers }));
  if (response.status === 401) {
    if (AUTH_ENABLED) {
      clearAccountDrafts();
      state.token = null;
      localStorage.removeItem("sxl.platform.token");
      localStorage.removeItem("sxl.platform.username");
      $("accountPanel").hidden = false;
      setAuthStatus("Your session expired. Sign in again to continue.");
      location.reload();
      throw new Error("Your session expired. Sign in again to continue.");
    }
    await login();
    response = await fetch(api(path), Object.assign({}, options, {
      headers: Object.assign({}, options.headers || {}, { authorization: `Bearer ${state.token}` })
    }));
  }
  return response;
}

async function uploadFile(file) {
  if (state.pendingUpload?.file === file && state.pendingUpload.fileId) {
    return { fileId: state.pendingUpload.fileId, filename: file.name, size: file.size, reused: true };
  }
  const fingerprint = `${file.name}:${file.size}:${file.lastModified}`;
  let pending = null;
  try { pending = JSON.parse(sessionStorage.getItem(PENDING_UPLOAD_KEY) || "null"); } catch { /* private browsing */ }
  if (!pending || pending.fingerprint !== fingerprint) {
    pending = { fingerprint, key: crypto.randomUUID() };
    try { sessionStorage.setItem(PENDING_UPLOAD_KEY, JSON.stringify(pending)); } catch { /* private browsing */ }
  }
  const form = new FormData();
  form.append("file", file);
  const response = await authFetch("/api/spreadsheets/upload", {
    method: "POST", headers: { "idempotency-key": pending.key }, body: form
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 409) {
      try { sessionStorage.removeItem(PENDING_UPLOAD_KEY); } catch { /* private browsing */ }
    }
    throw new Error(body.error || `upload failed: HTTP ${response.status}`);
  }
  state.pendingUpload = { file, fileId: body.fileId };
  return body;
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

function setWorkbook(workbook, remember = true) {
  state.workbook = workbook;
  state.selected = "A1";
  state.pending = {};
  updatePendingBar();
  $("workbookTitle").textContent = workbook.fileName || "untitled";
  renderSheetTabs();
  renderGrid();
  if (remember) {
    const reference = workbook.runId
      ? { runId: workbook.runId }
      : workbook.fileId ? { fileId: workbook.fileId, fileName: workbook.fileName } : null;
    try {
      if (reference) {
        sessionStorage.setItem(SELECTED_WORKBOOK_KEY, JSON.stringify({
          ...reference,
          account: AUTH_ENABLED ? localStorage.getItem("sxl.platform.username") : null
        }));
      } else sessionStorage.removeItem(SELECTED_WORKBOOK_KEY);
    } catch { /* private browsing may deny storage */ }
  }
}

async function restoreSelectedWorkbook() {
  let saved;
  try { saved = JSON.parse(sessionStorage.getItem(SELECTED_WORKBOOK_KEY) || "null"); }
  catch { /* private browsing or stale data */ }
  if (!saved) return false;
  const account = AUTH_ENABLED ? localStorage.getItem("sxl.platform.username") : null;
  if (AUTH_ENABLED && (!localStorage.getItem("sxl.platform.token") || !account || saved.account !== account)) {
    try { sessionStorage.removeItem(SELECTED_WORKBOOK_KEY); } catch { /* private browsing */ }
    return false;
  }
  try {
    const loaded = typeof saved.runId === "string" && saved.runId
      ? await loadWorkbookFromRun(saved.runId)
      : typeof saved.fileId === "string" && saved.fileId && typeof saved.fileName === "string"
        ? await loadWorkbookFromFileId(saved.fileId, saved.fileName) : false;
    if (loaded) {
      setStatus(`Restored ${state.workbook.fileName} from your workspace.`);
      return true;
    }
  } catch { /* missing or inaccessible workbook is cleared below */ }
  try { sessionStorage.removeItem(SELECTED_WORKBOOK_KEY); } catch { /* private browsing */ }
  return false;
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
  if (!wb) { $("pendingWrap").hidden = true; $("editLaneNote").hidden = true; return; }
  const activeName = wb.sheets[wb.active].name;
  const activePending = state.pending[activeName];
  const count = activePending ? activePending.size : 0;
  const elsewhere = Object.keys(state.pending)
    .filter((name) => name !== activeName && state.pending[name] && state.pending[name].size > 0)
    .reduce((sum, name) => sum + state.pending[name].size, 0);
  const hasPending = count > 0 || elsewhere > 0;
  $("pendingWrap").hidden = !hasPending;
  $("editLaneNote").hidden = !hasPending || $("editEngine").value !== "headless_value";
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
    + (wb.truncated ? " (preview limited; some cells or sheets are hidden)" : "");
  selectCell(state.selected, effectiveCell(sheet, state.selected));
}

// Pending edits override the stored cell for display and the formula bar.
function effectiveCell(sheet, address) {
  const pending = state.pending[sheet.name];
  const staged = pending && pending.get(address);
  if (staged) {
    return staged.formula
      ? { f: staged.formula, italic: staged.italic, wrap: staged.wrap, align: staged.align }
      : { v: staged.value, italic: staged.italic, wrap: staged.wrap, align: staged.align };
  }
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

// Excel's plain-text clipboard is tabular TSV. Parse quoted cells (including
// embedded tabs/newlines) before staging anything, so a failed paste cannot
// leave a partial draft. The submit API currently accepts at most 50 edits.
function parseGridPaste(text) {
  if (typeof text !== "string" || !text || text.length > 50000) {
    throw new Error("Paste must contain at most 50,000 characters.");
  }
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  let closed = false;
  let atStart = true;
  let cells = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') { quoted = false; closed = true; }
      else field += ch;
      if (field.length > 1000) throw new Error("A pasted cell exceeds the 1,000-character edit limit.");
      continue;
    }
    if (ch === '"' && atStart) { quoted = true; atStart = false; continue; }
    if (ch === "\t" || ch === "\r" || ch === "\n") {
      row.push(field);
      if (++cells > 50) throw new Error("Paste exceeds the 50-edit limit.");
      if (ch === "\t") {
        field = ""; closed = false; atStart = true;
      } else {
        rows.push(row);
        row = []; field = ""; closed = false; atStart = true;
        if (ch === "\r" && text[i + 1] === "\n") i++;
      }
      continue;
    }
    if (closed) throw new Error("Paste contains text after a quoted cell.");
    field += ch;
    atStart = false;
    if (field.length > 1000) throw new Error("A pasted cell exceeds the 1,000-character edit limit.");
  }
  if (quoted) throw new Error("Paste has an unfinished quoted cell.");
  if (field || row.length || !rows.length) {
    row.push(field);
    if (++cells > 50) throw new Error("Paste exceeds the 50-edit limit.");
  }
  if (row.length) rows.push(row);
  if (!rows.length || rows.some((cells) => !cells.length)) throw new Error("Paste has no cells.");
  return rows;
}

function stageGridPaste(text) {
  const wb = state.workbook;
  if (!wb) return false;
  const sheet = wb.sheets[wb.active];
  const start = addressParts(state.selected);
  if (!sheet || !start) return false;
  try {
    const rows = parseGridPaste(text);
    const startColumn = colToIndex(start.letters);
    const lastRow = start.row + rows.length - 1;
    const lastColumn = startColumn + Math.max(...rows.map((cells) => cells.length)) - 1;
    if (lastRow > GRID_MAX_ROWS || lastColumn > GRID_MAX_COLS ||
        (wb.truncated && (lastRow > state.rendered.rows || lastColumn > state.rendered.cols))) {
      throw new Error("Paste exceeds the visible grid limit; choose a smaller block.");
    }
    const pending = activePending();
    const edits = [];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const value = rows[r][c];
        if (value.length > 1000) throw new Error("A pasted cell exceeds the 1,000-character edit limit.");
        const address = `${columnToLetters(startColumn + c)}${start.row + r}`;
        edits.push([address, value.startsWith("=") ? { formula: value } : { value }]);
      }
    }
    if (new Set([...pending.keys(), ...edits.map(([address]) => address)]).size > 50) {
      throw new Error("Paste would exceed 50 staged edits on this sheet. Apply or discard the current edits first.");
    }
    for (const [address, edit] of edits) pending.set(address, edit);
    sheet.rows = Math.max(sheet.rows || 1, lastRow);
    sheet.cols = Math.max(sheet.cols || 1, lastColumn);
    updatePendingBar();
    renderGrid();
    setStatus(`Staged ${edits.length} pasted cell${edits.length === 1 ? "" : "s"} on ${sheet.name}. Apply to run the audited edit.`);
    return true;
  } catch (error) {
    setStatus(`Paste not staged: ${error.message || error}`);
    return false;
  }
}

function onGridPaste(event) {
  const tag = (event.target.tagName || "").toLowerCase();
  if (tag === "input" || tag === "textarea") return;
  const text = event.clipboardData?.getData("text/plain");
  if (typeof text !== "string" || !text) return;
  event.preventDefault();
  commitActiveEditor();
  stageGridPaste(text);
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
async function openWorkbookFile(file, requirePreview = false) {
  const previousFileId = state.pendingFileId;
  const previousUpload = state.pendingUpload;
  try {
    setStatus("Uploading file...");
    const uploaded = await uploadFile(file);
    state.pendingFileId = uploaded.fileId;
    const previewLoaded = await loadWorkbookFromFileId(uploaded.fileId, uploaded.filename || file.name);
    if (!previewLoaded) {
      if (requirePreview) {
        state.pendingFileId = previousFileId;
        state.pendingUpload = previousUpload;
        $("file").value = "";
      }
      return false;
    }
    setStatus(`Loaded ${file.name}${state.workbook?.truncated ? " with a limited grid preview" : ""}.`);
    return true;
  } catch (error) {
    if (requirePreview) {
      state.pendingFileId = previousFileId;
      state.pendingUpload = previousUpload;
      $("file").value = "";
    }
    setStatus(`ERROR: ${error.message || error}`);
    return false;
  }
}

async function showRecentWorkbooks() {
  const panel = $("recentFilesPanel");
  panel.hidden = !panel.hidden;
  $("recentFilesBtn").setAttribute("aria-expanded", String(!panel.hidden));
  if (panel.hidden) return;
  const list = $("recentFilesList");
  list.replaceChildren();
  $("recentFilesStatus").textContent = "Loading your workbooks…";
  try {
    const token = await ensureToken();
    const responses = await Promise.allSettled([
      authFetch("/api/spreadsheets/workbooks"), authFetch("/api/spreadsheets/runs")
    ]);
    const bodies = await Promise.all(responses.map(async (result) =>
      result.status === "fulfilled" && result.value.ok
        ? result.value.json().catch(() => null) : null));
    if (state.token !== token || panel.hidden) return;
    if (!bodies.some((body) => body && (Array.isArray(body.workbooks) || Array.isArray(body.runs)))) {
      throw new Error("workbook history unavailable");
    }
    const files = (Array.isArray(bodies[0]?.workbooks) ? bodies[0].workbooks : []).filter((item) => item &&
      typeof item.fileId === "string" && typeof item.filename === "string")
      .map((item) => ({ ...item, kind: "file" }));
    const runs = (Array.isArray(bodies[1]?.runs) ? bodies[1].runs : []).filter((run) => run && run.status === "completed" &&
      typeof run.runId === "string")
      .map((run) => ({ run, artifact: (Array.isArray(run.artifacts) ? run.artifacts : [])
        .find((name) => typeof name === "string" && /\.xlsx?$/i.test(name)) }))
      .filter(({ run, artifact }) => run.downloadUrl || artifact)
      .map(({ run, artifact }) => ({
        kind: "run", runId: run.runId,
        filename: run.downloadUrl ? "workbook.xlsx" : artifact,
        createdAt: run.updatedAt || run.createdAt,
        prompt: String(run.prompt || "").slice(0, 40)
      }));
    const items = [...files, ...runs]
      .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")))
      .slice(0, 30);
    $("recentFilesStatus").textContent = items.length ? "Select a workbook to reopen." : "No saved workbooks yet.";
    for (const file of items) {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = file.kind === "run" && file.prompt
        ? `${file.filename} · ${file.prompt}` : file.filename;
      const detail = document.createElement("small");
      const when = new Date(file.createdAt);
      detail.textContent = `${Number.isFinite(when.getTime()) ? when.toLocaleString() : "Saved"} · ${
        file.kind === "run" ? `Run ${file.runId.slice(0, 8)}` : `${Number(file.size) || 0} bytes`}`;
      button.appendChild(detail);
      button.addEventListener("click", async () => {
        if (state.token !== token) return;
        if (Object.values(state.pending).some((edits) => edits.size > 0) &&
            !window.confirm("Discard staged edits and open another workbook?")) return;
        const loaded = file.kind === "run"
          ? await loadWorkbookFromRun(file.runId) : await loadWorkbookFromFileId(file.fileId, file.filename);
        if (loaded) {
          panel.hidden = true;
          $("recentFilesBtn").setAttribute("aria-expanded", "false");
          setStatus(`Opened ${state.workbook.fileName} from your workspace.`);
        }
      });
      list.appendChild(button);
    }
  } catch (error) {
    if (!panel.hidden) $("recentFilesStatus").textContent = `Could not load workbooks: ${error.message || error}`;
  }
}

async function createNewWorkbook() {
  const button = $("newFile");
  button.disabled = true;
  try {
    const response = await fetch("blank.xlsx", { cache: "no-store" });
    if (!response.ok) throw new Error(`blank workbook unavailable: HTTP ${response.status}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.length < 100 || bytes.length > 100_000 ||
        bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
      throw new Error("blank workbook asset is invalid");
    }
    const file = new File([bytes], "new-workbook.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    if (await openWorkbookFile(file, true)) {
      $("file").value = "";
      setStatus("New workbook ready. Stage edits, then Apply to create an audited run.");
    }
  } catch (error) {
    setStatus(`ERROR: ${error.message || error}`);
  } finally {
    button.disabled = false;
  }
}

async function loadWorkbookFromFileId(fileId, filename) {
  const accountToken = AUTH_ENABLED ? await ensureToken() : null;
  const response = await authFetch(`/api/spreadsheets/workbook/${encodeURIComponent(fileId)}/sheet-data`);
  const body = await response.json().catch(() => ({}));
  if (AUTH_ENABLED && state.token !== accountToken) return false;
  if (!response.ok) {
    setStatus(body.error || `no grid preview for ${filename} (HTTP ${response.status})`);
    if (!state.workbook) setWorkbook(emptyWorkbook(filename));
    return false;
  }
  setWorkbook({
    fileName: filename || body.filename || "workbook.xlsx",
    fileId,
    runId: null,
    sheets: body.sheets,
    active: 0,
    truncated: body.truncated
  });
  return true;
}

async function loadWorkbookFromRun(runId, artifact) {
  const accountToken = AUTH_ENABLED ? await ensureToken() : null;
  const query = artifact ? `?artifact=${encodeURIComponent(artifact)}` : "";
  const response = await authFetch(`/api/spreadsheets/${encodeURIComponent(runId)}/sheet-data${query}`);
  const body = await response.json().catch(() => ({}));
  if (AUTH_ENABLED && state.token !== accountToken) return false;
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
  if (run.downloadUrl) add(run.revertedAt && run.revertJob?.runId ? "Download original result" : "Download workbook", api(run.downloadUrl));
  if (run.revertedAt && run.revertJob?.runId) {
    add("Download reverted workbook", api(`/api/spreadsheets/${encodeURIComponent(run.revertJob.runId)}/download`));
  }
  for (const artifact of run.artifacts || []) {
    if (artifact === "workbook.xlsx" && run.downloadUrl) continue;
    add(artifact === "sxl-audit-receipt.json" ? "Download audit receipt" : `Download ${artifact}`,
      api(`/api/spreadsheets/${run.runId}/artifacts/${encodeURIComponent(artifact)}`));
  }
  if (wrap.children.length) target.appendChild(wrap);
}

function snapshotLabel(snapshot, other, kind) {
  if (!snapshot || typeof snapshot !== "object") return "(empty)";
  if (kind === "presentation") return `style ${JSON.stringify(snapshot).slice(0, 200)}`;
  if (kind === "structural") return JSON.stringify(snapshot).slice(0, 200);
  const value = snapshot.formula != null ? String(snapshot.formula)
    : snapshot.value != null && snapshot.value !== "" ? String(snapshot.value) : "(empty)";
  const details = [];
  for (const [field, label] of [["format", "format"], ["note", "note"], ["fontColor", "font color"]]) {
    if (snapshot[field] !== other?.[field]) {
      details.push(`${label}: ${String(snapshot[field] ?? "(none)").replace(/\s+/g, " ").slice(0, 100)}`);
    }
  }
  return details.length ? `${value} [${details.join("; ")}]` : value;
}

function sameReceiptIds(actual, expected) {
  return Array.isArray(actual) && Array.isArray(expected) &&
    actual.every((id) => typeof id === "string" && id) &&
    expected.every((id) => typeof id === "string" && id) &&
    new Set(actual).size === actual.length &&
    new Set(expected).size === expected.length &&
    actual.length === expected.length && actual.every((id) => expected.includes(id));
}

function addAuditReview(target, run) {
  if (!(run.artifacts || []).includes("sxl-audit-receipt.json")) return;
  const button = document.createElement("button");
  button.className = "mini-btn";
  button.textContent = "Review audit changes";
  const panel = document.createElement("div");
  panel.className = "audit-review";
  let loaded = false;
  button.addEventListener("click", async () => {
    if (loaded) {
      panel.hidden = !panel.hidden;
      button.textContent = panel.hidden ? "Review audit changes" : "Hide audit changes";
      return;
    }
    button.disabled = true;
    panel.hidden = false;
    panel.textContent = "Loading audit receipt…";
    try {
      const response = await authFetch(`/api/spreadsheets/${encodeURIComponent(run.runId)}/artifacts/sxl-audit-receipt.json`);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const receipt = await response.json();
      if (receipt.schema !== "sxl.web-audit.v1" || receipt.runId !== run.runId || !Array.isArray(receipt.sessions)) {
        throw new Error("receipt does not match this run");
      }
      if (run.result?.runResultVersion === 1) {
        if (!sameReceiptIds(receipt.sessions.map((session) => session.sessionId), run.result.sessionIds)) {
          throw new Error("receipt sessions do not match this run's ledger");
        }
        const withSetIds = receipt.sessions.filter((session) => Array.isArray(session.mutationSetIds)).length;
        if (withSetIds !== 0 && (withSetIds !== receipt.sessions.length ||
            !sameReceiptIds(receipt.sessions.flatMap((session) => session.mutationSetIds), run.result.mutationSetIds))) {
          throw new Error("receipt mutation sets do not match this run's ledger");
        }
      }
      panel.textContent = "";
      const heading = document.createElement("strong");
      heading.textContent = `${receipt.changeCount} recorded change${receipt.changeCount === 1 ? "" : "s"}`
        + `${receipt.warningCount ? ` · ${receipt.warningCount} review warning${receipt.warningCount === 1 ? "" : "s"}` : ""}`
        + `${run.revertedAt ? " · reverted" : ""}`;
      panel.appendChild(heading);
      const list = document.createElement("ol");
      let shown = 0;
      for (const session of receipt.sessions) {
        for (const change of session.changes || []) {
          if (shown >= 100) break;
          const item = document.createElement("li");
          const before = snapshotLabel(change.before, change.after, change.kind);
          const after = snapshotLabel(change.after, change.before, change.kind);
          item.textContent = `${change.sheetName}!${change.address}: ${before} → ${after}`;
          if (change.sourceLabel) item.textContent += ` · Source label: ${change.sourceLabel}`;
          if (change.evidence?.note) item.textContent += ` · Evidence: ${String(change.evidence.note).slice(0, 220)}`;
          if (change.explanation) item.textContent += ` · ${change.explanation}`;
          if (Array.isArray(change.warnings) && change.warnings.length) {
            item.textContent += ` · Review: ${change.warnings.join("; ")}`;
          }
          if (change.mutationRecordId) item.textContent += ` · Ledger record: ${change.mutationRecordId}`;
          list.appendChild(item);
          shown++;
        }
        if (shown >= 100) break;
      }
      panel.appendChild(list);
      for (const session of receipt.sessions) {
        if (session.prompt || (Array.isArray(session.sources) && session.sources.length)) {
          const context = document.createElement("details");
          const summary = document.createElement("summary");
          summary.textContent = `Request and inputs for session ${session.sessionId}`;
          context.appendChild(summary);
          const request = document.createElement("p");
          request.textContent = `Request: ${String(session.prompt || "(not recorded)").slice(0, 1000)}`;
          context.appendChild(request);
          if (String(session.prompt || "").length > 1000) {
            const note = document.createElement("p");
            note.textContent = "Request shortened here; download the receipt for the full text.";
            context.appendChild(note);
          }
          const sources = document.createElement("p");
          sources.textContent = `Inputs: ${Array.isArray(session.sources) && session.sources.length
            ? session.sources.map((source) => String(source)).join(", ") : "No source attachments recorded"}`;
          context.appendChild(sources);
          panel.appendChild(context);
        }
        if (!Array.isArray(session.milestones) || !session.milestones.length) continue;
        const details = document.createElement("details");
        const summary = document.createElement("summary");
        summary.textContent = `Recorded milestones for session ${session.sessionId} (${session.milestones.length})`;
        details.appendChild(summary);
        const history = document.createElement("ol");
        for (const milestone of session.milestones.slice(0, 100)) {
          const item = document.createElement("li");
          item.textContent = `${String(milestone.status || "").replaceAll("_", " ")} · ${String(milestone.timestamp || "").slice(0, 40)}`;
          history.appendChild(item);
        }
        details.appendChild(history);
        if (session.milestones.length > 100) {
          const note = document.createElement("p");
          note.textContent = `Showing 100 of ${session.milestones.length} milestones; download the receipt for the full record.`;
          details.appendChild(note);
        }
        panel.appendChild(details);
      }
      if (receipt.changeCount > shown) {
        const note = document.createElement("p");
        note.textContent = `Showing ${shown} of ${receipt.changeCount}; download the receipt for the full record.`;
        panel.appendChild(note);
      }
      loaded = true;
      button.textContent = "Hide audit changes";
      button.disabled = false;
    } catch (error) {
      panel.textContent = `Audit receipt unavailable: ${error.message || error}`;
      button.disabled = false;
    }
  });
  target.appendChild(button);
  target.appendChild(panel);
  panel.hidden = true;
}

function addRunActivity(target, run) {
  const details = document.createElement("details");
  details.className = "run-activity";
  const label = document.createElement("summary");
  label.textContent = "Run activity";
  const panel = document.createElement("div");
  details.appendChild(label);
  details.appendChild(panel);
  let nextAfter = 0;
  let loading = false;
  let done = false;
  const list = document.createElement("ol");
  const load = async () => {
    if (!details.open || loading || done) return;
    loading = true;
    if (nextAfter === 0) panel.textContent = "Loading run activity…";
    try {
      const response = await authFetch(`/api/spreadsheets/${encodeURIComponent(run.runId)}/events` +
        (nextAfter ? `?after=${nextAfter}` : ""));
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const stream = await response.json();
      if (stream.eventStreamVersion !== 1 || stream.runId !== run.runId || !Array.isArray(stream.events) ||
          !Number.isInteger(stream.nextAfter) || stream.nextAfter < nextAfter) {
        throw new Error("run event response does not match this run");
      }
      let pageCursor = nextAfter;
      for (const event of stream.events) {
        if (!event || event.eventVersion !== 1 || event.runId !== run.runId ||
            (event.type !== "run_status" && event.type !== "run_progress") ||
            !Number.isInteger(event.seq) || event.seq <= pageCursor ||
            typeof event.status !== "string" || !event.status ||
            typeof event.at !== "string" || !Number.isFinite(Date.parse(event.at)) ||
            (event.type === "run_progress" &&
              (event.status !== "running" || typeof event.message !== "string" ||
               !/^[\x20-\x7e]{1,160}$/.test(event.message)))) {
          throw new Error("run event response contains an invalid event");
        }
        pageCursor = event.seq;
      }
      if (pageCursor !== stream.nextAfter) throw new Error("run event response has an invalid cursor");
      if (nextAfter === 0) { panel.textContent = ""; panel.appendChild(list); }
      for (const event of stream.events) {
        const item = document.createElement("li");
        const timestamp = new Date(event.at).toLocaleString();
        const sets = Array.isArray(event.mutationSetIds) ? event.mutationSetIds.length : 0;
        item.textContent = `${event.type === "run_progress" ? event.message : event.status} · ${timestamp}` +
          (sets ? ` · ${sets} audited set${sets === 1 ? "" : "s"}` : "");
        list.appendChild(item);
        nextAfter = event.seq;
        if (event.type !== "run_progress" && ["completed", "failed", "cancelled"].includes(event.status)) done = true;
      }
      if (!done && details.open) setTimeout(load, stream.events.length === 100 ? 0 : 2000);
    } catch (error) {
      panel.textContent = `Run activity unavailable: ${error.message || error}`;
      done = true;
    } finally {
      loading = false;
    }
  };
  details.addEventListener("toggle", async () => {
    if (details.open) await load();
  });
  target.appendChild(details);
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
  state.runId = runId;
  const cancelButton = $("cancelRun");
  cancelButton.hidden = false;
  cancelButton.disabled = false;
  cancelButton.textContent = "Cancel run";
  let last = null;
  try {
    while (state.polling && state.runId === runId) {
      const response = await authFetch(`/api/spreadsheets/${runId}`);
      if (!response.ok) {
        statusMessage(`status failed: HTTP ${response.status}`);
        return null;
      }
      const run = await response.json();
      if (run.status !== last) {
        last = run.status;
        statusMessage(`Run ${runId.slice(0, 8)}: ${run.status}…`);
      }
      if (["completed", "failed", "cancelled"].includes(run.status)) return run;
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }
    return null;
  } finally {
    if (state.runId === runId) {
      state.polling = false;
      cancelButton.hidden = true;
      cancelButton.disabled = false;
    }
  }
}

async function cancelActiveRun() {
  if (!state.polling || !state.runId) return;
  const runId = state.runId;
  const button = $("cancelRun");
  button.disabled = true;
  button.textContent = "Cancelling…";
  try {
    const response = await authFetch(`/api/spreadsheets/${runId}/cancel`, { method: "POST" });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(body.error || `HTTP ${response.status}`);
    setStatus(`Run ${runId.slice(0, 8)} cancelled.`);
  } catch (error) {
    setStatus(`Cancel failed: ${error.message || error}`);
    if (state.runId === runId) {
      button.disabled = false;
      button.textContent = "Cancel run";
    }
  }
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
        const card = addMessage("sys", `Loading run ${run.runId.slice(0, 8)}…`);
        try {
          const statusResponse = await authFetch(`/api/spreadsheets/${run.runId}`);
          const body = await statusResponse.json().catch(() => ({}));
          if (!statusResponse.ok) throw new Error(body.error || `HTTP ${statusResponse.status}`);
          const result = ["completed", "failed", "cancelled"].includes(body.status)
            ? body : await (async () => {
              card.textContent = "";
              const statusLine = document.createElement("span");
              card.appendChild(statusLine);
              addRunActivity(card, body);
              return pollRun(body.runId, (message) => { statusLine.textContent = message; });
            })();
          if (result) await renderRunResult(result, card);
          loadRuns();
        } catch (error) {
          card.className = "msg err";
          card.textContent = `Could not load run: ${error.message || error}`;
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

async function renderRunResult(run, target) {
  if (run.status === "completed") {
    target.className = "msg done";
    const summary = String(run.summary || "").split(/\r?\n/).filter((line) => {
      try {
        const value = JSON.parse(line);
        return !(value && typeof value === "object" && value.ledgerResultVersion !== undefined &&
          typeof value.sessionId === "string");
      } catch { return true; }
    }).join("\n").trim();
    target.textContent = summary ? summary.slice(0, 4000) : `Run ${run.runId.slice(0, 8)} completed.`;
    addDownloadButtons(target, run);
    addAuditReview(target, run);
    if (run.downloadUrl || (run.artifacts || []).some((name) => /\.xlsx?$/i.test(name))) {
      const resultRunId = run.revertedAt && run.revertJob?.runId ? run.revertJob.runId : run.runId;
      const loaded = await loadWorkbookFromRun(resultRunId);
      if (loaded) addMessage("sys", "Result workbook loaded into the grid.");
    }
    if ((run.mutationSetIds || []).length > 0) {
      target.appendChild(document.createElement("br"));
      if (run.revertedAt || run.revertJob) {
        const undoState = document.createElement("span");
        undoState.textContent = run.revertedAt ? "Reverted through the audited ledger."
          : `Undo ${run.revertJob.status}; open its run in History for details.`;
        target.appendChild(undoState);
      } else {
        const undoBtn = document.createElement("button");
        undoBtn.className = "mini-btn";
        undoBtn.textContent = "↩ Undo this run";
        undoBtn.title = "Replay the audited inverse (sxl revert) of this run";
        undoBtn.addEventListener("click", () => revertRun(run.runId, undoBtn));
        target.appendChild(undoBtn);
      }
    }
  } else if (run.status === "cancelled") {
    target.className = "msg err";
    target.textContent = "Run cancelled.";
  } else {
    target.className = "msg err";
    target.textContent = `Run ${run.status}: ${run.error || "no details"}`;
    const recorded = run.result?.mutationSetIds || run.mutationSetIds || [];
    if (recorded.length > 0) {
      const note = document.createElement("p");
      note.textContent = `${recorded.length} audited mutation set${recorded.length === 1 ? "" : "s"} recorded before failure. ` +
        (run.downloadUrl ? "Download the recovery workbook before retrying." : "No recovery workbook was published.");
      target.appendChild(note);
      const ids = document.createElement("details");
      const label = document.createElement("summary");
      label.textContent = "Show audit IDs";
      const values = document.createElement("pre");
      const sessions = run.result?.sessionIds || run.sessionIds || [];
      values.textContent = `Session IDs: ${sessions.join(", ") || "none"}\nMutation set IDs: ${recorded.join(", ")}`;
      ids.appendChild(label);
      ids.appendChild(values);
      target.appendChild(ids);
      addDownloadButtons(target, run);
      addAuditReview(target, run);
    }
  }
  addRunActivity(target, run);
}

async function submitRun(promptText, overrides = {}) {
  const prompt = (promptText !== undefined ? promptText : $("prompt").value).trim();
  if (!prompt && !overrides.edits) {
    setStatus("Enter a prompt first.");
    return;
  }
  $("send").disabled = true;
  if (prompt) addMessage("user", prompt);
  const statusMessage = addMessage("sys", "Submitting…");
  try {
    await ensureToken();
    let initFile = overrides.initFile || null;
    const file = $("file").files && $("file").files[0];
    if (file && !overrides.initFile) {
      if (!state.pendingUpload || state.pendingUpload.file !== file) {
        statusMessage.textContent = "Uploading attachment…";
        const uploaded = await uploadFile(file);
        state.pendingUpload = { file, fileId: uploaded.fileId };
      }
      initFile = state.pendingUpload.fileId;
    } else if (!initFile && state.pendingFileId) {
      initFile = state.pendingFileId;
    }
    const body = {
      prompt,
      mode: overrides.mode || ($("mode").value === "ask" ? "ask" : "action")
    };
    if (initFile) body.initFile = initFile;
    if (!initFile && state.pendingSubmission && !overrides.edits) {
      try {
        const prior = JSON.parse(state.pendingSubmission.bodyJson);
        if (prior.prompt === body.prompt && prior.mode === body.mode && prior.initFile) body.initFile = prior.initFile;
      } catch { /* a malformed saved request is replaced below */ }
    }
    if (overrides.edits) {
      body.edits = overrides.edits;
      if (overrides.sheet) body.sheet = overrides.sheet;
      if (overrides.executionLane) body.executionLane = overrides.executionLane;
    }
    if (!overrides.edits && $("model").value) body.model = $("model").value;
    const bodyJson = JSON.stringify(body);
    if (!state.pendingSubmission || state.pendingSubmission.bodyJson !== bodyJson) {
      state.pendingSubmission = { bodyJson, key: crypto.randomUUID() };
    }
    try { sessionStorage.setItem(PENDING_SUBMISSION_KEY, JSON.stringify(state.pendingSubmission)); }
    catch { /* retry still works in this page */ }
    statusMessage.textContent = "Queued…";
    const response = await authFetch("/api/spreadsheets", {
      method: "POST",
      headers: { "content-type": "application/json", "idempotency-key": state.pendingSubmission.key },
      body: bodyJson
    });
    const submitted = await response.json();
    if (!response.ok) throw new Error(submitted.error || `HTTP ${response.status}`);
    state.pendingSubmission = null;
    try { sessionStorage.removeItem(PENDING_SUBMISSION_KEY); } catch { /* private browsing */ }
    state.pendingUpload = null;
    try { sessionStorage.removeItem(PENDING_UPLOAD_KEY); } catch { /* private browsing */ }
    if (file && !overrides.initFile) $("file").value = "";
    if (promptText !== undefined && !overrides.edits) $("prompt").value = "";
    state.runId = submitted.runId;
    statusMessage.textContent = "";
    const statusLine = document.createElement("span");
    statusMessage.appendChild(statusLine);
    addRunActivity(statusMessage, submitted);
    const run = await pollRun(submitted.runId, (m) => { statusLine.textContent = m; });
    if (!run) return null;
    await renderRunResult(run, statusMessage);
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
  // Loading each result replaces the workbook and normally clears staged
  // edits. Keep the other sheets' maps until their own runs have committed.
  const stagedBySheet = state.pending;
  const activeName = wb.sheets[wb.active].name;
  const sheetNames = Object.keys(state.pending)
    .filter((name) => state.pending[name] && state.pending[name].size > 0)
    .sort((a, b) => (a === activeName ? -1 : b === activeName ? 1 : 0));
  if (!sheetNames.length) return;
  const executionLane = $("editEngine").value;
  const committed = [];
  if (executionLane === "headless_value") {
    if (!/\.xlsx$/i.test(wb.fileName || "")) {
      setStatus("Office-free edits require a backing .xlsx workbook.");
      return;
    }
    for (const sheetName of sheetNames) {
      for (const edit of state.pending[sheetName].values()) {
        if (typeof edit.value !== "string" || edit.value.startsWith("=") ||
            Object.keys(edit).some((key) => key !== "value")) {
          setStatus("Office-free edits support plain values only. Choose Live Excel for formulas or formatting.");
          return;
        }
      }
    }
  }
  for (const sheetName of sheetNames) {
    const pending = stagedBySheet[sheetName];
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
        const prior = committed.map(({ sheet, runId }) => `${sheet} (${runId.slice(0, 8)})`).join(", ");
        setStatus(`${prior ? `Partial edit: ${prior} committed. ` : ""}Could not prepare the result for ${sheetName}: ${error.message || error}. ${prior ? "Review History before retrying." : ""}`.trim());
        return;
      }
    }
    if (!fileId) {
      const prior = committed.map(({ sheet, runId }) => `${sheet} (${runId.slice(0, 8)})`).join(", ");
      setStatus(prior
        ? `Partial edit: ${prior} committed. No result workbook is available for ${sheetName}. Review History before retrying.`
        : "Open a workbook first — an empty sheet has nothing to edit.");
      return;
    }
    const edits = [...pending.entries()].map(([address, edit]) => ({
      address,
      ...(edit.formula ? { formula: edit.formula } : { value: edit.value }),
      ...(edit.format ? { format: edit.format } : {}),
      ...(edit.bold !== undefined ? { bold: edit.bold } : {}),
      ...(edit.italic !== undefined ? { italic: edit.italic } : {}),
      ...(edit.fontSize !== undefined ? { fontSize: edit.fontSize } : {}),
      ...(edit.fillColor ? { fillColor: edit.fillColor } : {}),
      ...(edit.fontColor ? { fontColor: edit.fontColor } : {}),
      ...(edit.align ? { align: edit.align } : {}),
      ...(edit.wrap !== undefined ? { wrap: edit.wrap } : {})
    }));
    const run = await submitRun("", {
      edits,
      sheet: sheetName,
      initFile: fileId,
      mode: "action",
      executionLane
    });
    // Each sheet is a separate durable run. Preserve and identify earlier
    // commits if a later run fails or its outcome cannot be confirmed.
    if (!run || run.status !== "completed") {
      if (committed.length) {
        const prior = committed.map(({ sheet, runId }) => `${sheet} (${runId.slice(0, 8)})`).join(", ");
        setStatus(`Partial edit: ${prior} committed. ${sheetName} ${run ? `ended ${run.status}` : "has an unconfirmed outcome"}. Review History before retrying.`);
      }
      return;
    }
    if (state.workbook?.runId !== run.runId) {
      state.pending = stagedBySheet;
      updatePendingBar();
      const prior = committed.map(({ sheet, runId }) => `${sheet} (${runId.slice(0, 8)})`).join(", ");
      setStatus(`${prior ? `Partial edit: ${prior} committed. ` : ""}${sheetName} run ${run.runId.slice(0, 8)} completed, but its workbook could not be loaded. Open it from History before applying the remaining edits.`);
      return;
    }
    committed.push({ sheet: sheetName, runId: run.runId });
    // Success: submitRun reloaded the result into state.workbook; drop this
    // sheet's staged edits and let the next iteration chain from the new run.
    delete stagedBySheet[sheetName];
    state.pending = stagedBySheet;
    updatePendingBar();
    renderGrid();
  }
}

// Undo (E184): replay the audited inverse of a completed run. Both local and
// hosted revert routes run `sxl revert` on the worker/server ledger — the
// refusal reasons from the ledger guards surface verbatim.
async function revertRun(runId, button) {
  if (button) button.disabled = true;
  const note = addMessage("sys", "Reverting…");
  try {
    await ensureToken();
    const response = await authFetch(`/api/spreadsheets/${runId}/revert`, { method: "POST" });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(body.error || `HTTP ${response.status}`);
    if (body.reverted) {
      note.className = "msg done";
      note.textContent = "Reverted — pre-run values restored through the audited ledger.";
      await loadWorkbookFromRun(runId);
    } else if (body.runId) {
      // Hosted: revert executed as its own queued job on the standing worker.
      note.textContent = "Revert job queued…";
      const job = await pollRun(body.runId, (m) => { note.textContent = m; });
      if (!job) return;
      if (job.status === "completed") {
        note.className = "msg done";
        note.textContent = "Reverted through the audited ledger.";
        await loadWorkbookFromRun(body.runId);
      } else {
        note.className = "msg err";
        note.textContent = `Revert failed: ${job.error || job.status}`;
      }
    }
    loadRuns();
  } catch (error) {
    note.className = "msg err";
    note.textContent = `Revert failed: ${error.message || error}`;
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
    const automatic = document.createElement("option");
    automatic.value = "";
    automatic.textContent = "Automatic (worker default)";
    select.appendChild(automatic);
    for (const model of body.models || []) {
      const option = document.createElement("option");
      option.value = model.id;
      option.textContent = model.id;
      select.appendChild(option);
    }
    if (state.pendingSubmission) {
      const prior = JSON.parse(state.pendingSubmission.bodyJson);
      if (prior.model && Array.from(select.options).some((option) => option.value === prior.model)) {
        select.value = prior.model;
      }
    }
  } catch {
    // run submits without a model field when empty
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  $("openFile").addEventListener("click", () => $("file").click());
  $("recentFilesBtn").addEventListener("click", showRecentWorkbooks);
  $("file").addEventListener("change", () => {
    const file = $("file").files && $("file").files[0];
    if (file) openWorkbookFile(file);
  });
  $("newFile").addEventListener("click", createNewWorkbook);
  $("exportFile").addEventListener("click", exportWorkbook);
  $("send").addEventListener("click", () => submitRun());
  $("cancelRun").addEventListener("click", cancelActiveRun);
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
  $("editEngine").addEventListener("change", updatePendingBar);
  $("discardEdits").addEventListener("click", discardEdits);
  $("refreshGrid").addEventListener("click", async () => {
    const wb = state.workbook;
    if (!wb) return;
    if (wb.runId) await loadWorkbookFromRun(wb.runId);
    else if (wb.fileId) await loadWorkbookFromFileId(wb.fileId, wb.fileName);
    else setStatus("Nothing to refresh yet.");
  });
  $("gridScroll").addEventListener("keydown", onGridKey);
  $("gridScroll").addEventListener("paste", onGridPaste);
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
      } else if (button.dataset.italic) {
        const sheet = state.workbook.sheets[state.workbook.active];
        const current = effectiveCell(sheet, state.selected);
        stageStyle({ italic: !(current && current.italic) });
      } else if (button.dataset.align) {
        stageStyle({ align: button.dataset.align });
      } else if (button.dataset.wrap !== undefined) {
        const sheet = state.workbook.sheets[state.workbook.active];
        const current = effectiveCell(sheet, state.selected);
        stageStyle({ wrap: !(current && current.wrap) });
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
  setWorkbook(emptyWorkbook("empty-sheet.xlsx"), false);
  if (AUTH_ENABLED) {
    await setupAuth();
    await handleAuthRedirect();
    if (!state.token && !localStorage.getItem("sxl.platform.token")) {
      $("accountPanel").hidden = false;
      setStatus("Sign in from Account to use the hosted workspace.");
    }
  }
  await restoreSelectedWorkbook();
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
