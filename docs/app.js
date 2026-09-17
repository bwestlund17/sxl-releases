"use strict";

const $ = (id) => document.getElementById(id);
const state = {
  token: null,
  runId: null,
  polling: false,
  // grid: { fileName, fileId?, runId?, sheets: [{name, cells, rows, cols}], active }
  workbook: null,
  selected: "A1",
  pendingFileId: null
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
  $("workbookTitle").textContent = workbook.fileName || "untitled";
  renderSheetTabs();
  renderGrid();
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
  const table = $("grid");
  table.innerHTML = "";
  const wb = state.workbook;
  if (!wb) return;
  const sheet = wb.sheets[wb.active] || wb.sheets[0];
  // Always render a comfortable viewport, up to the preview caps.
  const rows = Math.max(40, Math.min(sheet.rows || 1, GRID_MAX_ROWS));
  const cols = Math.max(12, Math.min(sheet.cols || 1, GRID_MAX_COLS));

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
      const cell = sheet.cells[address];
      if (cell) {
        if (cell.v !== undefined && cell.v !== null) {
          td.textContent = String(cell.v);
          if (typeof cell.v === "number") td.classList.add("num");
        } else if (cell.f) {
          td.textContent = cell.f;
          td.classList.add("fonly");
        }
      }
      if (address === state.selected) td.classList.add("sel");
      td.addEventListener("click", () => selectCell(address, sheet.cells[address]));
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  const shown = Object.keys(sheet.cells || {}).length;
  $("gridStatus").textContent = `${sheet.name}: ${shown} cell${shown === 1 ? "" : "s"} loaded`
    + (shown >= 20000 ? " (preview capped)" : "");
  selectCell(state.selected, (sheet.cells || {})[state.selected]);
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

async function submitRun(promptText) {
  const prompt = (promptText !== undefined ? promptText : $("prompt").value).trim();
  if (!prompt) {
    setStatus("Enter a prompt first.");
    return;
  }
  $("send").disabled = true;
  addMessage("user", prompt);
  if (promptText !== undefined) $("prompt").value = "";
  const statusMessage = addMessage("sys", "Submitting…");
  try {
    await ensureToken();
    let initFile;
    const file = $("file").files && $("file").files[0];
    if (file) {
      statusMessage.textContent = "Uploading attachment…";
      const uploaded = await uploadFile(file);
      initFile = uploaded.fileId;
      $("file").value = "";
    } else if (state.pendingFileId) {
      initFile = state.pendingFileId;
    }
    const body = {
      prompt,
      mode: $("mode").value === "ask" ? "ask" : "action"
    };
    if (initFile) body.initFile = initFile;
    if ($("model").value) body.model = $("model").value;
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
    if (!run) return;
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
  } catch (error) {
    statusMessage.className = "msg err";
    statusMessage.textContent = `ERROR: ${error.message || error}`;
  } finally {
    $("send").disabled = false;
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
