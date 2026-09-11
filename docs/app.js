"use strict";

const $ = (id) => document.getElementById(id);
const state = { token: null, runId: null, polling: false };

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

function setStatus(message) {
  $("status").textContent = message;
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
  let response = await fetch(path, Object.assign({}, options, { headers }));
  if (response.status === 401) {
    await login();
    response = await fetch(path, Object.assign({}, options, {
      headers: Object.assign({}, options.headers || {}, { authorization: `Bearer ${state.token}` })
    }));
  }
  return response;
}

async function uploadFile(file) {
  const form = new FormData();
  form.append("file", file);
  const response = await authFetch(api("/api/spreadsheets/upload"), { method: "POST", body: form });
  if (!response.ok) throw new Error(`upload failed: HTTP ${response.status}`);
  return response.json();
}

function renderDownloads(run) {
  const target = $("downloads");
  target.innerHTML = "";
  const addDownload = (label, url) => {
    const button = document.createElement("button");
    button.textContent = label;
    button.className = "secondary";
    button.style.marginRight = "8px";
    button.addEventListener("click", () => downloadUrl(url));
    target.appendChild(button);
  };
  if (run.downloadUrl) addDownload("Download workbook", api(run.downloadUrl));
  for (const artifact of run.artifacts || []) {
    addDownload(`Download ${artifact}`, api(`/api/spreadsheets/${run.runId}/artifacts/${encodeURIComponent(artifact)}`));
  }
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

function renderRun(run) {
  state.runId = run.runId;
  const lines = [
    `runId: ${run.runId}`,
    `status: ${run.status}`,
    `mode: ${run.mode}`,
    run.prompt ? `prompt: ${run.prompt}` : ""
  ].filter(Boolean);
  if (run.summary) lines.push("", "summary:", String(run.summary));
  if (run.error) lines.push("", `error: ${run.error}`);
  $("output").textContent = lines.join("\n");
  renderDownloads(run);
}

async function pollRun(runId) {
  state.polling = true;
  while (state.polling && state.runId === runId) {
    const response = await authFetch(api(`/api/spreadsheets/${runId}`));
    if (!response.ok) {
      setStatus(`status failed: HTTP ${response.status}`);
      state.polling = false;
      return;
    }
    const run = await response.json();
    renderRun(run);
    setStatus(`Run ${runId.slice(0, 8)}: ${run.status}`);
    if (["completed", "failed", "cancelled"].includes(run.status)) {
      state.polling = false;
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  state.polling = false;
}

async function submitRun() {
  const prompt = $("prompt").value.trim();
  if (!prompt) {
    setStatus("Enter a prompt first.");
    return;
  }
  $("submit").disabled = true;
  try {
    await ensureToken();
    setStatus("Preparing input...");
    let initFile;
    const file = $("file").files && $("file").files[0];
    if (file) {
      const uploaded = await uploadFile(file);
      initFile = uploaded.fileId;
    }
    setStatus("Submitting run...");
    const body = {
      prompt,
      mode: $("mode").value === "ask" ? "ask" : "action"
    };
    if (initFile) body.initFile = initFile;
    if ($("model").value.trim()) body.model = $("model").value.trim();
    const response = await authFetch(api("/api/spreadsheets"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(body)
    });
    const submitted = await response.json();
    if (!response.ok) throw new Error(submitted.error || `HTTP ${response.status}`);
    renderRun({
      runId: submitted.runId,
      status: submitted.status || "queued",
      mode: body.mode,
      prompt
    });
    setStatus(`Queued ${submitted.runId}`);
    await pollRun(submitted.runId);
    await loadRuns();
  } catch (error) {
    setStatus(`ERROR: ${error.message || error}`);
  } finally {
    $("submit").disabled = false;
  }
}

async function loadRuns() {
  try {
    const response = await authFetch(api("/api/spreadsheets/runs"));
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const body = await response.json();
    const list = $("runs");
    list.innerHTML = "";
    for (const run of body.runs || []) {
      const item = document.createElement("li");
      const link = document.createElement("a");
      link.textContent = `${run.runId.slice(0, 8)} · ${run.status} · ${String(run.prompt || "").slice(0, 70)}`;
      link.addEventListener("click", async () => {
        const statusResponse = await authFetch(api(`/api/spreadsheets/${run.runId}`));
        if (statusResponse.ok) renderRun(await statusResponse.json());
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

window.addEventListener("DOMContentLoaded", () => {
  $("submit").addEventListener("click", submitRun);
  $("refresh").addEventListener("click", loadRuns);
  loadRuns();
});
