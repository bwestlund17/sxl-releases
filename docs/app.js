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

// ---- Credit billing (Stripe Checkout) -------------------------------------
function setBillingStatus(message) {
  $("billingStatus").textContent = message;
}

async function buyPack(packId) {
  setBillingStatus("Starting checkout…");
  try {
    const response = await authFetch(api("/v1/billing/checkout"), {
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

window.addEventListener("DOMContentLoaded", async () => {
  $("submit").addEventListener("click", submitRun);
  $("refresh").addEventListener("click", loadRuns);
  if (AUTH_ENABLED) {
    await setupAuth();
    await handleAuthRedirect();
  }
  loadBilling();
  loadRuns();
});
