/* Read-only shortcut inventory and delivery proof shared by both surfaces. */
(function (g) {
  "use strict";
  function signature(key) { return String(key || "").toUpperCase().split(/[+\s]+/).filter(Boolean).sort().join(" "); }
  function createMonitor(profile, runtime, surface) {
    var active = new Set(runtime.commands(profile).map(function (c) { return c.macro; }));
    var delivered = new Map(), listeners = [], testing = false, host = null;
    function changed() { listeners.forEach(function (fn) { fn(); }); }
    return {
      subscribe: function (fn) { listeners.push(fn); },
      test: function (enabled) { testing = Boolean(enabled); changed(); },
      testing: function () { return testing; },
      host: function (snapshot) { host = snapshot; changed(); },
      observe: function (macro, source) {
        if (!active.has(macro)) return false;
        var sources = delivered.get(macro) || {};
        sources[source] = (sources[source] || 0) + 1;
        delivered.set(macro, sources); changed();
        return testing;
      },
      rows: function () {
        return profile.commands.filter(function (c) { return c.host === "Excel"; }).map(function (c) {
          var implemented = active.has(c.macro), key = implemented ? runtime.shortcutKey(profile, c) : null;
          var assignment = "Not implemented", warning = "";
          if (implemented && !key) assignment = "Button only; no shared binding";
          else if (implemented && surface === "web") assignment = "Grid binding; browser ownership unverified";
          else if (implemented) {
            assignment = !host ? "Worksheet assignment not checked" : host.error ? "Assignment unknown: " + host.error :
              !host.supported ? "Worksheet shortcut API unavailable; pane keys can be tested" :
              !Object.prototype.hasOwnProperty.call(host.shortcuts, "SXLProfile_" + c.macro) ? "Missing from loaded Excel profile; reload the add-in" :
              host.shortcuts["SXLProfile_" + c.macro] === null ? "Disabled by an Excel/add-in conflict choice" :
              signature(host.shortcuts["SXLProfile_" + c.macro]) === signature(key) ? "Assigned in Excel; delivery unverified" :
              "Excel uses " + host.shortcuts["SXLProfile_" + c.macro] + "; differs from shared key";
            if (host && host.conflicts && host.conflicts.some(function (row) { return row.inUse && signature(row.shortcut) === signature(key); })) warning = "Key is reported in use; this alone does not identify its owner.";
          }
          var sources = delivered.get(c.macro);
          return { macro: c.macro, imported: c.key || "Unbound", shared: key || "—", assignment: assignment, warning: warning,
            delivery: sources ? Object.keys(sources).map(function (s) { return s + ": " + sources[s]; }).join(", ") : "Not observed this session" };
        });
      }
    };
  }
  async function inspectOffice(office, keys, timeoutMs) {
    var timer;
    try {
      if (!office || !office.context || !office.context.requirements ||
          !office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1") ||
          !office.context.requirements.isSetSupported("SharedRuntime", "1.1") ||
          !office.actions || typeof office.actions.getShortcuts !== "function") return { supported: false };
      return await Promise.race([
        (async function () {
          var shortcuts = await office.actions.getShortcuts();
          if (!shortcuts || typeof shortcuts !== "object" || Array.isArray(shortcuts)) throw new Error("Excel returned an invalid shortcut profile");
          var conflicts = [], conflictError;
          if (typeof office.actions.areShortcutsInUse === "function") {
            try { conflicts = await office.actions.areShortcutsInUse(keys.map(function (k) { return k.replace(/ /g, "+"); })); }
            catch (e) { conflictError = "Conflict check failed: " + (e.message || e); }
          } else conflictError = "Conflict API unavailable";
          return { supported: true, shortcuts: shortcuts, conflicts: Array.isArray(conflicts) ? conflicts : [], conflictError: conflictError };
        })(),
        new Promise(function (_, reject) { timer = setTimeout(function () { reject(new Error("Excel shortcut query timed out; assignment remains unknown")); }, timeoutMs || 5000); })
      ]);
    } catch (e) { return { supported: true, error: String(e.message || e) }; }
    finally { clearTimeout(timer); }
  }
  function mount(root, monitor, profile, runtime, office) {
    var doc = root.ownerDocument, generation = 0;
    function element(tag, text, parent) { var el = doc.createElement(tag); if (text) el.textContent = text; if (parent) parent.appendChild(el); return el; }
    root.className = "shortcut-diagnostics";
    var summary = element("summary", "Shortcut diagnostics", root);
    var description = element("p", "Imported settings, shared bindings and actual delivery are separate. Counts below are observations in this session, not proof of a saved workbook change.", root);
    var label = element("label", "Find a command or key ", root), search = element("input", "", label);
    search.type = "search"; search.setAttribute("aria-label", "Find a command or key");
    var controls = element("div", "", root), test = element("button", "Start delivery test", controls), stop = element("button", "Stop delivery test", controls);
    var refresh = office ? element("button", "Check Excel assignments", controls) : null;
    var status = element("p", "", root); status.setAttribute("role", "status");
    var wrap = element("div", "", root); wrap.className = "shortcut-table";
    var table = element("table", "", wrap), head = element("tr", "", element("thead", "", table));
    var columns = ["Command", "Imported key", "Shared key", "Assignment", "Observed delivery"];
    columns.forEach(function (s) { element("th", s, head).scope = "col"; });
    var body = element("tbody", "", table);
    function render() {
      body.replaceChildren();
      var query = search.value.toLowerCase();
      monitor.rows().filter(function (r) { return [r.macro, r.imported, r.shared, r.assignment].join(" ").toLowerCase().includes(query); }).forEach(function (r) {
        var row = element("tr", "", body);
        [r.macro, r.imported, r.shared, r.assignment + (r.warning ? " · " + r.warning : ""), r.delivery].forEach(function (s, i) { element("td", s, row).setAttribute("data-label", columns[i]); });
      });
      stop.disabled = !monitor.testing(); test.disabled = monitor.testing();
      status.textContent = monitor.testing() ? "Delivery test active: supported shortcuts are observed only; formatting is not staged. Focus the grid or Excel worksheet, then press a shared key. Stop the test to resume commands." :
        "Commands run normally. Start a delivery test to check keys without staging formatting. Text editors, IME and AltGr keep their normal input.";
    }
    test.addEventListener("click", function () { monitor.test(true); });
    stop.addEventListener("click", function () { monitor.test(false); });
    search.addEventListener("input", render);
    root.addEventListener("toggle", function () { if (!root.open) monitor.test(false); });
    if (doc.addEventListener) doc.addEventListener("visibilitychange", function () { if (doc.hidden) monitor.test(false); });
    if (office && office.onReady) office.onReady(function () {
      if (office.addin && typeof office.addin.onVisibilityModeChanged === "function") {
        Promise.resolve().then(function () { return office.addin.onVisibilityModeChanged(function (message) {
          if (message.visibilityMode !== "Taskpane") monitor.test(false);
        }); }).catch(function () { status.textContent += " Pane visibility could not be tracked; stop the delivery test before hiding the pane."; });
      }
    });
    async function check() {
      if (!refresh) return;
      var request = ++generation; refresh.disabled = true; refresh.textContent = "Checking Excel assignments…";
      var keys = runtime.commands(profile).map(function (c) { return runtime.shortcutKey(profile, c); }).filter(Boolean);
      var snapshot = await inspectOffice(office, keys);
      if (request !== generation) return;
      monitor.host(snapshot); refresh.disabled = false; refresh.textContent = "Check Excel assignments";
      if (snapshot.conflictError) status.textContent += " " + snapshot.conflictError;
    }
    monitor.subscribe(render); render();
    if (refresh) refresh.addEventListener("click", check);
    return { check: check };
  }
  var api = { signature: signature, createMonitor: createMonitor, inspectOffice: inspectOffice, mount: mount };
  if (typeof module === "object" && module.exports) module.exports = api; else g.SXLShortcutDiagnostics = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
