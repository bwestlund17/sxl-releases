/* Shared deterministic profile dispatch. No host writes or DOM dependencies. */
(function (g) {
  "use strict";
  function commands(profile) {
    return profile.commands.filter(function (c) { return c.host === "Excel" && profile.cycles[c.macro] && profile.cycles[c.macro].length; });
  }
  function supportedOfficeKey(key) { return /^(?:(?:Ctrl|Alt|Shift) )+[A-Z0-9]$/i.test(key); }
  // Excel accepts the legacy unquoted L prefix in this installed LIBOR format;
  // SSF needs it quoted. This affects display only, never the stored format.
  function previewFormat(format) { return String(format).replace(/(^|;)L(?=[+-]0)/g, '$1"L"'); }
  function match(profile, event) {
    if (event.repeat || event.isComposing || event.metaKey) return null;
    var target = event.target || {};
    if (/^(INPUT|TEXTAREA|SELECT)$/i.test(target.tagName || "") || target.isContentEditable) return null;
    return commands(profile).find(function (command) {
      // Both hosts expose buttons for keys outside Office's portable grammar.
      if (!supportedOfficeKey(command.key)) return false;
      var parts = command.key.split(" ");
      var char = parts.pop().toUpperCase();
      var code = /^Digit[0-9]$/.test(event.code || "") ? event.code.slice(5) : /^Key[A-Z]$/.test(event.code || "") ? event.code.slice(3) : String(event.key).toUpperCase();
      return code === char && Boolean(event.ctrlKey) === parts.includes("Ctrl") && Boolean(event.shiftKey) === parts.includes("Shift") && Boolean(event.altKey) === parts.includes("Alt");
    }) || null;
  }
  function createCycle(profile) {
    var last = null;
    var next = function (macro, selection, current) {
      var entries = profile.cycles[macro];
      if (!entries || !entries.length) throw new Error("Unsupported imported command: " + macro);
      var index;
      var field = Object.keys(entries[0].patch)[0];
      if (last && last.macro === macro && last.selection === selection && (!current || current[field] === undefined)) index = (last.index + 1) % entries.length;
      else {
        // A repeated cycle resumes from the cell's existing format when known.
        var existing = entries.findIndex(function (entry) { return Object.keys(entry.patch).every(function (field) { return current && current[field] === entry.patch[field]; }); });
        index = (existing + 1) % entries.length;
      }
      last = { macro: macro, selection: selection, index: index };
      return { index: index, label: entries[index].label, patch: Object.assign({}, entries[index].patch) };
    };
    next.reset = function () { last = null; };
    return next;
  }
  var api = { commands: commands, match: match, createCycle: createCycle, supportedOfficeKey: supportedOfficeKey, previewFormat: previewFormat };
  if (typeof module === "object" && module.exports) module.exports = api;
  else g.SXLProfileRuntime = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
