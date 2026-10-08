/* Shared deterministic profile dispatch. No host writes or DOM dependencies. */
(function (g) {
  "use strict";
  function commands(profile) {
    return profile.commands.filter(function (c) { return c.host === "Excel" && ((profile.cycles[c.macro] && profile.cycles[c.macro].length) || (profile.operations || []).includes(c.macro)); });
  }
  function isDecimal(macro) { return macro === "IncrDecimal" || macro === "DecrDecimal"; }
  function decimalFormat(format, value, direction) {
    format = format || "General";
    if (/^general$/i.test(format)) {
      if (typeof value !== "number" || !Number.isFinite(value)) return format;
      var parts = String(Math.abs(value)).toLowerCase().split("e");
      var decimals = Math.max(0, (parts[0].split(".")[1] || "").length - Number(parts[1] || 0));
      decimals = Math.max(0, decimals + direction);
      if (decimals > 30) throw new Error("Excel decimal precision is limited to 30 places.");
      return "0" + (decimals ? "." + "0".repeat(decimals) : "");
    }
    // Mask literal and directive tokens at their original offsets. Semicolons,
    // dots and digits inside quotes, escapes, colors/conditions, spacing and
    // fill directives must never be mistaken for a numeric placeholder.
    var mask = "", quote = false, bracket = false, elapsed = [];
    for (var i = 0; i < format.length; i++) {
      var ch = format[i];
      if (ch === '"' && !bracket) { quote = !quote; mask += " "; }
      else if (quote) mask += " ";
      else if (ch === "[") { if (/^\[[hms]+\]/i.test(format.slice(i))) elapsed.push(i); bracket = true; mask += " "; }
      else if (bracket) { if (ch === "]") bracket = false; mask += " "; }
      else if (ch === "\\" || ch === "_" || ch === "*") { mask += " "; if (i + 1 < format.length) { mask += " "; i++; } }
      else mask += ch;
    }
    if (quote || bracket) throw new Error("Cannot adjust an unterminated number format.");
    var start = 0, sections = [];
    for (var end = 0; end <= format.length; end++) {
      if (end < format.length && mask[end] !== ";") continue;
      var raw = format.slice(start, end), visible = mask.slice(start, end);
      // Date/time, fraction and text sections retain their complete format.
      if (/[ymdh\/]/i.test(visible) || /(^|[^a-z])s{1,2}([^a-z]|$)/i.test(visible) || elapsed.some(function (index) { return index >= start && index < end; })) sections.push(raw);
      else {
        var exponent = /E[+-]?[0#?]+/i.exec(visible);
        var limit = exponent ? exponent.index : visible.length;
        var numeric = /[0#?]+(?:,[0#?]+)*(?:\.[0#?]*)?/.exec(visible.slice(0, limit));
        if (!numeric) sections.push(raw);
        else {
          var token = numeric[0], dot = token.indexOf("."), precision = dot < 0 ? 0 : token.length - dot - 1;
          if (direction > 0) {
            if (precision >= 30) throw new Error("Excel decimal precision is limited to 30 places.");
            token += (dot < 0 ? "." : "") + "0";
          } else if (precision > 0) {
            token = token.slice(0, -1); if (token.endsWith(".")) token = token.slice(0, -1);
          }
          sections.push(raw.slice(0, numeric.index) + token + raw.slice(numeric.index + numeric[0].length));
        }
      }
      start = end + 1;
    }
    var result = sections.join(";");
    if (result.length > 255) throw new Error("The adjusted number format exceeds Excel's 255-character limit.");
    return result;
  }
  function supportedOfficeKey(key) { return /^(?:(?:Ctrl|Alt|Shift) )+[A-Z0-9]$/i.test(key); }
  function shortcutKey(profile, command) {
    var key = command.portableKey || command.key;
    if (!supportedOfficeKey(key)) return null;
    // An alternate must never steal an imported command, including one whose
    // implementation is still pending. Modifier order/case don't change a chord.
    var signature = key.toUpperCase().split(/\s+/).sort().join(" ");
    if (command.portableKey && profile.commands.some(function (other) {
      return other !== command && other.host === "Excel" &&
        [other.key, other.portableKey].some(function (candidate) {
          return candidate && candidate.toUpperCase().split(/\s+/).sort().join(" ") === signature;
        });
    })) return null;
    return key;
  }
  function keyLabel(profile, command) {
    var key = shortcutKey(profile, command);
    if (!key) return "button only; imported " + (command.key || "unbound");
    return key + (command.portableKey ? "; imported " + command.key : "");
  }
  // Excel accepts the legacy unquoted L prefix in this installed LIBOR format;
  // SSF needs it quoted. This affects display only, never the stored format.
  function previewFormat(format) { return String(format).replace(/(^|;)L(?=[+-]0)/g, '$1"L"'); }
  function match(profile, event) {
    if (event.repeat || event.isComposing || event.metaKey || (event.getModifierState && event.getModifierState("AltGraph"))) return null;
    var target = event.target || {};
    if (/^(INPUT|TEXTAREA|SELECT)$/i.test(target.tagName || "") || target.isContentEditable) return null;
    return commands(profile).find(function (command) {
      var key = shortcutKey(profile, command);
      if (!key) return false;
      var parts = key.split(" ");
      var char = parts.pop().toUpperCase();
      var code = /^Digit[0-9]$/.test(event.code || "") ? event.code.slice(5) : /^Key[A-Z]$/.test(event.code || "") ? event.code.slice(3) : String(event.key).toUpperCase();
      return code === char && Boolean(event.ctrlKey) === parts.includes("Ctrl") && Boolean(event.shiftKey) === parts.includes("Shift") && Boolean(event.altKey) === parts.includes("Alt");
    }) || null;
  }
  function createCycle(profile) {
    var last = null;
    var next = function (macro, selection, current) {
      if (isDecimal(macro) && (profile.operations || []).includes(macro)) {
        var format = decimalFormat(current && current.format, current && current.value, macro === "IncrDecimal" ? 1 : -1);
        return { label: macro === "IncrDecimal" ? "Increase decimal precision" : "Decrease decimal precision", patch: { format: format } };
      }
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
    next.plan = function (macro, selection, cells) {
      if (!cells.length) throw new Error("Select cells before formatting.");
      var entry = isDecimal(macro) ? null : next(macro, selection, cells[0]);
      return { label: entry ? entry.label : "Adjust each cell's decimal precision", patches: cells.map(function (cell) {
        var patch = (entry || next(macro, selection, cell)).patch;
        return Object.assign({ address: cell.address }, isDecimal(macro) && patch.format === (cell.format || "General") ? {} : patch);
      }) };
    };
    return next;
  }
  var api = { commands: commands, match: match, createCycle: createCycle, isDecimal: isDecimal, decimalFormat: decimalFormat, supportedOfficeKey: supportedOfficeKey, shortcutKey: shortcutKey, keyLabel: keyLabel, previewFormat: previewFormat };
  if (typeof module === "object" && module.exports) module.exports = api;
  else g.SXLProfileRuntime = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
