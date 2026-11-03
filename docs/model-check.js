/* Read-only model checks shared by the web grid and the Office pane. */
(function (g) {
  "use strict";
  // A check never writes and never guesses: every finding names the cell and
  // the evidence it was derived from, and rules that need the whole sheet are
  // skipped (and reported as skipped) on a partial scan.
  var ERROR = /^#(REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM|GETTING_DATA)/;
  var BLUE = "#0000FF";
  var RULES = {
    formula_error: { severity: "critical", label: "Formula error" },
    inconsistent_formula: { severity: "warning", label: "Inconsistent formula" },
    broken_reference: { severity: "critical", label: "Broken reference" },
    external_link: { severity: "warning", label: "External link" },
    uncolored_input: { severity: "warning", label: "Input not coloured" },
    input_colored_formula: { severity: "warning", label: "Formula coloured as an input" },
    unused_input: { severity: "warning", label: "Input with no dependents" }
  };
  function columnNumber(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
    return n;
  }
  function columnLetters(index) {
    var label = "", n = index;
    while (n > 0) { var r = (n - 1) % 26; label = String.fromCharCode(65 + r) + label; n = Math.floor((n - 1) / 26); }
    return label;
  }
  function key(sheet, address) { return String(sheet).toLowerCase() + "!" + String(address).toUpperCase(); }
  function isFormula(cell) { return typeof cell.formula === "string" && cell.formula.charAt(0) === "="; }
  function sameColor(value, expected) {
    if (typeof value !== "string") return false;
    return ("#" + value.replace(/^#/, "").toUpperCase()) === expected;
  }
  // Return every A1/range reference a formula depends on, resolved against the
  // formula's own sheet unless the reference names another one.
  function precedentKeys(sheet, formula) {
    var keys = [], i = 0, text = String(formula), target = sheet;
    while (i < text.length) {
      var ch = text[i];
      if (ch === '"') {
        var close = i + 1;
        while (close < text.length) { if (text[close] === '"') { if (text[close + 1] === '"') { close += 2; continue; } break; } close++; }
        i = close + 1; continue;
      }
      if (ch === "'") {
        var end = i + 1, name = "";
        while (end < text.length) {
          if (text[end] === "'") { if (text[end + 1] === "'") { name += "'"; end += 2; continue; } break; }
          name += text[end]; end++;
        }
        if (text[end + 1] === "!") { target = name; i = end + 2; continue; }
        i = end + 1; continue;
      }
      var sheetMatch = /^([A-Za-z0-9_. ]{1,31})!/.exec(text.slice(i));
      if (sheetMatch && !/[A-Za-z0-9_.]$/.test(text.slice(0, i))) {
        target = sheetMatch[1]; i += sheetMatch[0].length; continue;
      }
      var range = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6}))?(?![A-Za-z0-9_(.])/.exec(text.slice(i));
      if (!range || (i > 0 && /[A-Za-z0-9_.$]/.test(text[i - 1]))) { i++; continue; }
      var c1 = columnNumber(range[2]), r1 = Number(range[4]);
      var c2 = range[6] ? columnNumber(range[6]) : c1, r2 = range[8] ? Number(range[8]) : r1;
      var cells = (Math.abs(c2 - c1) + 1) * (Math.abs(r2 - r1) + 1);
      if (cells > 1000) { keys.push({ sheet: target, oversized: true }); i += range[0].length; continue; }
      for (var row = Math.min(r1, r2); row <= Math.max(r1, r2); row++) {
        for (var col = Math.min(c1, c2); col <= Math.max(c1, c2); col++) keys.push({ sheet: target, address: columnLetters(col) + row });
      }
      // A sheet prefix applies to the reference that follows it only.
      target = sheet;
      i += range[0].length;
    }
    return keys;
  }
  // One-hop formula graph over the same snapshot, for Trace In/Out. Keys are
  // "sheet!A1" (lower-cased sheet), so callers can ask about any cell without
  // caring which sheet the formula lives on.
  function buildGraph(cells) {
    var precedents = new Map(), dependents = new Map(), display = new Map();
    (cells || []).forEach(function (cell) {
      if (!cell || !cell.sheet) return;
      display.set(String(cell.sheet).toLowerCase(), cell.sheet);
      if (!cell.address || !isFormula(cell)) return;
      var self = key(cell.sheet, cell.address);
      precedentKeys(cell.sheet, cell.formula).forEach(function (ref) {
        if (ref.oversized) return;
        display.set(String(ref.sheet).toLowerCase(), ref.sheet);
        var id = key(ref.sheet, ref.address);
        if (!precedents.has(self)) precedents.set(self, []);
        precedents.get(self).push(id);
        if (!dependents.has(id)) dependents.set(id, []);
        dependents.get(id).push(self);
      });
    });
    var read = function (map, sheet, address) { return (map.get(key(sheet, address)) || []).slice(); };
    return {
      precedentsOf: function (sheet, address) { return read(precedents, sheet, address); },
      dependentsOf: function (sheet, address) { return read(dependents, sheet, address); },
      // Every cell that would recalculate if this one changed, breadth-first
      // and cycle-safe. Depth is capped because a real model's dependency chains
      // are not bounded and an overlay is only useful if it stays readable.
      transitiveDependents: function (sheet, address, depth) {
        var cap = Number.isInteger(depth) && depth > 0 ? depth : 5;
        var seen = new Set([key(sheet, address)]), frontier = [key(sheet, address)], levels = [];
        for (var step = 0; step < cap && frontier.length; step++) {
          var next = [];
          frontier.forEach(function (id) {
            (dependents.get(id) || []).forEach(function (child) {
              if (seen.has(child)) return;
              seen.add(child); next.push(child); levels.push({ id: child, depth: step + 1 });
            });
          });
          frontier = next;
        }
        return levels;
      },
      // "model!A1" back to {sheet, address}, restoring the sheet's own case.
      parse: function (id) {
        var bang = String(id).lastIndexOf("!");
        var sheet = String(id).slice(0, bang);
        return { sheet: display.get(sheet.toLowerCase()) || sheet, address: String(id).slice(bang + 1) };
      }
    };
  }
  // Canonical "same shape" signature for a formula: every relative reference
  // becomes its offset from the cell that owns the formula, absolute ones keep
  // their sheet coordinate. Two cells with equal signatures are the same
  // calculation copied across, which is what the consistency rule compares.
  function relativeSignature(formula, address) {
    var match = /^([A-Z]+)(\d+)$/.exec(String(address).toUpperCase());
    if (!match) return null;
    var baseColumn = columnNumber(match[1]), baseRow = Number(match[2]);
    var text = String(formula), out = "", i = 0;
    while (i < text.length) {
      var ch = text[i];
      if (ch === '"' || ch === "'") {
        var close = i + 1;
        while (close < text.length) { if (text[close] === ch) { if (text[close + 1] === ch) { close += 2; continue; } break; } close++; }
        out += text.slice(i, close + 1); i = close + 1; continue;
      }
      if (i > 0 && /[A-Za-z0-9_.$]/.test(text[i - 1])) { out += ch; i++; continue; }
      var range = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6}))?(?![A-Za-z0-9_(.])/.exec(text.slice(i));
      if (!range) { out += ch; i++; continue; }
      var token = function (colAnchor, colText, rowAnchor, rowText) {
        var col = colAnchor ? "C" + columnNumber(colText) : "C[" + (columnNumber(colText) - baseColumn) + "]";
        var row = rowAnchor ? "R" + Number(rowText) : "R[" + (Number(rowText) - baseRow) + "]";
        return row + col;
      };
      out += token(range[1] === "$", range[2], range[3] === "$", range[4]);
      if (range[6]) out += ":" + token(range[5] === "$", range[6], range[7] === "$", range[8]);
      i += range[0].length;
    }
    return out;
  }
  function columnOf(match) { return match && match[1] ? columnNumber(match[1]) : 0; }
  // A cell whose formula differs from BOTH neighbours while those neighbours
  // agree is the classic copy/paste break. Interior-only keeps totals and
  // deliberately different row/column ends out of the report. The same check
  // runs along rows and down columns because a model is built both ways.
  function inconsistentRuns(list, findings) {
    [{ axis: "row", group: function (m) { return m[2]; }, order: function (m) { return columnOf(m); } },
     { axis: "column", group: function (m) { return m[1]; }, order: function (m) { return Number(m[2]); } }].forEach(function (run) {
      var groups = new Map();
      list.forEach(function (cell) {
        if (!isFormula(cell)) return;
        var match = /^([A-Z]+)(\d+)$/.exec(String(cell.address).toUpperCase());
        if (!match) return;
        var id = String(cell.sheet).toLowerCase() + "!" + run.group(match);
        if (!groups.has(id)) groups.set(id, []);
        groups.get(id).push({ cell: cell, rank: run.order(match) });
      });
      groups.forEach(function (entries) {
        entries.sort(function (a, b) { return a.rank - b.rank; });
        for (var index = 1; index < entries.length - 1; index++) {
          var before = entries[index - 1], current = entries[index], after = entries[index + 1];
          if (current.rank !== before.rank + 1 || after.rank !== current.rank + 1) continue;
          var previous = relativeSignature(before.cell.formula, before.cell.address);
          var mine = relativeSignature(current.cell.formula, current.cell.address);
          var next = relativeSignature(after.cell.formula, after.cell.address);
          if (!previous || !mine || !next) continue;
          if (previous === next && mine !== previous) {
            findings.push(finding("inconsistent_formula", current.cell,
              "breaks the " + run.axis + "'s formula pattern: its neighbours agree but it does not",
              current.cell.formula));
          }
        }
      });
    });
  }
  function finding(rule, cell, detail, evidence) {
    return { rule: rule, severity: RULES[rule].severity, label: RULES[rule].label, sheet: cell.sheet, address: cell.address, detail: detail, evidence: evidence };
  }
  // `cells` are {sheet, address, value, formula, format, fontColor}; `complete`
  // means the scan covered each sheet's whole used range, which is what makes
  // the dependency rule honest.
  function checkModel(cells, options) {
    options = options || {};
    var list = (cells || []).filter(function (cell) { return cell && cell.address && cell.sheet; });
    // Colour rules need each cell's own font colour; a host that cannot read
    // 5,000 of them says so instead of reporting the cell as uncoloured.
    var colors = options.colors !== false;
    var byKey = new Map(), dependents = new Map(), findings = [], skipped = [];
    list.forEach(function (cell) { byKey.set(key(cell.sheet, cell.address), cell); });
    list.forEach(function (cell) {
      if (!isFormula(cell)) return;
      precedentKeys(cell.sheet, cell.formula).forEach(function (ref) {
        if (ref.oversized) return;
        var id = key(ref.sheet, ref.address);
        if (!dependents.has(id)) dependents.set(id, new Set());
        dependents.get(id).add(key(cell.sheet, cell.address));
      });
    });
    list.forEach(function (cell) {
      var valueText = cell.value === null || cell.value === undefined ? "" : String(cell.value);
      // A #REF! in the formula is the precise cause; report that rather than
      // the generic error it produces.
      if (isFormula(cell) && cell.formula.indexOf("#REF!") >= 0) {
        findings.push(finding("broken_reference", cell, "formula contains #REF!", cell.formula));
        return;
      }
      if (ERROR.test(valueText)) {
        findings.push(finding("formula_error", cell, valueText, isFormula(cell) ? cell.formula : valueText));
        return;
      }
      if (isFormula(cell)) {
        if (/\[[^\]]+\]/.test(cell.formula)) findings.push(finding("external_link", cell, "formula points at another workbook", cell.formula));
        if (colors && sameColor(cell.fontColor, BLUE)) findings.push(finding("input_colored_formula", cell, "a formula is painted as a hard-coded input", cell.formula));
        return;
      }
      var numeric = typeof cell.value === "number" || typeof cell.value === "boolean";
      if (!numeric) return;
      if (colors && !sameColor(cell.fontColor, BLUE)) findings.push(finding("uncolored_input", cell, "hard-coded input is not blue", valueText));
      if (options.complete) {
        var id = key(cell.sheet, cell.address);
        if (!dependents.has(id)) findings.push(finding("unused_input", cell, "no scanned formula depends on this input", valueText));
      }
    });
    inconsistentRuns(list, findings);
    if (!options.complete) skipped.push("unused_input");
    if (!colors) skipped.push("uncolored_input", "input_colored_formula");
    var order = { critical: 0, warning: 1, info: 2 };
    findings.sort(function (a, b) {
      return (order[a.severity] - order[b.severity]) || a.sheet.localeCompare(b.sheet) ||
        (columnNumber(/^[A-Z]+/.exec(a.address)[0]) - columnNumber(/^[A-Z]+/.exec(b.address)[0])) ||
        (Number(/\d+$/.exec(a.address)[0]) - Number(/\d+$/.exec(b.address)[0]));
    });
    var counts = { critical: 0, warning: 0, info: 0 };
    findings.forEach(function (item) { counts[item.severity]++; });
    return {
      scanned: { cells: list.length, sheets: [...new Set(list.map(function (cell) { return cell.sheet; }))], complete: Boolean(options.complete) },
      counts: counts, findings: findings, skippedRules: skipped, rules: Object.keys(RULES)
    };
  }
  var api = { checkModel: checkModel, precedentKeys: precedentKeys, relativeSignature: relativeSignature, buildGraph: buildGraph, rules: RULES, ERROR: ERROR };
  if (typeof module === "object" && module.exports) module.exports = api; else g.SXLModelCheck = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
