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
  // ---- A1 reference rebasing and Fast Fill --------------------------------
  // Filling a selection down or right copies the first row/column and rebases
  // its RELATIVE references. `$` anchors hold, quoted sheet names and string
  // literals are copied verbatim, and a reference that would leave the grid is
  // refused instead of becoming #REF! (a silent #REF! would be an unaudited
  // corruption of the model's lineage).
  var GRID_ROWS = 1048576, GRID_COLS = 16384;
  var FILL_MACROS = { ProCopyRight: "right", ProCopyDown: "down" };
  var FILL_LIMIT = 50;
  function fillDirection(macro) { return FILL_MACROS[macro] || null; }
  function isFill(macro) { return Object.prototype.hasOwnProperty.call(FILL_MACROS, macro); }
  function columnLetters(index) {
    var label = "", n = Number(index);
    if (!Number.isInteger(n) || n < 1 || n > GRID_COLS) throw new Error("Column " + index + " is outside the worksheet.");
    while (n > 0) { var r = (n - 1) % 26; label = String.fromCharCode(65 + r) + label; n = Math.floor((n - 1) / 26); }
    return label;
  }
  function columnNumber(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
    return n;
  }
  function refuseOutside(text) { throw new Error("Fill would move " + text + " outside the worksheet."); }
  // Shift one reference token; `kind` is cell, column or row.
  function shiftReference(kind, parts, dRow, dCol) {
    var text = parts.text;
    if (kind === "cell") {
      var col = parts.colAnchor ? columnNumber(parts.col) : columnNumber(parts.col) + dCol;
      var row = parts.rowAnchor ? Number(parts.row) : Number(parts.row) + dRow;
      if (col < 1 || col > GRID_COLS || row < 1 || row > GRID_ROWS) refuseOutside(text);
      return (parts.colAnchor ? "$" : "") + columnLetters(col) + (parts.rowAnchor ? "$" : "") + row;
    }
    var shifted = null;
    if (kind === "column") {
      shifted = [parts.first, parts.second].map(function (part) {
        var value = part.anchor ? columnNumber(part.value) : columnNumber(part.value) + dCol;
        if (value < 1 || value > GRID_COLS) refuseOutside(text);
        return (part.anchor ? "$" : "") + columnLetters(value);
      });
      return shifted.join(":");
    }
    shifted = [parts.first, parts.second].map(function (part) {
      var value = part.anchor ? Number(part.value) : Number(part.value) + dRow;
      if (value < 1 || value > GRID_ROWS) refuseOutside(text);
      return (part.anchor ? "$" : "") + value;
    });
    return shifted.join(":");
  }
  var IDENTIFIER = /[A-Za-z0-9_.]/;
  // Try to read one reference token at the start of `rest`.
  function readReference(rest) {
    var m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_])/.exec(rest);
    if (m) return { kind: "column", text: m[0], first: { anchor: m[1] === "$", value: m[2] }, second: { anchor: m[3] === "$", value: m[4] } };
    m = /^(\$?)([1-9]\d{0,6}):(\$?)([1-9]\d{0,6})(?![A-Za-z0-9_.])/.exec(rest);
    if (m) return { kind: "row", text: m[0], first: { anchor: m[1] === "$", value: m[2] }, second: { anchor: m[3] === "$", value: m[4] } };
    m = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?![A-Za-z0-9_(.])/.exec(rest);
    if (m) return { kind: "cell", text: m[0], colAnchor: m[1] === "$", col: m[2], rowAnchor: m[3] === "$", row: m[4] };
    return null;
  }
  function shiftFormulaReferences(formula, dRow, dCol) {
    if (typeof formula !== "string" || (!dRow && !dCol)) return formula;
    var out = "", i = 0;
    while (i < formula.length) {
      var ch = formula[i];
      // Double-quoted strings and single-quoted sheet names travel unchanged.
      if (ch === '"' || ch === "'") {
        var end = i + 1;
        while (end < formula.length) {
          if (formula[end] === ch) { if (formula[end + 1] === ch) { end += 2; continue; } break; }
          end++;
        }
        out += formula.slice(i, end + 1); i = end + 1; continue;
      }
      if (i > 0 && IDENTIFIER.test(formula[i - 1])) { out += ch; i++; continue; }
      var token = readReference(formula.slice(i));
      if (!token) { out += ch; i++; continue; }
      out += shiftReference(token.kind, token, dRow, dCol); i += token.text.length;
    }
    return out;
  }
  // Fill a selection from its first row (down) or first column (right).
  // `lookup(address)` returns { value, formula, format }; the caller owns where
  // that snapshot comes from (committed grid + staged drafts, or live Excel).
  // Only values, formulas and number formats are copied: visual styles stay
  // with the dedicated colour/format commands so both surfaces behave alike.
  function planFill(direction, rect, lookup, limit) {
    limit = limit || FILL_LIMIT;
    if (direction !== "down" && direction !== "right") throw new Error("Fill direction must be down or right.");
    var top = Number(rect && rect.top), bottom = Number(rect && rect.bottom);
    var left = Number(rect && rect.left), right = Number(rect && rect.right);
    if (![top, bottom, left, right].every(function (v) { return Number.isInteger(v) && v >= 1; })) throw new Error("Select a range to fill.");
    if (direction === "down" && bottom <= top) throw new Error("Select the source row plus the rows to fill.");
    if (direction === "right" && right <= left) throw new Error("Select the source column plus the columns to fill.");
    var targets = [], row, col;
    if (direction === "down") {
      for (row = top + 1; row <= bottom; row++) for (col = left; col <= right; col++) targets.push({ address: columnLetters(col) + row, from: columnLetters(col) + top, dRow: row - top, dCol: 0 });
    } else {
      for (col = left + 1; col <= right; col++) for (row = top; row <= bottom; row++) targets.push({ address: columnLetters(col) + row, from: columnLetters(left) + row, dRow: 0, dCol: col - left });
    }
    if (targets.length > limit) throw new Error("Fill would write " + targets.length + " cells (max " + limit + ").");
    var edits = [], unchanged = 0, source = targets[0].from;
    targets.forEach(function (target) {
      var from = lookup(target.from) || {};
      var current = lookup(target.address) || {};
      var edit = { address: target.address };
      if (typeof from.formula === "string" && from.formula.charAt(0) === "=") edit.formula = shiftFormulaReferences(from.formula, target.dRow, target.dCol);
      else {
        if (typeof from.value === "boolean") throw new Error("Fast Fill cannot copy the boolean in " + target.from + ": the audited value path writes numbers or text.");
        edit.value = from.value === undefined || from.value === null ? "" : String(from.value);
      }
      // Only a format that actually differs travels with the fill: writing the
      // target's existing format would add a redundant mutation attribute, and
      // in the Office-free lane ANY extra key makes the edit ineligible.
      var currentFormat = current.format || "";
      if (typeof from.format === "string" && from.format && from.format !== currentFormat) edit.format = from.format;
      // A blank target filled from a blank source is not a change; without this
      // every repeat press would stage a no-op clear on each empty cell.
      var currentIsFormula = typeof current.formula === "string" && current.formula.charAt(0) === "=";
      var currentText = currentIsFormula ? current.formula : (current.value === undefined || current.value === null ? "" : String(current.value));
      var editText = edit.formula !== undefined ? edit.formula : edit.value;
      if (currentText === editText && currentIsFormula === (edit.formula !== undefined)) unchanged++;
      else edits.push(edit);
    });
    return { label: "Fast Fill " + (direction === "down" ? "Down" : "Right"), source: source, targets: targets.length, unchanged: unchanged, edits: edits };
  }
  // ---- Sign flip and Wall Street auto-colour ------------------------------
  var AUTO_COLOR = { input: "#0000FF", formula: "#000000", link: "#008000" };
  function sameColor(a, b) {
    var norm = function (value) { return typeof value === "string" ? "#" + value.replace(/^#/, "").toUpperCase() : ""; };
    return norm(a) === norm(b);
  }
  // Flip the sign of every numeric cell and formula in the selection. Text,
  // blanks and errors are left alone rather than coerced.
  function planFlipSign(cells) {
    var edits = [], skipped = [];
    (cells || []).forEach(function (cell) {
      var address = cell && cell.address;
      if (!address) return;
      var formula = typeof cell.formula === "string" && cell.formula.charAt(0) === "=" ? cell.formula : null;
      if (formula) {
        edits.push({ address: address, formula: "=-(" + formula.slice(1) + ")" });
        return;
      }
      var value = cell.value;
      if (typeof value === "number" && Number.isFinite(value)) {
        edits.push({ address: address, value: String(value === 0 ? 0 : -value) });
        return;
      }
      skipped.push(address);
    });
    return { label: "Flip sign", targets: (cells || []).length, skipped: skipped, edits: edits };
  }
  // Flatten replaces formulas with the value they last calculated to. It is
  // deliberately destructive to lineage, so it refuses rather than guesses: a
  // formula with no calculated value (never opened in Excel, or a web preview
  // with no cached result) and an error value both abort the whole selection.
  // The ledger then classifies each cell as a formula overwritten by a
  // hardcode, which is exactly the warning a reviewer needs to see.
  function planFlatten(cells) {
    var edits = [], skipped = [];
    (cells || []).forEach(function (cell) {
      var address = cell && cell.address;
      if (!address) return;
      var formula = typeof cell.formula === "string" && cell.formula.charAt(0) === "=" ? cell.formula : null;
      if (!formula) { skipped.push(address); return; }
      var value = cell.value;
      if (value === undefined || value === null || value === "") {
        throw new Error("Flatten needs the calculated value of " + address + ". Recalculate the workbook (or open it in Excel) before flattening.");
      }
      if (typeof value === "boolean") throw new Error("Flatten cannot write the boolean result of " + address + "; the audited value path writes numbers or text.");
      var text = String(value);
      if (/^#(REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM|GETTING_DATA)/.test(text)) {
        throw new Error("Flatten refuses the error value " + text + " in " + address + ".");
      }
      if (text === formula) { skipped.push(address); return; }
      edits.push({ address: address, value: text });
    });
    return { label: "Flatten", targets: (cells || []).length, skipped: skipped, edits: edits };
  }
  // Wall Street convention: hard-coded inputs blue, formulas black, links to
  // another sheet green. Only cells whose colour would actually change travel.
  function planAutoColor(cells) {
    var edits = [], skipped = [];
    (cells || []).forEach(function (cell) {
      var address = cell && cell.address;
      if (!address) return;
      var formula = typeof cell.formula === "string" && cell.formula.charAt(0) === "=" ? cell.formula : null;
      var color;
      if (formula) color = formula.indexOf("!") >= 0 ? AUTO_COLOR.link : AUTO_COLOR.formula;
      else if (typeof cell.value === "number" || typeof cell.value === "boolean") color = AUTO_COLOR.input;
      else if (typeof cell.value === "string" && cell.value !== "") color = AUTO_COLOR.formula;
      else { skipped.push(address); return; }
      if (sameColor(cell.fontColor, color)) return;
      edits.push({ address: address, fontColor: color });
    });
    return { label: "AutoColor selection", targets: (cells || []).length, skipped: skipped, edits: edits };
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
  // Capture each gesture's target immediately, then read/commit in order.
  // A lost write response must stop the queue rather than replay a command.
  function createCommandQueue() {
    var tail = Promise.resolve(), pending = 0, blocked = null;
    return {
      busy: function () { return pending > 0; },
      blocked: function () { return blocked; },
      halt: function (reason) { blocked = reason; },
      enqueue: function (capture, execute) {
        if (blocked) return Promise.reject(new Error(blocked));
        if (pending >= 50) return Promise.reject(new Error("Wait for queued shortcuts to finish before pressing more keys."));
        pending++;
        var captured;
        try { captured = Promise.resolve(capture()); }
        catch (error) { captured = Promise.reject(error); }
        // Capture may fail before its queue slot starts; keep that rejection
        // handled now and deliver it to the caller when the slot is reached.
        captured.catch(function () {});
        var result = tail.then(function () {
          if (blocked) throw new Error(blocked);
          return captured.then(execute);
        }).finally(function () { pending--; });
        tail = result.catch(function () {});
        return result;
      }
    };
  }
  var api = { commands: commands, match: match, createCycle: createCycle, createCommandQueue: createCommandQueue, isDecimal: isDecimal, decimalFormat: decimalFormat, supportedOfficeKey: supportedOfficeKey, shortcutKey: shortcutKey, keyLabel: keyLabel, previewFormat: previewFormat, planFlipSign: planFlipSign, planAutoColor: planAutoColor, planFlatten: planFlatten, AUTO_COLOR: AUTO_COLOR, isFill: isFill, fillDirection: fillDirection, columnLetters: columnLetters, shiftFormulaReferences: shiftFormulaReferences, planFill: planFill };
  if (typeof module === "object" && module.exports) module.exports = api;
  else g.SXLProfileRuntime = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
