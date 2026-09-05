/* HK Daily XLSM in-place editor
   Updates the uploaded .xlsm ZIP parts. Does not build a new workbook. */

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const STATUS_CODES = ["ARR","DEP","OCC","B2B","VAC","OOO","OS","SV","STB","DU","DC","VM-ARR"];
const STATUS_SET = new Set(STATUS_CODES);

const state = {
  fileName: "",
  zip: null,
  sst: [],
  sstDoc: null,
  sheets: [],          // {name, path, doc, xmlText}
  cells: {},           // sheetName -> {ref: cellMeta}
  dirty: new Map(),    // "sheet!A1" -> new value (string)
  styleByStatus: {},   // status -> style index string
};

const $ = (id) => document.getElementById(id);

function toast(msg, cls = "ok") {
  const el = $("toast");
  if (!el) return;
  el.className = "toast show" + (cls === "warn" ? " warn" : "");
  el.textContent = msg;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("show"), 4200);
}

function colRow(ref) {
  const m = String(ref).match(/^([A-Z]+)(\d+)$/);
  if (!m) return { col: "", row: 0, colNum: 0 };
  const col = m[1];
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: +m[2], colNum: n };
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function xml(str) {
  return new DOMParser().parseFromString(str, "application/xml");
}

function xmlText(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function local(el, name) {
  return el.getElementsByTagNameNS(NS, name);
}

function cellTextFromSst(idx) {
  const si = state.sst[idx];
  return si == null ? "" : si;
}

function parseSharedStrings(sstXml) {
  const doc = xml(sstXml);
  const sis = [...doc.getElementsByTagNameNS(NS, "si")];
  const out = sis.map((si) => {
    const ts = [...si.getElementsByTagNameNS(NS, "t")];
    return ts.map((t) => t.textContent || "").join("");
  });
  return { doc, list: out };
}

function parseCellValue(c, sst) {
  const t = c.getAttribute("t") || "";
  const f = c.getElementsByTagNameNS(NS, "f")[0];
  const v = c.getElementsByTagNameNS(NS, "v")[0];
  const is = c.getElementsByTagNameNS(NS, "is")[0];
  let formula = f ? f.textContent : "";
  let value = "";
  if (t === "s" && v) value = sst[Number(v.textContent)] ?? "";
  else if (t === "inlineStr" && is) {
    value = [...is.getElementsByTagNameNS(NS, "t")].map((x) => x.textContent || "").join("");
  } else if (t === "b" && v) value = v.textContent === "1" ? "TRUE" : "FALSE";
  else if (v) value = v.textContent || "";
  else if (is) value = [...is.getElementsByTagNameNS(NS, "t")].map((x) => x.textContent || "").join("");
  return { type: t, formula, value, style: c.getAttribute("s") || "" };
}

function excelSerialToDate(n) {
  const num = Number(n);
  if (!Number.isFinite(num) || num < 20000 || num > 80000) return null;
  const utc = new Date(Math.round((num - 25569) * 86400 * 1000));
  if (Number.isNaN(utc.getTime())) return null;
  return utc.toISOString().slice(0, 10);
}

function dateToSerial(yyyyMmDd) {
  const [y, m, d] = yyyyMmDd.split("-").map(Number);
  const utc = Date.UTC(y, m - 1, d);
  return Math.round(utc / 86400 / 1000 + 25569);
}

async function loadXlsm(file) {
  toast("Opening file…");
  state.fileName = file.name;
  state.dirty.clear();
  state.styleByStatus = {};
  state.cells = {};
  $("fileName").textContent = file.name;
  if ($("sideFile")) $("sideFile").textContent = file.name;

  const zip = await JSZip.loadAsync(file);
  if (!zip.file("xl/vbaProject.bin")) {
    toast("Warning: no VBA project found. File can still be edited.", "warn");
  }
  state.zip = zip;

  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const relXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const wbDoc = xml(wbXml);
  const relDoc = xml(relXml);

  const relMap = {};
  const relNodes = [
    ...relDoc.getElementsByTagNameNS(PKG_REL_NS, "Relationship"),
    ...relDoc.getElementsByTagName("Relationship"),
  ];
  for (const rel of relNodes) {
    relMap[rel.getAttribute("Id")] = rel.getAttribute("Target");
  }

  const sstFile = zip.file("xl/sharedStrings.xml");
  if (sstFile) {
    const parsed = parseSharedStrings(await sstFile.async("string"));
    state.sstDoc = parsed.doc;
    state.sst = parsed.list;
  } else {
    state.sstDoc = null;
    state.sst = [];
  }

  const sheets = [];
  for (const sh of wbDoc.getElementsByTagNameNS(NS, "sheet")) {
    const name = sh.getAttribute("name");
    const rid = sh.getAttributeNS(REL_NS, "id") || sh.getAttribute("r:id");
    let target = relMap[rid];
    if (!target) continue;
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.\//, "");
    const xmlText = await zip.file(target).async("string");
    const doc = xml(xmlText);
    sheets.push({ name, path: target, doc, xmlText });
  }
  state.sheets = sheets;

  for (const sheet of sheets) {
    const map = {};
    const cs = sheet.doc.getElementsByTagNameNS(NS, "c");
    for (const c of cs) {
      const ref = c.getAttribute("r");
      if (!ref) continue;
      const meta = parseCellValue(c, state.sst);
      meta.el = c;
      map[ref] = meta;
      const val = String(meta.value || "").trim().toUpperCase();
      if (STATUS_SET.has(val) && meta.style) {
        if (!state.styleByStatus[val]) state.styleByStatus[val] = meta.style;
      }
    }
    state.cells[sheet.name] = map;
  }

  $("saveBtn").disabled = false;
  $("resetBtn").disabled = false;
  const vba = !!state.zip.file("xl/vbaProject.bin");
  if ($("sideSub")) $("sideSub").textContent = state.sheets.length + " sheets · VBA " + (vba ? "preserved" : "not found");
  renderAll();
  toast("Workbook loaded. Same XLSM will be saved — colours and macros kept.");
}

function currentValue(sheet, ref) {
  const key = sheet + "!" + ref;
  if (state.dirty.has(key)) return state.dirty.get(key);
  return state.cells[sheet]?.[ref]?.value ?? "";
}

function setValue(sheet, ref, value) {
  const key = sheet + "!" + ref;
  const orig = state.cells[sheet]?.[ref]?.value ?? "";
  if (String(value) === String(orig)) state.dirty.delete(key);
  else state.dirty.set(key, value);
  const badge = $("dirtyCount");
  if (badge) {
    badge.textContent = state.dirty.size ? state.dirty.size + " change" + (state.dirty.size === 1 ? "" : "s") : "";
    badge.classList.toggle("hidden", !state.dirty.size);
  }
}

function sheetByName(name) {
  return state.sheets.find((s) => s.name === name);
}

function mainSheetName() {
  return state.sheets[0]?.name || "JULY- 2026";
}

/* ---------- write-back into original XML ---------- */
function applyEditsToXml() {
  const grouped = {};
  for (const [key, val] of state.dirty.entries()) {
    const i = key.indexOf("!");
    const sheet = key.slice(0, i);
    const ref = key.slice(i + 1);
    (grouped[sheet] ||= []).push({ ref, val });
  }

  for (const [sheetName, edits] of Object.entries(grouped)) {
    const sheet = sheetByName(sheetName);
    if (!sheet) continue;
    const map = state.cells[sheetName];

    for (const { ref, val } of edits) {
      let meta = map[ref];
      let el = meta?.el;
      if (!el) {
        el = ensureCell(sheet.doc, ref);
        meta = map[ref] = { type: "", formula: "", value: "", style: el.getAttribute("s") || "", el };
      }

      const upper = String(val).trim().toUpperCase();
      if (STATUS_SET.has(upper) && state.styleByStatus[upper]) {
        el.setAttribute("s", state.styleByStatus[upper]);
        meta.style = state.styleByStatus[upper];
      }

      // Keep formula if user didn't intend to replace a formula with a constant
      // If original had a formula and new value equals displayed cached value, skip.
      if (meta.formula && String(val) === String(meta.value)) continue;

      // Drop formula if user typed a new literal
      const f = el.getElementsByTagNameNS(NS, "f")[0];
      if (f && String(val) !== String(meta.formula)) f.parentNode.removeChild(f);

      // Write as inline string or number so we never corrupt sharedStrings indexes
      const isNum = val !== "" && /^-?\d+(\.\d+)?$/.test(String(val)) && !STATUS_SET.has(upper);
      const looksDate = /^\d{4}-\d{2}-\d{2}$/.test(String(val));

      [...el.getElementsByTagNameNS(NS, "v")].forEach((n) => n.parentNode.removeChild(n));
      [...el.getElementsByTagNameNS(NS, "is")].forEach((n) => n.parentNode.removeChild(n));

      if (looksDate) {
        el.removeAttribute("t");
        const v = sheet.doc.createElementNS(NS, "v");
        v.textContent = String(dateToSerial(val));
        el.appendChild(v);
      } else if (isNum) {
        el.removeAttribute("t");
        const v = sheet.doc.createElementNS(NS, "v");
        v.textContent = String(val);
        el.appendChild(v);
      } else {
        el.setAttribute("t", "inlineStr");
        const is = sheet.doc.createElementNS(NS, "is");
        const t = sheet.doc.createElementNS(NS, "t");
        t.setAttribute("xml:space", "preserve");
        t.textContent = String(val);
        is.appendChild(t);
        el.appendChild(is);
      }
      meta.value = val;
    }
  }
}

function ensureCell(doc, ref) {
  const { col, row } = colRow(ref);
  const sheetData = doc.getElementsByTagNameNS(NS, "sheetData")[0];
  let rowEl = null;
  for (const r of sheetData.getElementsByTagNameNS(NS, "row")) {
    if (r.getAttribute("r") === String(row)) { rowEl = r; break; }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(NS, "row");
    rowEl.setAttribute("r", String(row));
    sheetData.appendChild(rowEl);
  }
  for (const c of rowEl.getElementsByTagNameNS(NS, "c")) {
    if (c.getAttribute("r") === ref) return c;
  }
  const c = doc.createElementNS(NS, "c");
  c.setAttribute("r", ref);
  rowEl.appendChild(c);
  return c;
}

async function downloadSameFile() {
  if (!state.zip) return;
  applyEditsToXml();
  for (const sheet of state.sheets) {
    const text = xmlText(sheet.doc);
    state.zip.file(sheet.path, text);
  }
  const blob = await state.zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
    compression: "DEFLATE",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = state.fileName || "HK Daily.xlsm";
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  toast("Downloaded " + (state.fileName || "HK Daily.xlsm") + " — open it in Excel to keep using macros.");
}

/* ---------- UI builders ---------- */
function field(sheet, ref, label, type = "text") {
  const raw = currentValue(sheet, ref);
  const meta = state.cells[sheet]?.[ref];
  const isFormula = !!(meta && meta.formula);
  let display = raw;
  if (type === "date") {
    const iso = excelSerialToDate(raw);
    if (iso) display = iso;
    else if (/^\d{4}-\d{2}-\d{2}/.test(String(raw))) display = String(raw).slice(0, 10);
  }
  return `<div class="field">
    <label>${esc(label)} <small>${ref}${isFormula ? " · formula" : ""}</small></label>
    <input data-sheet="${esc(sheet)}" data-ref="${ref}" data-kind="${type}" value="${esc(display)}"/>
  </div>`;
}

function readOnly(label, value) {
  return `<div class="field"><label>${esc(label)}</label><div class="readonly">${esc(value ?? "—")}</div></div>`;
}

function pageHead(title, sub) {
  return `<div class="page-head"><h2>${esc(title)}</h2><p>${esc(sub)}</p></div>`;
}

function kpiStrip() {
  const s = mainSheetName();
  const villas = collectVillaTriples(s);
  const counts = {};
  for (const v of villas) {
    const st = String(v.status || "").toUpperCase();
    if (st) counts[st] = (counts[st] || 0) + 1;
  }
  const occPct = currentValue(s, "F10");
  const occShow = occPct && !Number.isNaN(Number(occPct)) ? (Number(occPct) * 100).toFixed(1) + "%" : esc(occPct || "—");
  const items = [
    ["Occupancy", occShow, "from sheet formula"],
    ["Arrivals", currentValue(s, "F11") || "—", "today"],
    ["Departures", currentValue(s, "F12") || "—", "today"],
    ["Occupied villas", currentValue(s, "F14") || String(counts.OCC || "—"), ""],
    ["Adults", currentValue(s, "F16") || "—", "in house"],
    ["Children", currentValue(s, "F17") || "—", "in house"],
  ];
  return `<div class="kpis">${items.map(([k,v,n]) => `<div class="kpi"><span>${k}</span><b>${esc(v)}</b><em>${esc(n)}</em></div>`).join("")}</div>`;
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function bindInputs(root) {
  root.querySelectorAll("input, select, textarea").forEach((el) => {
    if (!el.dataset.sheet || !el.dataset.ref) return;
    el.addEventListener("change", onFieldChange);
    el.addEventListener("input", onFieldChange);
  });
}

function onFieldChange(e) {
  const el = e.target;
  const val = el.value;
  setValue(el.dataset.sheet, el.dataset.ref, val);
  if (el.dataset.kind === "status") {
    const st = String(val).toUpperCase();
    const card = el.closest(".villa");
    if (card) {
      card.className = "villa st-" + st;
      const chip = card.querySelector(".chip");
      if (chip) { chip.className = "chip " + st; chip.textContent = val || "—"; }
    }
  }
}

function renderAll() {
  if (!state.sheets.length) return;
  $("app").classList.remove("hidden");
  $("empty").classList.add("hidden");
  showTab("daily");
}

function showTab(id) {
  document.querySelectorAll("[data-tab]").forEach((b) => b.classList.toggle("active", b.dataset.tab === id));
  const host = $("view");
  host.innerHTML = "";
  if (id === "daily") host.appendChild(viewDaily());
  if (id === "villas") host.appendChild(viewVillas());
  if (id === "lists") host.appendChild(viewLists());
  if (id === "staff") host.appendChild(viewStaff());
  if (id === "match") host.appendChild(viewGenericSheet(state.sheets[1]?.name, "MATCH"));
  if (id === "supervisor") host.appendChild(viewGenericSheet(state.sheets[2]?.name, "Supervisor allocation"));
  if (id === "raw") host.appendChild(viewRaw());
}

function viewDaily() {
  const s = mainSheetName();
  const el = document.createElement("div");
  el.className = "grid";
  el.innerHTML = `
    ${pageHead("Daily overview", "Duty desk, in-house counts and guest notes. Formula totals stay in Excel.")}
    ${kpiStrip()}
    <div class="cards">
      <div class="card">
        <h3>Duty desk</h3>
        <div class="form-grid">
          ${field(s, "I7", "Core value of the day")}
          ${field(s, "O4", "Date")}
          ${field(s, "AG1", "Manager on duty / chef")}
          ${field(s, "AG2", "Housekeeping day duty")}
          ${field(s, "AG3", "Housekeeping night duty")}
          ${field(s, "AG4", "Security on duty")}
          ${field(s, "AG5", "Engineering hotline")}
        </div>
      </div>
      <div class="card">
        <h3>In-house</h3>
        <p class="hint">Edit guest counts. Occupancy and villa totals remain Excel formulas and refresh when the file is opened.</p>
        <div class="form-grid">
          ${field(s, "F16", "Adults in house", "number")}
          ${field(s, "F17", "Children in house", "number")}
          ${readOnly("Occupied villas", currentValue(s, "F14"))}
          ${readOnly("Vacant villas", currentValue(s, "F15"))}
          ${readOnly("Arrival total", currentValue(s, "F11"))}
          ${readOnly("Departure total", currentValue(s, "F12"))}
          ${readOnly("Villa moves", currentValue(s, "F13"))}
        </div>
      </div>
      <div class="card">
        <h3>Guest notes</h3>
        <div class="form-grid">
          ${field(s, "G8", "VIP arrivals")}
          ${field(s, "L8", "Next-day preferences")}
          ${field(s, "L12", "Floor note")}
        </div>
      </div>
    </div>`;
  bindInputs(el);
  return el;
}

function collectVillaTriples(sheetName) {
  const map = state.cells[sheetName] || {};
  const out = [];
  // Layout: groups of V# / CAT / STAT starting at A35, repeating every 4 columns, two blocks
  const startRows = [36, 52];
  const headerRows = [35, 51];
  const nameRows = [32, 48];
  const groupCols = [1, 5, 9, 13, 17, 21, 25, 29, 33]; // A E I M Q U Y AC AG
  for (let block = 0; block < 2; block++) {
    const headerRow = headerRows[block];
    const start = startRows[block];
    const nameRow = nameRows[block];
    for (const c of groupCols) {
      const vRef = colLetter(c) + headerRow;
      if (String(currentValue(sheetName, vRef)).replace(/\s/g, "") !== "V#") continue;
      const attendant = currentValue(sheetName, colLetter(c) + nameRow);
      for (let r = start; r < start + 14; r++) {
        const vn = currentValue(sheetName, colLetter(c) + r);
        const cat = currentValue(sheetName, colLetter(c + 1) + r);
        const st = currentValue(sheetName, colLetter(c + 2) + r);
        if (!vn || String(vn).toUpperCase() === "V#") continue;
        out.push({
          attendant,
          villa: String(vn),
          cat: String(cat || ""),
          status: String(st || ""),
          refs: { v: colLetter(c) + r, cat: colLetter(c + 1) + r, st: colLetter(c + 2) + r },
        });
      }
    }
  }
  return out;
}

function viewVillas() {
  const s = mainSheetName();
  const villas = collectVillaTriples(s);
  const el = document.createElement("div");
  const groups = {};
  for (const v of villas) {
    const g = v.attendant || "Unassigned";
    (groups[g] ||= []).push(v);
  }
  const opts = STATUS_CODES.map((c) => `<option value="${c}">${c}</option>`).join("");
  el.innerHTML = `
    ${pageHead("Villa board", "Status colours in Excel follow styles already stored in this workbook.")}
    <div class="card">
      <h3>Legend</h3>
      <div class="legend">${STATUS_CODES.map((c) => `<span class="chip ${c}">${c}</span>`).join("")}</div>
      <div class="searchbar"><input id="villaFilter" type="search" placeholder="Search villa, attendant or status"/></div>
    </div>
    <div id="villaGroups" class="grid"></div>`;
  const host = el.querySelector("#villaGroups");
  for (const [name, list] of Object.entries(groups)) {
    const card = document.createElement("div");
    card.className = "card villa-group";
    card.innerHTML = `<h3>${esc(name)} · ${list.length} villas</h3><div class="villa-grid"></div>`;
    const grid = card.querySelector(".villa-grid");
    for (const v of list) {
      const st = String(v.status || "").toUpperCase();
      const box = document.createElement("div");
      box.className = "villa st-" + st;
      box.dataset.hay = (v.villa + " " + name + " " + v.status + " " + v.cat).toLowerCase();
      box.innerHTML = `
        <div class="top"><div class="vn">${esc(v.villa)}</div><span class="chip ${esc(st)}">${esc(v.status || "—")}</span></div>
        <div class="cat">${esc(v.cat || "")}</div>
        <select data-sheet="${esc(s)}" data-ref="${v.refs.st}" data-kind="status">
          <option value=""></option>
          ${opts}
        </select>`;
      const sel = box.querySelector("select");
      sel.value = String(v.status || "").toUpperCase();
      if (sel.value && ![...sel.options].some((o) => o.value === sel.value)) {
        const extra = document.createElement("option");
        extra.value = v.status;
        extra.textContent = v.status;
        sel.appendChild(extra);
        sel.value = v.status;
      }
      grid.appendChild(box);
    }
    host.appendChild(card);
  }
  bindInputs(el);
  el.querySelector("#villaFilter").addEventListener("input", (e) => {
    const q = e.target.value.toLowerCase().trim();
    el.querySelectorAll(".villa").forEach((n) => n.classList.toggle("hidden", q && !n.dataset.hay.includes(q)));
  });
  return el;
}

function rangeInputs(sheet, startCol, startRow, rows, cols, headers, title, note) {
  let html = `<div class="card"><h3>${esc(title || "Section")}</h3>`;
  if (note) html += `<p class="hint">${esc(note)}</p>`;
  html += `<div class="table-wrap"><table><thead><tr>`;
  headers.forEach((h) => { html += `<th>${esc(h || "")}</th>`; });
  html += "</tr></thead><tbody>";
  for (let r = 0; r < rows; r++) {
    html += "<tr>";
    for (let c = 0; c < cols; c++) {
      const ref = colLetter(startCol + c) + (startRow + r);
      const val = currentValue(sheet, ref);
      html += `<td><input data-sheet="${esc(sheet)}" data-ref="${ref}" value="${esc(val)}"/></td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table></div></div>";
  return html;
}

function viewLists() {
  const s = mainSheetName();
  const el = document.createElement("div");
  el.innerHTML = `
    ${pageHead("Moves and traffic", "Villa moves, arrivals, departures and celebration rooms.")}
    ${rangeInputs(s, 1, 20, 10, 4, ["From", "To", "From", "To"], "Villa moves", "Paired from / to rooms")}
    ${rangeInputs(s, 9, 19, 12, 7, ["1","2","3","4","5","6","7"], "Arrivals", "Villa numbers arriving today")}
    ${rangeInputs(s, 17, 19, 12, 7, ["1","2","3","4","5","6","7"], "Departures", "Villa numbers departing today")}
    ${rangeInputs(s, 25, 20, 10, 9, ["Honeymoon","","","Birthday","","","Anniversary","",""], "Celebrations", "Special-occasion villas")}
  `;
  bindInputs(el);
  return el;
}

function viewStaff() {
  const s = mainSheetName();
  const el = document.createElement("div");
  el.innerHTML = `
    ${pageHead("Team and leave", "Leave day counts stay as Excel formulas.")}
    ${rangeInputs(s, 1, 67, 12, 10, ["Name","","","Type","","","Start","","","End"], "Attendants on leave", "Name, leave type, start and end dates")}
    ${rangeInputs(s, 1, 87, 8, 16, ["Name","","","Role","","","","Status","","","Section","","","","Contact"], "Housekeeping supervisors", "")}
    ${rangeInputs(s, 1, 101, 8, 8, ["No","Name","","","","Role","","Section"], "Public area", "")}
  `;
  bindInputs(el);
  return el;
}

function viewGenericSheet(name, title) {
  const el = document.createElement("div");
  if (!name || !state.cells[name]) {
    el.innerHTML = `${pageHead(title || "Sheet", "")}<div class="card"><p class="hint">This sheet was not found in the uploaded workbook.</p></div>`;
    return el;
  }
  const refs = Object.keys(state.cells[name]).sort((a, b) => {
    const A = colRow(a), B = colRow(b);
    return A.row - B.row || A.colNum - B.colNum;
  });
  const rows = {};
  for (const ref of refs) {
    const { row } = colRow(ref);
    (rows[row] ||= []).push(ref);
  }
  let html = `${pageHead(title || name, "Every populated cell from the original sheet. Overwrite a formula only if you want a constant.")}
    <div class="card">
    <div class="searchbar"><input id="genFilter" type="search" placeholder="Find room, name or status"/></div>
    <div class="table-wrap"><table><tbody>`;
  for (const [row, list] of Object.entries(rows)) {
    html += `<tr data-row="${row}">`;
    for (const ref of list) {
      const meta = state.cells[name][ref];
      const val = currentValue(name, ref);
      html += `<td><small>${ref}${meta.formula ? " ƒ" : ""}</small><br>
        <input data-sheet="${esc(name)}" data-ref="${ref}" value="${esc(val)}"/></td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table></div></div>";
  el.innerHTML = html;
  bindInputs(el);
  el.querySelector("#genFilter")?.addEventListener("input", (e) => {
    const q = e.target.value.toLowerCase().trim();
    el.querySelectorAll("tr[data-row]").forEach((tr) => {
      tr.classList.toggle("hidden", q && !tr.textContent.toLowerCase().includes(q));
    });
  });
  return el;
}

function viewRaw() {
  const el = document.createElement("div");
  const names = state.sheets.map((s) => s.name);
  el.innerHTML = `${pageHead("Cell inspector", "Edit any address when a field is not on the other screens. Style index is kept.")}
    <div class="card">
    <div class="form-grid">
      <div class="field"><label>Sheet</label><select id="rawSheet">${names.map((n) => `<option>${esc(n)}</option>`).join("")}</select></div>
      <div class="field"><label>Cell</label><input id="rawRef" placeholder="C36"/></div>
      <div class="field"><label>Value</label><input id="rawVal" placeholder="New value"/></div>
    </div>
    <p class="hint" id="rawPreview"></p>
    <button id="rawApply" class="primary" type="button">Queue change</button>
  </div>`;
  const preview = () => {
    const sh = el.querySelector("#rawSheet").value;
    const ref = el.querySelector("#rawRef").value.trim().toUpperCase();
    const meta = state.cells[sh]?.[ref];
    el.querySelector("#rawPreview").textContent = meta
      ? `${sh}!${ref} = ${JSON.stringify(currentValue(sh, ref))}  style=${meta.style}  formula=${meta.formula || "—"}`
      : (ref ? "Cell empty or not in original file (will be created on save)." : "");
  };
  el.querySelector("#rawRef").addEventListener("input", preview);
  el.querySelector("#rawSheet").addEventListener("change", preview);
  el.querySelector("#rawApply").addEventListener("click", () => {
    const sh = el.querySelector("#rawSheet").value;
    const ref = el.querySelector("#rawRef").value.trim().toUpperCase();
    const val = el.querySelector("#rawVal").value;
    if (!ref) return;
    setValue(sh, ref, val);
    preview();
    toast("Queued " + sh + "!" + ref);
  });
  return el;
}

function openFile(file) {
  if (!file) return;
  loadXlsm(file).catch((err) => { console.error(err); toast(String(err), "warn"); });
}

/* ---------- wire ---------- */
window.addEventListener("DOMContentLoaded", () => {
  $("fileInput").addEventListener("change", (e) => openFile(e.target.files?.[0]));
  $("openBtn").addEventListener("click", () => $("fileInput").click());
  $("saveBtn").addEventListener("click", () => {
    $("saveCopy").textContent = "Download “" + (state.fileName || "HK Daily.xlsm") + "” with " + state.dirty.size + " edited cell" + (state.dirty.size === 1 ? "" : "s") + ". Macros and colours stay in the original workbook.";
    $("saveModal").classList.add("show");
  });
  $("saveCancel").addEventListener("click", () => $("saveModal").classList.remove("show"));
  $("saveConfirm").addEventListener("click", () => {
    $("saveModal").classList.remove("show");
    downloadSameFile().catch((err) => { console.error(err); toast(String(err), "warn"); });
  });
  $("resetBtn").addEventListener("click", () => {
    state.dirty.clear();
    const badge = $("dirtyCount");
    if (badge) { badge.textContent = ""; badge.classList.add("hidden"); }
    showTab(document.querySelector(".side-nav button.active")?.dataset.tab || "daily");
    toast("Edits discarded. The uploaded workbook is unchanged.");
  });
  document.querySelectorAll("[data-tab]").forEach((b) => {
    b.addEventListener("click", () => showTab(b.dataset.tab));
  });

  const drop = $("dropZone");
  if (drop) {
    drop.addEventListener("click", () => $("fileInput").click());
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("drag"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
    drop.addEventListener("drop", (e) => {
      e.preventDefault();
      drop.classList.remove("drag");
      openFile(e.dataTransfer.files?.[0]);
    });
  }
});
