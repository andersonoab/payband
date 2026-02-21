// app.js - v4
// Importa 1 Excel e calcula Compa + Status.
// Prioridade:
// 1) Se existir aba "Pay Bands": usa tabela (Pay Positioning 80/100/120 x Sonova Level A-J) comparando por Job Family (Position Role Family externalName).
// 2) Se não existir ou faltar faixa: estima por grupo + level + moeda (P20/P50/P80 ou 0,8/1,2).
//
// v4: Colunas extras e filtros dinâmicos a partir da aba Excel Output (ex: CDC, Chefia, Tipo, etc).
// - Exibe colunas extras selecionáveis (Colunas extras).
// - Cria filtros automáticos para colunas categóricas (2..30 valores únicos).

(() => {
  "use strict";

  const STORAGE_KEY = "sonova_comp_bands_v4";
  const STORAGE_META = "sonova_comp_bands_meta_v4";
  const STORAGE_COLS = "sonova_comp_bands_cols_v4";

  const el = (id) => document.getElementById(id);

  const ui = {
    fileInput: el("fileInput"),
    btnExportTxt: el("btnExportTxt"),
    btnExportXlsx: el("btnExportXlsx"),
    btnClear: el("btnClear"),
    btnApply: el("btnApply"),
    btnReset: el("btnReset"),

    fSearch: el("fSearch"),
    fJobFamily: el("fJobFamily"),
    fBand: el("fBand"),
    fLevel: el("fLevel"),
    fStatus: el("fStatus"),
    fCurrency: el("fCurrency"),
    fMinCompa: el("fMinCompa"),
    fMaxCompa: el("fMaxCompa"),

    extraFilters: el("extraFilters"),
    colPickerBody: el("colPickerBody"),

    dataInfo: el("dataInfo"),
    tableInfo: el("tableInfo"),

    kpiTotal: el("kpiTotal"),
    kpiBelow: el("kpiBelow"),
    kpiWithin: el("kpiWithin"),
    kpiAbove: el("kpiAbove"),
    kpiAvgCompa: el("kpiAvgCompa"),

    tblHead: el("tblHead"),
    tblBody: el("tblBody"),
    diagBox: el("diagBox"),
    buildInfo: el("buildInfo"),
  };

  const state = {
    rows: [],
    meta: null,
    filtered: [],
    extraColumns: [],
    visibleExtraColumns: [],
    extraFilterValues: {}, // {col: selectedValue}
    extraFilterModes: {}, // {col: 'select'|'contains'}
    sort: { key: null, dir: 1 }, // dir: 1 asc, -1 desc
  };

  function nowISO() { return new Date().toISOString(); }
  function safeText(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

  function parseNumber(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    const s = String(v).trim();
    if (!s) return null;
    const s1 = s.replace(/\s/g, "");
    const hasComma = s1.includes(",");
    const hasDot = s1.includes(".");
    let normalized = s1;

    if (hasComma && hasDot) {
      const lastComma = s1.lastIndexOf(",");
      const lastDot = s1.lastIndexOf(".");
      if (lastComma > lastDot) normalized = s1.replace(/\./g, "").replace(",", ".");
      else normalized = s1.replace(/,/g, "");
    } else if (hasComma && !hasDot) {
      normalized = s1.replace(/\./g, "").replace(",", ".");
    } else {
      normalized = s1.replace(/,/g, "");
    }

    const n = Number(normalized);
    return Number.isFinite(n) ? n : null;
  }

  function fmtMoney(n, currency) {
    if (!Number.isFinite(n)) return "";
    const cur = currency || "BRL";
    try {
      return new Intl.NumberFormat("pt-BR", { style: "currency", currency: cur, maximumFractionDigits: 0 }).format(n);
    } catch {
      return String(Math.round(n));
    }
  }

  function fmtNum(n, d = 2) {
    if (!Number.isFinite(n)) return "";
    return n.toFixed(d).replace(".", ",");
  }

  function getSortValue(row, key) {
    if (!key) return null;
    if (key.startsWith("extra:")) {
      const c = key.slice("extra:".length);
      const v = row.extras ? row.extras[c] : null;
      const n = parseNumber(v);
      if (n !== null) return n;
      return safeText(v).toLowerCase();
    }
    const direct = row[key];

    // números prioritários
    if (["baseSalary","p80","p100","p120","compa"].includes(key)) {
      return Number.isFinite(direct) ? direct : null;
    }

    if (typeof direct === "number") return Number.isFinite(direct) ? direct : null;
    return safeText(direct).toLowerCase();
  }

  function applySort(rows) {
    const key = state.sort && state.sort.key ? state.sort.key : null;
    const dir = state.sort && state.sort.dir ? state.sort.dir : 1;
    if (!key) return rows;

    const arr = rows.slice();
    arr.sort((a,b) => {
      const va = getSortValue(a, key);
      const vb = getSortValue(b, key);

      if (va === null && vb === null) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;

      // number vs string
      const na = typeof va === "number";
      const nb = typeof vb === "number";
      if (na && nb) return (va - vb) * dir;
      if (!na && !nb) return va.localeCompare(vb, "pt-BR") * dir;
      // mix
      return (String(va)).localeCompare(String(vb), "pt-BR") * dir;
    });
    return arr;
  }

  function setSort(key) {
    if (!key) return;
    if (!state.sort) state.sort = { key: null, dir: 1 };
    if (state.sort.key === key) state.sort.dir = state.sort.dir * -1;
    else {
      state.sort.key = key;
      state.sort.dir = 1;
    }
  }

  function setSelectOptions(selectEl, values, allLabel = "Todos") {
    const uniq = Array.from(new Set(values.map(safeText).filter(Boolean))).sort((a,b)=>a.localeCompare(b,"pt-BR"));
    selectEl.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = allLabel;
    selectEl.appendChild(optAll);
    for (const v of uniq) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      selectEl.appendChild(o);
    }
  }

  function badgeNode(status) {
    const s = safeText(status);
    const span = document.createElement("span");
    span.className = "badge";
    span.textContent = s || "";
    if (s.startsWith("Dentro")) span.classList.add("ok");
    else if (s.startsWith("Abaixo")) span.classList.add("danger");
    else if (s.startsWith("Acima")) span.classList.add("warn");
    else span.classList.add("na");
    return span;
  }

  function buildBandViz({ p80, p100, p120, salary }) {
    const w = 260;
    const h = 26;
    const pad = 6;

    const vals = [p80, p100, p120, salary].filter(Number.isFinite);
    if (!vals.length) return "";

    const minV = Math.min(...vals);
    const maxV = Math.max(...vals);
    const span = Math.max(1, maxV - minV);
    const left = minV - span * 0.10;
    const right = maxV + span * 0.10;

    const x = (v) => {
      const t = (v - left) / (right - left);
      return pad + t * (w - pad*2);
    };

    const x80 = Number.isFinite(p80) ? x(p80) : null;
    const x100 = Number.isFinite(p100) ? x(p100) : null;
    const x120 = Number.isFinite(p120) ? x(p120) : null;
    const xs = Number.isFinite(salary) ? x(salary) : null;

    const yMid = 14;

    const rect = (x1, x2) => {
      const a = Math.min(x1, x2);
      const b = Math.max(x1, x2);
      return `<rect class="range" x="${a}" y="${yMid-7}" width="${Math.max(1,b-a)}" height="14" rx="7" ry="7"></rect>`;
    };

    const line = (xx, cls) => `<line class="${cls}" x1="${xx}" y1="${yMid-9}" x2="${xx}" y2="${yMid+9}"></line>`;
    const axis = `<line class="axis" x1="${pad}" y1="${yMid}" x2="${w-pad}" y2="${yMid}"></line>`;

    let out = `<div class="bandviz"><svg viewBox="0 0 ${w} ${h}" role="img">`;
    out += axis;
    if (x80 !== null && x120 !== null) out += rect(x80, x120);
    if (x100 !== null) out += line(x100, "mid");
    if (xs !== null) {
      out += `<line class="ptline" x1="${xs}" y1="${yMid-9}" x2="${xs}" y2="${yMid+9}"></line>`;
      out += `<circle class="pt" cx="${xs}" cy="${yMid}" r="3.2"></circle>`;
    }
    out += `</svg></div>`;
    return out;
  }

  function guessColumn(headers, patterns) {
    const h = headers.map(safeText);
    for (const pat of patterns) {
      const re = pat instanceof RegExp ? pat : new RegExp(pat, "i");
      const found = h.find(col => re.test(col));
      if (found) return found;
    }
    return null;
  }

  function sheetToJson(wb, sheetName, opts = {}) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true, ...opts });
  }

  function detectMainSheet(wb) {
    const names = wb.SheetNames || [];
    const m1 = names.find(n => /excel\s*output/i.test(n));
    if (m1) return m1;

    for (const name of names) {
      const rows = sheetToJson(wb, name, { range: 0 });
      const headers = Object.keys(rows[0] || {}).map(safeText).join(" | ").toLowerCase();
      if (headers.includes("base salary") || headers.includes("salary") || headers.includes("sonova level")) return name;
    }
    return names[0] || null;
  }

  function findPayBandsSheet(wb) {
    const names = wb.SheetNames || [];
    return names.find(n => /pay\s*bands?/i.test(n)) || null;
  }

  function normalizeLevel(v) {
    const s = safeText(v).toUpperCase();
    const m = s.match(/([A-J])/);
    return m ? m[1] : "";
  }

  function normalizeCurrency(v) {
    const s = safeText(v).toUpperCase();
    return s || "BRL";
  }

  function normKey(parts) {
    return parts.map(p => safeText(p).toUpperCase()).join("|");
  }

  function normalizeGroupMatch(baseGroup, availableGroups) {
    const g = safeText(baseGroup);
    if (!g) return g;
    const hit = (availableGroups || []).find(x => safeText(x).toUpperCase() === g.toUpperCase());
    return hit || g;
  }

  function mapGroupFallback(baseGroup, jobFamily, availableGroups) {
    const direct = normalizeGroupMatch(baseGroup, availableGroups);
    if ((availableGroups || []).some(x => safeText(x).toUpperCase() === safeText(direct).toUpperCase())) return direct;

    const jf = safeText(jobFamily).toLowerCase();
    const hasHearing = jf.includes("hearing");

    const gH = (availableGroups || []).find(x => x.toLowerCase() === "hearing care") || "";
    const gAll = (availableGroups || []).find(x => x.toLowerCase().includes("all families")) || "";

    if (hasHearing && gH) return gH;
    if (!hasHearing && gAll) return gAll;

    return safeText(baseGroup);
  }

  function parseEmployees(rows) {
    if (!rows.length) return { employees: [], info: "Sem dados.", extraColumns: [] };
    const headers = Object.keys(rows[0] || {}).map(safeText);

    const colFirst = guessColumn(headers, [/^first\s*name$/i, /primeiro\s*nome/i]);
    const colLast  = guessColumn(headers, [/^last\s*name$/i, /sobrenome/i]);
    const colName  = guessColumn(headers, [/employee\s*name/i, /^name$/i, /nome/i]);
    const colId    = guessColumn(headers, [/sonova\s*id/i, /employee\s*id/i, /\bid\b/i, /matr/i]);

    // Garantir leitura do "Position Role Family (externalName)"
    const colJobFamily = guessColumn(headers, [
      /position\s*role\s*family\s*\(externalname\)/i,
      /role\s*family.*externalname/i,
      /position\s*role\s*family/i,
      /job\s*family/i,
      /family/i
    ]);

    const colBand  = guessColumn(headers, [/pay\s*band/i, /\bbanda?\b/i, /band/i]);
    const colLevel = guessColumn(headers, [/position\s*sonova\s*level/i, /sonova\s*level/i, /externalname/i, /^level$/i]);
    const colCurrency = guessColumn(headers, [/currency/i, /moeda/i]);
    const colSalary = guessColumn(headers, [/base\s*salary\s*100/i, /base\s*salary/i, /salary/i, /sal[aá]rio/i]);

    const coreCols = new Set([colFirst,colLast,colName,colId,colJobFamily,colBand,colLevel,colCurrency,colSalary].filter(Boolean));

    const extraColumns = headers.filter(h => h && !coreCols.has(h));

    const employees = rows.map((r, idx) => {
      const first = safeText(colFirst ? r[colFirst] : "");
      const last  = safeText(colLast ? r[colLast] : "");
      const fullName = safeText(colName ? r[colName] : "") || safeText([first, last].filter(Boolean).join(" "));
      const salary = parseNumber(colSalary ? r[colSalary] : null);

      const extras = {};
      for (const h of extraColumns) {
        const v = r[h];
        const n = parseNumber(v);
        extras[h] = (n !== null) ? n : safeText(v);
      }

      return {
        rowIndex: idx + 1,
        employeeId: safeText(colId ? r[colId] : ""),
        employeeName: fullName,
        jobFamily: safeText(colJobFamily ? r[colJobFamily] : ""),
        payBand: safeText(colBand ? r[colBand] : ""),
        levelRaw: safeText(colLevel ? r[colLevel] : ""),
        levelLetter: normalizeLevel(colLevel ? r[colLevel] : ""),
        currency: normalizeCurrency(colCurrency ? r[colCurrency] : "BRL"),
        baseSalary: Number.isFinite(salary) ? salary : null,
        extras
      };
    }).filter(e => Number.isFinite(e.baseSalary));

    const info = `Colaboradores: ${employees.length} (aba carregada).`;
    return { employees, info, extraColumns };
  }

  function percentile(sorted, p) {
    const n = sorted.length;
    if (!n) return null;
    if (n === 1) return sorted[0];
    const idx = (n - 1) * p;
    const lo = Math.floor(idx);
    const hi = Math.ceil(idx);
    if (lo === hi) return sorted[lo];
    const w = idx - lo;
    return sorted[lo] * (1 - w) + sorted[hi] * w;
  }

  function buildBandsFromEmployees(employees) {
    const map = new Map();

    for (const e of employees) {
      const group = safeText(e.payBand) || safeText(e.jobFamily) || "Sem grupo";
      const level = safeText(e.levelLetter) || "Sem level";
      const key = normKey([group, e.currency, level]);
      if (!map.has(key)) map.set(key, { group, currency: e.currency, level, salaries: [] });
      map.get(key).salaries.push(e.baseSalary);
    }

    const bands = [];
    for (const obj of map.values()) {
      const arr = obj.salaries.filter(Number.isFinite).sort((a,b)=>a-b);
      const p100 = percentile(arr, 0.50);
      let p80 = percentile(arr, 0.20);
      let p120 = percentile(arr, 0.80);

      if (arr.length < 5 && Number.isFinite(p100)) {
        p80 = p100 * 0.80;
        p120 = p100 * 1.20;
      }

      bands.push({
        group: obj.group,
        currency: obj.currency,
        level: obj.level,
        p80, p100, p120,
        n: arr.length,
        source: (arr.length < 5) ? "Estimado (0,8/1,2)" : "Estimado (P20/P50/P80)"
      });
    }
    return bands;
  }

  function parsePayBandsTable(wb) {
    const sheetName = findPayBandsSheet(wb);
    if (!sheetName) return { sheetName: null, lookup: null, groups: [] };

    const rows = sheetToJson(wb, sheetName, { range: 1 });
    if (!rows.length) return { sheetName, lookup: null, groups: [] };

    const headers = Object.keys(rows[0] || {}).map(safeText);

    const colGroup = guessColumn(headers, [/job\s*family/i, /family/i, /grupo/i]);
    const colCurrency = guessColumn(headers, [/currency/i, /moeda/i]);
    const colPos = guessColumn(headers, [/pay\s*positioning/i, /positioning/i, /\bposi/i]);

    const levelCols = {};
    for (const L of ["A","B","C","D","E","F","G","H","I","J"]) {
      const exact = headers.find(h => safeText(h).toUpperCase() === L);
      if (exact) levelCols[L] = exact;
    }

    if (!colGroup || !colCurrency || !colPos || Object.keys(levelCols).length < 6) {
      return { sheetName, lookup: null, groups: [] };
    }

    const lookup = new Map();
    const groups = new Set();

    for (const r of rows) {
      const group = safeText(r[colGroup]);
      const cur = normalizeCurrency(r[colCurrency]);
      const pos = parseNumber(r[colPos]);
      if (!group || !Number.isFinite(pos)) continue;

      groups.add(group);

      const levels = {};
      for (const L of Object.keys(levelCols)) {
        const v = parseNumber(r[levelCols[L]]);
        if (Number.isFinite(v)) levels[L] = v;
      }

      const k = normKey([group, cur, pos]);
      lookup.set(k, { group, currency: cur, pos, levels });
    }

    return { sheetName, lookup, groups: Array.from(groups) };
  }

  function buildBandsFromPayTableForEmployees(employees, payTable) {
    if (!payTable || !payTable.lookup) return [];

    const map = new Map();
    for (const e of employees) {
      const baseGroup = safeText(e.payBand) || safeText(e.jobFamily) || "Sem grupo";
      const group = mapGroupFallback(baseGroup, e.jobFamily, payTable.groups || []);
      const level = safeText(e.levelLetter);
      const cur = normalizeCurrency(e.currency);

      if (!level) continue;

      const k80 = normKey([group, cur, 80]);
      const k100 = normKey([group, cur, 100]);
      const k120 = normKey([group, cur, 120]);

      const r80 = payTable.lookup.get(k80);
      const r100 = payTable.lookup.get(k100);
      const r120 = payTable.lookup.get(k120);

      const p80 = r80 && r80.levels ? r80.levels[level] : null;
      const p100 = r100 && r100.levels ? r100.levels[level] : null;
      const p120 = r120 && r120.levels ? r120.levels[level] : null;

      if (!Number.isFinite(p100)) continue;

      const key = normKey([group, cur, level]);
      if (!map.has(key)) {
        map.set(key, {
          group,
          currency: cur,
          level,
          p80: Number.isFinite(p80) ? p80 : null,
          p100: p100,
          p120: Number.isFinite(p120) ? p120 : null,
          n: 0,
          source: "Tabela (Pay Bands)"
        });
      }
    }
    return Array.from(map.values());
  }

  function joinAndCompute(employees, bandsTable, bandsEstimated, payTableMeta) {
    const lookup = new Map();
    for (const b of (bandsTable || [])) {
      const key = normKey([b.group, b.currency, b.level]);
      lookup.set(key, b);
    }
    for (const b of (bandsEstimated || [])) {
      const key = normKey([b.group, b.currency, b.level]);
      if (!lookup.has(key)) lookup.set(key, b);
    }

    const availableGroups = (payTableMeta && payTableMeta.groups) ? payTableMeta.groups : [];

    return employees.map(e => {
      const baseGroup = safeText(e.payBand) || safeText(e.jobFamily) || "Sem grupo";
      const group = mapGroupFallback(baseGroup, e.jobFamily, availableGroups);
      const levelKey = safeText(e.levelLetter) || "Sem level";
      const cur = normalizeCurrency(e.currency);

      const key = normKey([group, cur, levelKey]);
      let b = lookup.get(key) || null;

      if (!b && levelKey === "Sem level") {
        const key2 = normKey([group, cur, "Sem level"]);
        b = lookup.get(key2) || null;
      }

      const p80 = b ? b.p80 : null;
      const p100 = b ? b.p100 : null;
      const p120 = b ? b.p120 : null;

      const salary = e.baseSalary;
      const compa = (Number.isFinite(salary) && Number.isFinite(p100) && p100 !== 0) ? (salary / p100) : null;

      let status = "Sem faixa";
      if (Number.isFinite(salary) && Number.isFinite(p80) && Number.isFinite(p120)) {
        if (salary < p80) status = "Abaixo (<80)";
        else if (salary > p120) status = "Acima (>120)";
        else status = "Dentro (80-120)";
      } else if (Number.isFinite(salary) && Number.isFinite(p100)) {
        status = "Sem P80/P120";
      }

      return {
        rowIndex: e.rowIndex,
        employeeId: e.employeeId,
        employeeName: e.employeeName,
        jobFamily: e.jobFamily,
        payBand: e.payBand,
        level: safeText(e.levelRaw) || "",
        levelLetter: safeText(e.levelLetter) || "",
        currency: cur,
        baseSalary: salary,

        group,
        p80, p100, p120,
        compa,
        status,
        bandSource: b ? (b.source || "Estimado") : "Nenhuma",
        extras: e.extras || {}
      };
    });
  }

  function computeKPIs(rows) {
    const total = rows.length;
    const below = rows.filter(r => r.status.startsWith("Abaixo")).length;
    const within = rows.filter(r => r.status.startsWith("Dentro")).length;
    const above = rows.filter(r => r.status.startsWith("Acima")).length;
    const compas = rows.map(r => r.compa).filter(Number.isFinite);
    const avgCompa = compas.length ? compas.reduce((a,b)=>a+b,0)/compas.length : null;

    ui.kpiTotal.textContent = String(total);
    ui.kpiBelow.textContent = String(below);
    ui.kpiWithin.textContent = String(within);
    ui.kpiAbove.textContent = String(above);
    ui.kpiAvgCompa.textContent = Number.isFinite(avgCompa) ? fmtNum(avgCompa, 2) : "0,00";
  }

  function buildTableHeader() {
    if (!ui.tblHead) return;
    ui.tblHead.innerHTML = "";

    const tr = document.createElement("tr");
    const th = (label, key, cls) => {
      const x = document.createElement("th");
      x.textContent = label;
      if (cls) x.className = cls;
      if (key) {
        x.classList.add("sortable");
        x.setAttribute("data-key", key);
        x.addEventListener("click", () => {
          setSort(key);
          applyFilters();
        });
      }
      return x;
    };

    tr.appendChild(th("Colaborador","employeeName"));
    tr.appendChild(th("Job Family","jobFamily"));
    tr.appendChild(th("Pay Band","payBand"));
    tr.appendChild(th("Level","level"));

    for (const c of (state.visibleExtraColumns || [])) {
      tr.appendChild(th(c, "extra:" + c));
    }

    tr.appendChild(th("Salário","baseSalary","num"));
    tr.appendChild(th("P80","p80","num"));
    tr.appendChild(th("P100","p100","num"));
    tr.appendChild(th("P120","p120","num"));
    tr.appendChild(th("Compa","compa","num"));
    tr.appendChild(th("Status","status"));
    tr.appendChild(th("Faixa", null));

    ui.tblHead.appendChild(tr);

    // indicador de ordenação
    const ths = ui.tblHead.querySelectorAll("th.sortable");
    ths.forEach(h => {
      h.classList.remove("sort-asc");
      h.classList.remove("sort-desc");
      const k = h.getAttribute("data-key");
      if (state.sort && state.sort.key && k === state.sort.key) {
        if (state.sort.dir === 1) h.classList.add("sort-asc");
        else h.classList.add("sort-desc");
      }
    });
  }

  function renderTable(rows) {
    buildTableHeader();
    ui.tblBody.innerHTML = "";
    const frag = document.createDocumentFragment();

    for (const r of rows) {
      const tr = document.createElement("tr");

      const td = (text, cls) => {
        const x = document.createElement("td");
        if (cls) x.className = cls;
        x.textContent = text;
        return x;
      };

      tr.appendChild(td(r.employeeName || r.employeeId || `Linha ${r.rowIndex}`));
      tr.appendChild(td(r.jobFamily || ""));
      tr.appendChild(td(r.payBand || ""));
      tr.appendChild(td(r.level || ""));

      for (const c of (state.visibleExtraColumns || [])) {
        const v = r.extras ? r.extras[c] : "";
        if (typeof v === "number" && Number.isFinite(v)) tr.appendChild(td(String(v), "num"));
        else tr.appendChild(td(safeText(v)));
      }

      tr.appendChild(td(fmtMoney(r.baseSalary, r.currency), "num"));
      tr.appendChild(td(fmtMoney(r.p80, r.currency), "num"));
      tr.appendChild(td(fmtMoney(r.p100, r.currency), "num"));
      tr.appendChild(td(fmtMoney(r.p120, r.currency), "num"));
      tr.appendChild(td(Number.isFinite(r.compa) ? fmtNum(r.compa, 2) : "", "num"));

      const tdStatus = document.createElement("td");
      tdStatus.appendChild(badgeNode(r.status));
      tr.appendChild(tdStatus);

      const tdViz = document.createElement("td");
      tdViz.innerHTML = buildBandViz({ p80: r.p80, p100: r.p100, p120: r.p120, salary: r.baseSalary });
      tr.appendChild(tdViz);

      frag.appendChild(tr);
    }

    ui.tblBody.appendChild(frag);
  }

  function renderDiag(rows, payMeta) {
    const semFaixa = rows.filter(r => r.status === "Sem faixa").length;
    const semP80120 = rows.filter(r => r.status === "Sem P80/P120").length;

    const tabela = rows.filter(r => (r.bandSource || "").startsWith("Tabela")).length;
    const estimados = rows.filter(r => (r.bandSource || "").startsWith("Estimado")).length;

    const byGroup = new Map();
    for (const r of rows) {
      const g = safeText(r.group || "Sem grupo");
      if (!byGroup.has(g)) byGroup.set(g, { g, below: 0, total: 0 });
      const o = byGroup.get(g);
      o.total++;
      if (r.status.startsWith("Abaixo")) o.below++;
    }
    const top = Array.from(byGroup.values()).sort((a,b)=>b.below-a.below).slice(0,5);

    const box = (title, lines) => {
      const div = document.createElement("div");
      div.className = "box";
      const h = document.createElement("h4");
      h.textContent = title;
      const p = document.createElement("p");
      p.innerHTML = lines.join("<br>");
      div.appendChild(h);
      div.appendChild(p);
      return div;
    };

    ui.diagBox.innerHTML = "";
    ui.diagBox.appendChild(box("Leitura do arquivo", [
      `Registros: ${rows.length}`,
      `Faixa por tabela: ${tabela}`,
      `Faixa estimada: ${estimados}`,
      `Sem faixa: ${semFaixa}`,
      `Sem P80/P120: ${semP80120}`,
      payMeta && payMeta.sheetName ? `Aba Pay Bands: ${payMeta.sheetName}` : `Aba Pay Bands: não encontrada`,
      `Colunas extras detectadas: ${(state.extraColumns || []).length}`,
      `Colunas extras visíveis: ${(state.visibleExtraColumns || []).length}`
    ]));

    ui.diagBox.appendChild(box("Top grupos com Abaixo (<80)", top.length
      ? top.map(x => `${x.g}: abaixo ${x.below} | total ${x.total}`)
      : ["Sem dados suficientes."]));

    ui.diagBox.appendChild(box("Regra aplicada", [
      `Comparação Pay Bands: Job Family (Position Role Family externalName).`,
      `Grupo: Pay Band (se existir) senão Job Family.`,
      `Prioridade: Pay Bands (80/100/120 x level) -> Estimado.`,
      `Level: extrai letra A-J do Sonova Level.`
    ]));
  }

  function buildExtraFilters(rows) {
    if (!ui.extraFilters) return;
    ui.extraFilters.innerHTML = "";

    const cols = state.extraColumns || [];
    const selected = new Set(state.visibleExtraColumns || []);

    // prioriza filtros das colunas marcadas
    const ordered = Array.from(selected).filter(c => cols.includes(c));

    // garante que filtros persistidos não travem quando coluna sumir
    for (const k of Object.keys(state.extraFilterValues || {})) {
      if (!cols.includes(k)) {
        delete state.extraFilterValues[k];
        delete state.extraFilterModes[k];
      }
    }

    const makeField = (labelText) => {
      const wrap = document.createElement("div");
      wrap.className = "field";

      const lab = document.createElement("label");
      lab.textContent = labelText;

      wrap.appendChild(lab);
      return wrap;
    };

    const valuesFor = (col) => rows
      .map(r => r.extras ? r.extras[col] : null)
      .map(v => safeText(v))
      .filter(Boolean);

    for (const col of ordered) {
      const vals = valuesFor(col);
      const uniq = Array.from(new Set(vals));
      const maxLen = uniq.reduce((m, x) => Math.max(m, x.length), 0);

      const wrap = makeField(col);

      // Se for categórico pequeno: select. Senão: input de busca (contains).
      const isSmallCategorical = (uniq.length >= 2 && uniq.length <= 30 && maxLen <= 60);

      if (isSmallCategorical) {
        const sel = document.createElement("select");
        sel.setAttribute("data-col", col);
        sel.setAttribute("data-mode", "select");

        const optAll = document.createElement("option");
        optAll.value = "";
        optAll.textContent = "Todos";
        sel.appendChild(optAll);

        uniq.sort((a,b)=>a.localeCompare(b,"pt-BR"));
        for (const v of uniq) {
          const o = document.createElement("option");
          o.value = v;
          o.textContent = v;
          sel.appendChild(o);
        }

        const saved = safeText(state.extraFilterValues[col] || "");
        sel.value = saved;

        state.extraFilterModes[col] = "select";

        sel.addEventListener("change", () => {
          state.extraFilterValues[col] = sel.value;
          state.extraFilterModes[col] = "select";
          applyFilters();
        });

        const w2 = document.createElement("div");
        w2.className = "inputwrap";
        const b2 = document.createElement("button");
        b2.className = "clearbtn";
        b2.type = "button";
        b2.textContent = "X";
        w2.appendChild(sel);
        w2.appendChild(b2);
        const sync2 = () => { b2.style.display = sel.value ? "inline-flex" : "none"; };
        b2.addEventListener("click", () => { sel.value = ""; sync2(); state.extraFilterValues[col] = ""; applyFilters(); sel.focus(); });
        sel.addEventListener("change", sync2);
        sync2();
        wrap.appendChild(w2);
      } else {
        const inp = document.createElement("input");
        inp.type = "text";
        inp.placeholder = "Digite para filtrar";
        inp.setAttribute("data-col", col);
        inp.setAttribute("data-mode", "contains");

        const saved = safeText(state.extraFilterValues[col] || "");
        inp.value = saved;

        state.extraFilterModes[col] = "contains";

        inp.addEventListener("input", () => {
          state.extraFilterValues[col] = inp.value;
          state.extraFilterModes[col] = "contains";
          applyFilters();
        });

        const w3 = document.createElement("div");
        w3.className = "inputwrap";
        const b3 = document.createElement("button");
        b3.className = "clearbtn";
        b3.type = "button";
        b3.textContent = "X";
        w3.appendChild(inp);
        w3.appendChild(b3);
        wrap.appendChild(w3);
        wireClearButton(inp, b3);

        // dica rápida (sem poluir): se uniq não for absurdo, adiciona datalist
        if (uniq.length > 0 && uniq.length <= 200) {
          const dlId = "dl_" + col.replace(/[^a-z0-9]/gi, "_").toLowerCase();
          inp.setAttribute("list", dlId);
          const dl = document.createElement("datalist");
          dl.id = dlId;
          uniq.slice(0, 200).sort((a,b)=>a.localeCompare(b,"pt-BR")).forEach(v => {
            const o = document.createElement("option");
            o.value = v;
            dl.appendChild(o);
          });
          wrap.appendChild(dl);
        }
      }

      ui.extraFilters.appendChild(wrap);
    }
  }

  function guessDefaultVisibleColumns(cols) {
    const wants = [/cdc/i, /chef/i, /manager/i, /supervisor/i, /tipo/i, /type/i, /cost\s*center/i, /\bcc\b/i];
    const pick = [];
    for (const c of cols) {
      if (wants.some(re => re.test(c))) pick.push(c);
    }
    return pick.slice(0, 6);
  }

  function saveVisibleColumns() {
    localStorage.setItem(STORAGE_COLS, JSON.stringify(state.visibleExtraColumns || []));
  }

  function loadVisibleColumns() {
    const raw = localStorage.getItem(STORAGE_COLS);
    if (!raw) return null;
    try {
      const v = JSON.parse(raw);
      return Array.isArray(v) ? v : null;
    } catch {
      return null;
    }
  }

  function buildColPicker() {
    if (!ui.colPickerBody) return;
    ui.colPickerBody.innerHTML = "";

    const cols = state.extraColumns || [];
    const visible = new Set(state.visibleExtraColumns || []);

    for (const c of cols) {
      const label = document.createElement("label");
      label.className = "chk";

      const inp = document.createElement("input");
      inp.type = "checkbox";
      inp.checked = visible.has(c);

      inp.addEventListener("change", () => {
        const set = new Set(state.visibleExtraColumns || []);
        if (inp.checked) set.add(c);
        else set.delete(c);
        state.visibleExtraColumns = Array.from(set).sort((a,b)=>a.localeCompare(b,"pt-BR"));
        saveVisibleColumns();
        buildExtraFilters(state.rows);
        applyFilters();
      });

      const span = document.createElement("span");
      span.textContent = c;

      label.appendChild(inp);
      label.appendChild(span);

      ui.colPickerBody.appendChild(label);
    }
  }

  function applyFilters() {
    const q = safeText(ui.fSearch.value).toLowerCase();
    const jobFamily = safeText(ui.fJobFamily.value);
    const band = safeText(ui.fBand.value);
    const level = safeText(ui.fLevel.value);
    const status = safeText(ui.fStatus.value);
    const currency = safeText(ui.fCurrency.value);
    const minCompa = parseNumber(ui.fMinCompa.value);
    const maxCompa = parseNumber(ui.fMaxCompa.value);

    let rows = state.rows.slice();

    if (q) {
      rows = rows.filter(r => {
        const extraHay = state.extraColumns.map(c => safeText((r.extras || {})[c])).join(" ");
        const hay = [
          r.employeeName, r.employeeId, r.jobFamily, r.payBand, r.level, r.status, extraHay
        ].map(safeText).join(" ").toLowerCase();
        return hay.includes(q);
      });
    }

    if (jobFamily) rows = rows.filter(r => r.jobFamily === jobFamily);
    if (band) rows = rows.filter(r => r.payBand === band);
    if (level) rows = rows.filter(r => safeText(r.level) === level);
    if (status) rows = rows.filter(r => r.status === status);
    if (currency) rows = rows.filter(r => r.currency === currency);

    if (Number.isFinite(minCompa)) rows = rows.filter(r => Number.isFinite(r.compa) && r.compa >= minCompa);
    if (Number.isFinite(maxCompa)) rows = rows.filter(r => Number.isFinite(r.compa) && r.compa <= maxCompa);

    // filtros extras (dinâmicos)
    for (const col of Object.keys(state.extraFilterValues || {})) {
      const v = safeText(state.extraFilterValues[col]);
      if (!v) continue;

      const mode = safeText((state.extraFilterModes || {})[col]) || "select";

      if (mode === "contains") {
        const vv = v.toLowerCase();
        rows = rows.filter(r => safeText((r.extras || {})[col]).toLowerCase().includes(vv));
      } else {
        rows = rows.filter(r => safeText((r.extras || {})[col]) === v);
      }
    }
rows = applySort(rows);

    state.filtered = rows;

    computeKPIs(rows);
    renderTable(rows);
    renderDiag(rows, state.meta ? state.meta.payBandsMeta : null);

    ui.tableInfo.textContent = `Exibindo ${rows.length} de ${state.rows.length}.`;
  }

  function resetFilters() {
    ui.fSearch.value = "";
    ui.fJobFamily.value = "";
    ui.fBand.value = "";
    ui.fLevel.value = "";
    ui.fStatus.value = "";
    ui.fCurrency.value = "";
    ui.fMinCompa.value = "";
    ui.fMaxCompa.value = "";

    state.extraFilterValues = {};
    if (ui.extraFilters) {
      const selects = ui.extraFilters.querySelectorAll("select[data-col]");
      selects.forEach(s => s.value = "");
    }

    applyFilters();
  }

  function fillFiltersFromRows(rows) {
    setSelectOptions(ui.fJobFamily, rows.map(r => r.jobFamily), "Todos");
    setSelectOptions(ui.fBand, rows.map(r => r.payBand), "Todos");
    setSelectOptions(ui.fLevel, rows.map(r => safeText(r.level)), "Todos");
    setSelectOptions(ui.fStatus, rows.map(r => r.status), "Todos");
    setSelectOptions(ui.fCurrency, rows.map(r => r.currency), "Todas");
  }

  function saveToStorage(rows, meta) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
    localStorage.setItem(STORAGE_META, JSON.stringify(meta || {}));
  }

  function loadFromStorage() {
    const raw = localStorage.getItem(STORAGE_KEY);
    const meta = localStorage.getItem(STORAGE_META);
    if (!raw) return false;
    try {
      const rows = JSON.parse(raw);
      const m = meta ? JSON.parse(meta) : null;
      if (!Array.isArray(rows)) return false;
      state.rows = rows;
      state.meta = m;
      return true;
    } catch {
      return false;
    }
  }

  function clearStorage() {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(STORAGE_META);
    localStorage.removeItem(STORAGE_COLS);
  }

  function exportTxt(rows, meta) {
    const m = meta || {};
    const lines = [];
    lines.push("SONOVA | COMPENSATION BANDS | USO INTERNO");
    lines.push(`Gerado em: ${new Date().toLocaleString("pt-BR")}`);
    if (m.sourceFile) lines.push(`Arquivo: ${m.sourceFile}`);
    if (m.sheet) lines.push(`Aba: ${m.sheet}`);
    if (m.payBandsMeta && m.payBandsMeta.sheetName) lines.push(`Pay Bands: ${m.payBandsMeta.sheetName}`);
    lines.push("");

    const extra = state.visibleExtraColumns || [];
    const header = ["EmployeeName","JobFamily","PayBand","Level"].concat(extra).concat(["Currency","BaseSalary","P80","P100","P120","Compa","Status","FonteFaixa"]);
    lines.push(header.join(" | "));
    lines.push("");

    for (const r of rows) {
      const extraVals = extra.map(c => safeText((r.extras || {})[c]));
      lines.push([
        safeText(r.employeeName || r.employeeId),
        safeText(r.jobFamily),
        safeText(r.payBand),
        safeText(r.level),
        ...extraVals,
        safeText(r.currency),
        Number.isFinite(r.baseSalary) ? String(Math.round(r.baseSalary)) : "",
        Number.isFinite(r.p80) ? String(Math.round(r.p80)) : "",
        Number.isFinite(r.p100) ? String(Math.round(r.p100)) : "",
        Number.isFinite(r.p120) ? String(Math.round(r.p120)) : "",
        Number.isFinite(r.compa) ? r.compa.toFixed(4) : "",
        safeText(r.status),
        safeText(r.bandSource),
      ].join(" | "));
    }

    const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `sonova_comp_bands_${new Date().toISOString().slice(0,10)}.txt`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }


  function exportXlsx(rows, meta) {
    if (!window.XLSX) return;

    const m = meta || {};
    const extraCols = (state.extraColumns || []).slice(); // todas as colunas extras detectadas

    const baseHeader = [
      "EmployeeName",
      "EmployeeId",
      "JobFamily",
      "PayBand",
      "Level",
      "LevelLetter",
      "Currency",
      "BaseSalary",
      "P80",
      "P100",
      "P120",
      "Compa",
      "Status",
      "FonteFaixa"
    ];

    const header = baseHeader.concat(extraCols);

    const aoa = [];
    aoa.push(header);

    for (const r of rows) {
      const row = [
        safeText(r.employeeName),
        safeText(r.employeeId),
        safeText(r.jobFamily),
        safeText(r.payBand),
        safeText(r.level),
        safeText(r.levelLetter),
        safeText(r.currency),
        Number.isFinite(r.baseSalary) ? Math.round(r.baseSalary) : "",
        Number.isFinite(r.p80) ? Math.round(r.p80) : "",
        Number.isFinite(r.p100) ? Math.round(r.p100) : "",
        Number.isFinite(r.p120) ? Math.round(r.p120) : "",
        Number.isFinite(r.compa) ? Number(r.compa.toFixed(4)) : "",
        safeText(r.status),
        safeText(r.bandSource)
      ];

      for (const c of extraCols) {
        const v = (r.extras || {})[c];
        const n = parseNumber(v);
        row.push(n !== null ? n : safeText(v));
      }

      aoa.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // formata números principais
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = 1; R <= range.e.r; R++) {
      // BaseSalary..P120
      for (let C of [7,8,9,10]) { // 0-based columns in aoa: BaseSalary=7
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (ws[addr] && typeof ws[addr].v === "number") ws[addr].z = "#,##0";
      }
      // Compa
      const compaAddr = XLSX.utils.encode_cell({ r: R, c: 11 });
      if (ws[compaAddr] && typeof ws[compaAddr].v === "number") ws[compaAddr].z = "0.00";
    }

    ws["!freeze"] = { xSplit: 0, ySplit: 1 };

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Export");

    const date = new Date().toISOString().slice(0,10);
    const fname = `sonova_comp_bands_export_${date}.xlsx`;

    XLSX.writeFile(wb, fname);
  }


  async function handleFile(file) {
    if (!window.XLSX) {
      ui.dataInfo.textContent = "Biblioteca XLSX não carregou. Verifique conexão ou bloqueio de CDN.";
      return;
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheet = detectMainSheet(wb);
    if (!sheet) throw new Error("Sem abas no arquivo.");

    const payMeta = parsePayBandsTable(wb);

    const rawRows = sheetToJson(wb, sheet);
    const parsed = parseEmployees(rawRows);

    state.extraColumns = (parsed.extraColumns || []).slice().sort((a,b)=>a.localeCompare(b,"pt-BR"));

    const savedCols = loadVisibleColumns();
    state.visibleExtraColumns = savedCols ? savedCols.filter(c => state.extraColumns.includes(c)) : guessDefaultVisibleColumns(state.extraColumns);

    buildColPicker();

    const bandsEstimated = buildBandsFromEmployees(parsed.employees);
    const bandsTable = payMeta && payMeta.lookup ? buildBandsFromPayTableForEmployees(parsed.employees, payMeta) : [];

    const merged = joinAndCompute(parsed.employees, bandsTable, bandsEstimated, payMeta);

    const usedTable = merged.some(r => (r.bandSource || "").startsWith("Tabela"));
    const meta = {
      importedAt: nowISO(),
      sourceFile: file.name,
      sheet,
      employeeInfo: parsed.info,
      bandInfo: usedTable
        ? "Faixas: Pay Bands (tabela) com fallback para estimativa."
        : "Faixas estimadas automaticamente por grupo + level + moeda.",
      payBandsMeta: {
        sheetName: payMeta ? payMeta.sheetName : null,
        hasTable: !!(payMeta && payMeta.lookup),
        groups: payMeta ? (payMeta.groups || []) : []
      },
      extraColumns: state.extraColumns,
      visibleExtraColumns: state.visibleExtraColumns,
      version: "v4.2"
    };

    state.rows = merged;
    state.meta = meta;

    saveToStorage(merged, meta);

    ui.dataInfo.textContent = `${meta.sourceFile} | ${meta.employeeInfo} | ${meta.bandInfo}`;
    fillFiltersFromRows(merged);
    buildExtraFilters(merged);
    applyFilters();
  }

  
  function wireClearButton(inputEl, btnEl) {
    if (!inputEl || !btnEl) return;

    const sync = () => {
      const has = safeText(inputEl.value).length > 0;
      btnEl.style.display = has ? "inline-flex" : "none";
    };

    btnEl.addEventListener("click", () => {
      inputEl.value = "";
      sync();
      applyFilters();
      inputEl.focus();
    });

    inputEl.addEventListener("input", sync);
    sync();
  }

  function initUI() {
    ui.buildInfo.textContent = "Padrão Sonova | v4.2";

    wireClearButton(ui.fSearch, el("btnClearSearch"));
    wireClearButton(ui.fMinCompa, el("btnClearMinCompa"));
    wireClearButton(ui.fMaxCompa, el("btnClearMaxCompa"));

    ui.btnApply.addEventListener("click", applyFilters);
    ui.btnReset.addEventListener("click", resetFilters);

    ui.fSearch.addEventListener("keydown", (e) => {
      if (e.key === "Enter") applyFilters();
    });

    ui.fileInput.addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0];
      if (!f) return;
      try {
        await handleFile(f);
      } catch (err) {
        console.error(err);
        ui.dataInfo.textContent = "Falha ao importar. Confirme .xlsx e colunas de Salary e Currency.";
      } finally {
        ui.fileInput.value = "";
      }
    });

    ui.btnClear.addEventListener("click", () => {
      clearStorage();
      state.rows = [];
      state.filtered = [];
      state.meta = null;
      state.extraColumns = [];
      state.visibleExtraColumns = [];
      state.extraFilterValues = {};

      ui.dataInfo.textContent = "Storage limpo. Importe um Excel.";
      ui.tableInfo.textContent = "";
      ui.tblHead.innerHTML = "";
      ui.tblBody.innerHTML = "";
      computeKPIs([]);
      fillFiltersFromRows([]);
      if (ui.extraFilters) ui.extraFilters.innerHTML = "";
      if (ui.colPickerBody) ui.colPickerBody.innerHTML = "";
      renderDiag([], null);
    });

    ui.btnExportTxt.addEventListener("click", () => {
      if (!state.filtered.length) return;
      exportTxt(state.filtered, state.meta);
    });

    if (ui.btnExportXlsx) {
      ui.btnExportXlsx.addEventListener("click", () => {
        if (!state.filtered.length) return;
        exportXlsx(state.filtered, state.meta);
      });
    }
}

  function boot() {
    initUI();

    const loaded = loadFromStorage();
    if (loaded) {
      const m = state.meta || {};
      ui.dataInfo.textContent = m.sourceFile
        ? `Storage carregado: ${m.sourceFile} | ${m.employeeInfo || ""} | ${m.bandInfo || ""}`
        : "Storage carregado.";

      // reconstruir colunas extras e seleção
      state.extraColumns = Array.isArray(m.extraColumns) ? m.extraColumns : [];
      const savedCols = loadVisibleColumns();
      state.visibleExtraColumns = savedCols ? savedCols.filter(c => state.extraColumns.includes(c)) : (Array.isArray(m.visibleExtraColumns) ? m.visibleExtraColumns : []);

      if (!state.visibleExtraColumns.length) state.visibleExtraColumns = guessDefaultVisibleColumns(state.extraColumns);

      buildColPicker();
      fillFiltersFromRows(state.rows);
      buildExtraFilters(state.rows);
      applyFilters();
    } else {
      fillFiltersFromRows([]);
      resetFilters();
      ui.dataInfo.textContent = "Importe um Excel (.xlsx). O sistema usa Pay Bands (se existir) e faz fallback para estimativa. Colunas extras viram filtros automaticamente.";
    }
  }

  boot();
})();
