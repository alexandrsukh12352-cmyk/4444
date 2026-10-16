"use strict";

const $ = (sel) => document.querySelector(sel);
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmtInt = (v) => (v == null || v === "" ? "—" : Math.round(v).toLocaleString("ru-RU"));
const fmtMoney = (v) => (v == null || v === "" ? "—" : Number(v).toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const pct = (a, b) => (b ? Math.round((a / b) * 100) : 0);
const fmtPct = (v, signed = false) => {
  if (v == null || v === "" || Number.isNaN(Number(v))) return "—";
  const n = Number(v) * 100;
  const body = Math.abs(n).toLocaleString("ru-RU", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
  if (signed) return n > 0.05 ? `+${body}` : n < -0.05 ? `−${body}` : "0%";
  return (n < -0.05 ? "−" : "") + body;
};
const fmtPp = (v) => {
  if (v == null || Number.isNaN(Number(v))) return "—";
  const n = Number(v) * 100;
  const body = Math.abs(n).toLocaleString("ru-RU", { maximumFractionDigits: 1 }) + " п.п.";
  return (n > 0.05 ? "+" : n < -0.05 ? "−" : "") + body;
};

const S = {
  init: null,
  scan: null,
  analysis: null,
  folder: "",
  opts: {
    brand: "", group: "", base: "pst", sources: {}, filters: {}, abc: [], cross_modes: {}, cross_split: "none",
    price_source: "both", price_pick: "min", monitoring_only: true, in_stock_only: false, subgroup_mode: "sheets", suggest_threshold: 0.3,
    report_kind: "full", green_when: "above", rec: {}, ui: "template",
  },
  priceFiles: [],
  editing: null, // настройка прайса: { id, draft, file }
  collapsed: {},
  analyzeSeq: 0,
  settings: null,
  settingsDirty: false,
  stockSlice: "both",
  stockEdit: "sale",
  stockProfiles: null,
  reportArea: "group",
  tab: "report",
  pa: { brand: "", path: [], hidden: {} },
};

// ------------------------------------------------------------------ сеть и уведомления

async function api(path, body) {
  const res = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body || {}) });
  let data = {};
  try { data = await res.json(); } catch { /* пусто */ }
  if (!res.ok) throw new Error(data.error || `Ошибка ${res.status}`);
  return data;
}

function toast(text, kind = "") {
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.textContent = text;
  $("#toasts").appendChild(el);
  setTimeout(() => el.remove(), kind === "bad" ? 7000 : 4000);
}

function seg(el, value, onChange) {
  if (!el) return;
  const paint = (current) => {
    el.querySelectorAll("button").forEach((b) => {
      b.classList.toggle("on", b.dataset.v === current);
      b.onclick = () => {
        if (b.dataset.v === current) return;
        paint(b.dataset.v);
        onChange(b.dataset.v);
      };
    });
  };
  paint(value);
}

function segHtml(choices, value, attr) {
  return `<div class="seg" ${attr}>${Object.entries(choices).map(([k, t]) =>
    `<button data-v="${esc(k)}" class="${k === value ? "on" : ""}">${esc(t)}</button>`).join("")}</div>`;
}

// ------------------------------------------------------------------ старт

let bootTimer = null;

function showBoot(title) {
  const boot = $("#boot");
  if (!boot) return;
  const head = $("#boot-title");
  if (head) head.textContent = title || (S.init ? "Читаю папки" : "Открываю программу");
  const steps = $("#boot-steps");
  if (steps) steps.innerHTML = `<li class="active"><i></i><div><b>Жду ответ программы</b></div><span></span></li>`;
  S.bootHold = false;
  S.bootSince = Date.now() / 1000;
  boot.classList.remove("hidden");
  const err = $("#boot-error");
  const retry = $("#boot-retry");
  if (err) { err.classList.add("hidden"); err.textContent = ""; }
  if (retry) retry.hidden = true;
  if (!boot.dataset.bound) {
    boot.dataset.bound = "1";
    if (retry) retry.onclick = () => { if (S.init) rescan(true); else start(); };
  }
  clearInterval(bootTimer);
  bootTimer = setInterval(pollBoot, 400);
  pollBoot();
}

function hideBoot() {
  clearInterval(bootTimer);
  const boot = $("#boot");
  if (boot) boot.classList.add("hidden");
}

function showBootError(message) {
  S.bootHold = true;
  clearInterval(bootTimer);
  const boot = $("#boot");
  if (boot) boot.classList.remove("hidden");
  const err = $("#boot-error");
  const retry = $("#boot-retry");
  if (err) { err.classList.remove("hidden"); err.textContent = message || "Неизвестная ошибка"; }
  if (retry) retry.hidden = false;
}

function bootRow(ph) {
  const time = ph.state === "skip" ? "не понадобилось" : ph.elapsed != null ? `${String(ph.elapsed).replace(".", ",")} с` : "";
  const pr = ph.state === "active" && ph.progress ? ph.progress : null;
  const bar = pr ? `<div class="boot-bar"><i style="width:${pct(pr[0], pr[1])}%"></i></div>` : "";
  const detail = (ph.state === "active" || ph.state === "error") && ph.detail ? `<div class="muted">${esc(ph.detail)}</div>` : "";
  return `<li class="${ph.state}"><i></i><div><b>${esc(ph.title)}</b>${detail}${bar}</div><span>${esc(time)}</span></li>`;
}

async function pollBoot() {
  if (S.bootHold) return;
  try {
    const res = await fetch("/api/progress");
    const p = await res.json();
    if (p.done && !p.error && p.started && p.started + 0.3 < (S.bootSince || 0)) return;
    const time = $("#boot-time");
    if (time) time.textContent = p.started ? `${String(p.elapsed).replace(".", ",")} с` : "";
    const steps = $("#boot-steps");
    if (steps && (p.phases || []).length) steps.innerHTML = p.phases.map(bootRow).join("");
    const log = $("#boot-log");
    if (log) {
      const rows = p.log || [];
      log.innerHTML = rows.map((row, i) => `<li class="${i === rows.length - 1 ? "on" : ""}"><b>${esc(row.stage)}</b>${row.detail ? ` · ${esc(row.detail)}` : ""}</li>`).join("");
    }
    if (p.error) showBootError(p.error);
  } catch (e) {
    const steps = $("#boot-steps");
    if (steps) steps.innerHTML = `<li class="error"><i></i><div><b>Нет ответа от программы</b><div class="muted">${esc(e.message || "")} Проверьте, что чёрное окно запуска открыто.</div></div><span></span></li>`;
  }
}

async function start() {
  showBoot();
  S.init = await api("/api/init");
  const o = S.init.options || {};
  Object.assign(S.opts, {
    base: o.base || "pst", sources: o.sources || {}, filters: o.filters || {}, abc: o.abc || [],
    cross_modes: o.cross_modes || {}, cross_split: S.init.cross_splits[o.cross_split] ? o.cross_split : "none",
    price_source: S.init.price_sources[o.price_source] ? o.price_source : "both",
    monitoring_only: o.monitoring_only !== false, in_stock_only: !!o.in_stock_only,
    subgroup_mode: o.subgroup_mode || "sheets", suggest_threshold: o.suggest_threshold ?? 0.3,
    brand: o.brand || "", group: o.group || "",
    report_kind: "full",
    sheet_filters: Array.isArray(o.sheet_filters) ? o.sheet_filters.slice() : (S.init.default_sheet_filters || ["overstock", "sale", "competitor_price"]).slice(),
    green_when: o.green_when === "below" ? "below" : "above",
    money_base: o.money_base === "potential" ? "potential" : "sales",
    rec: { ...(S.init.rec || {}), ...(o.rec || {}), up: { ...(S.init.default_up || {}), ...((o.rec || {}).up || {}) } },
    ui: "custom",
  });
  S.recCommon = clone(S.opts.rec);
  S.recMine = o.rec_mine && typeof o.rec_mine === "object" && Object.keys(o.rec_mine).length
    ? clone(o.rec_mine) : thresholdSnapshot(S.recCommon);
  S.recGroups = o.rec_groups && typeof o.rec_groups === "object" ? clone(o.rec_groups) : {};
  S.recBundled = clone(S.init.bundled_rec || {});
  S.recBundledGroups = clone(S.init.bundled_groups || {});
  // Пороги — всегда свои у каждой группы: «одни на всех» убраны из окна.
  S.recScope = "group";
  S.recSource = o.rec_source === "yours" ? "yours" : (o.rec_source === "mine" ? "mine" : (Object.keys(S.recGroups).length ? "mine" : "yours"));
  S.recGroupSource = o.rec_group_source && typeof o.rec_group_source === "object" ? clone(o.rec_group_source) : {};
  S.recDirty = false;
  S.custom = seedCustom(o);
  // «По шаблону» больше нет: всегда сохранённые настройки, стандарт — кнопкой «Сбросить к стандарту».
  if (o.custom && Object.keys(o.custom).length) applyCustom();
  else applyTemplate();
  S.opts.filters = foldClearance(S.opts.filters);
  ensureStockProfiles();
  S.folder = S.init.folder;
  $("#folder").value = S.folder;
  $("#pick-folder").classList.toggle("hidden", !S.init.can_pick_folder);
  bind();
  renderPriceControls();
  renderSplit();
  renderRecForm();
  paintUi();
  let recOpen = false;
  try { recOpen = localStorage.getItem("recOpen") === "1"; } catch (e) { /* без памяти */ }
  setRecOpen(recOpen);
  paintRepTabs();
  paintReportArea();
  S.ready = true;
  const tag = $("#build-tag");
  if (tag) tag.textContent = S.init.build_label || "";
  renderSheetFilters();
  renderStockForm();
  const opened = await rescan();
  if (opened) hideBoot();
}

const CUSTOM_KEYS = ["base", "sources", "filters", "abc", "cross_modes", "cross_split", "price_source", "price_pick",
  "monitoring_only", "in_stock_only", "subgroup_mode", "suggest_threshold", "green_when"];
const REC_NUMS = ["tolerance", "down", "min_vr", "top_factor", "mid_factor", "rounding", "max_lift"];

function clone(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function foldClearance(filters) {
  const raw = { ...(filters || {}) };
  const disc = raw.discontinued;
  delete raw.discontinued;
  const rank = { include: 0, only: 1, exclude: 2 };
  const sale = raw.sale;
  let chosen = rank[sale] != null ? sale : "include";
  if (rank[disc] > rank[chosen]) chosen = disc;
  if (chosen && chosen !== "include") raw.sale = chosen;
  else delete raw.sale;
  return raw;
}

function seedCustom(o) {
  if (o.custom && typeof o.custom === "object" && Object.keys(o.custom).length) {
    const snap = clone(o.custom);
    if (snap.filters) snap.filters = foldClearance(snap.filters);
    return snap;
  }
  const snap = {};
  for (const key of CUSTOM_KEYS) if (o[key] !== undefined) snap[key] = clone(o[key]);
  return snap;
}

function applyCustom() {
  const saved = S.custom || {};
  for (const key of CUSTOM_KEYS) {
    if (saved[key] !== undefined) S.opts[key] = clone(saved[key]);
  }
}

function customSnapshot() {
  if (S.opts.ui === "custom") {
    const snap = {};
    for (const key of CUSTOM_KEYS) snap[key] = clone(S.opts[key]);
    return snap;
  }
  const snap = clone(S.custom || {});
  if (snap.rec && S.opts.rec) {
    snap.rec.enabled = !!S.opts.rec.enabled;
    snap.rec.target = S.opts.rec.target || snap.rec.target || "main";
  }
  return snap;
}

let persistTimer = null;
function persist(saveMineNumbers = false) {
  if (!S.ready) return;
  S.custom = customSnapshot();
  clearTimeout(persistTimer);
  persistTimer = setTimeout(() => {
    const options = {
      ui: S.opts.ui, brand: S.opts.brand, group: S.opts.group, report_kind: S.opts.report_kind,
      sheet_filters: S.opts.sheet_filters, custom: S.custom,
      rec_scope: S.recScope, rec_source: S.recSource, rec_group_source: S.recGroupSource,
      rec: { enabled: !!(S.opts.rec && S.opts.rec.enabled), target: (S.opts.rec && S.opts.rec.target) || "main" },
      money_base: S.opts.money_base === "potential" ? "potential" : "sales",
    };
    if (saveMineNumbers) {
      options.save_rec = true;
      options.rec_mine = S.recMine;
      options.rec_groups = S.recGroups;
    }
    api("/api/options", { options }).catch(() => {});
  }, saveMineNumbers ? 0 : 350);
}

function bind() {
  document.querySelectorAll(".tab").forEach((t) => (t.onclick = () => showTab(t.dataset.tab)));
  const diagRefresh = $("#diag-refresh");
  if (diagRefresh) diagRefresh.onclick = () => loadDiagnose();
  const diagWrite = $("#diag-write");
  if (diagWrite) diagWrite.onclick = () => writeDiagnose();
  const diagOpen = $("#diag-open");
  if (diagOpen) diagOpen.onclick = () => {
    const path = diagOpen.dataset.path;
    if (!path) return;
    api("/api/open", { path }).catch((e) => toast(e.message, "bad"));
  };
  $("#rescan").onclick = () => rescan(true);
  const pxRescan = $("#px-rescan");
  if (pxRescan) pxRescan.onclick = () => rescan(true);
  const pxLayout = $("#px-layout");
  if (pxLayout) pxLayout.onclick = () => writePriceLayout();
  const pxCrosses = $("#px-crosses");
  if (pxCrosses) pxCrosses.onclick = () => writeCrossPrices();
  const paFrom = $("#pa-from");
  if (paFrom) paFrom.onchange = () => loadAnalytics({ date_from: paFrom.value });
  const paTo = $("#pa-to");
  if (paTo) paTo.onchange = () => loadAnalytics({ date_to: paTo.value, date_from: "" });
  const paBrand = $("#pa-brand");
  if (paBrand) paBrand.onchange = () => { S.pa.brand = paBrand.value; S.pa.path = []; paintAnalytics(); };
  const paExport = $("#pa-export");
  if (paExport) paExport.onclick = () => writePriceExport();
  const openResult = () => {
    if (!S.pxFolder) { toast("Папка результата откроется после первой выгрузки"); return; }
    api("/api/open", { path: S.pxFolder }).catch((e) => toast(e.message, "bad"));
  };
  const pxOpen = $("#px-open");
  if (pxOpen) pxOpen.onclick = openResult;
  const paOpen = $("#pa-open");
  if (paOpen) paOpen.onclick = openResult;
  seg($("#report-area"), S.reportArea, (v) => { S.reportArea = v; paintReportArea(); });
  const pxSearch = $("#px-search");
  if (pxSearch) pxSearch.oninput = () => { S.pxQuery = pxSearch.value; renderPriceAudit(); };
  const pxFilter = $("#px-filter");
  if (pxFilter) seg(pxFilter, S.pxFilter || "all", (v) => { S.pxFilter = v; renderPriceAudit(); });
  $("#folder").addEventListener("keydown", (e) => { if (e.key === "Enter") rescan(); });
  $("#open-folder").onclick = () => api("/api/open", { path: $("#folder").value }).catch((e) => toast(e.message, "bad"));
  $("#pick-folder").onclick = async () => {
    try {
      const r = await api("/api/pick-folder", { folder: $("#folder").value });
      if (r.folder) { $("#folder").value = r.folder; rescan(); }
    } catch (e) { toast(e.message, "bad"); }
  };
  $("#preset-all").onclick = () => { S.opts.filters = {}; S.opts.abc = []; renderFilters(); analyze(); };
  $("#preset-standard").onclick = () => {
    applyTemplate();
    renderFilters(); renderPriceControls(); renderSplit();
    analyze();
    toast("Настройки отчёта — стандартные: все товары, цены из АЦПКД и прайсов, кроссы по подсказке", "good");
  };
  $("#preset-clean").onclick = () => { S.opts.filters = { ...S.init.clean_preset }; renderFilters(); analyze(); };
  $("#apply-suggest").onclick = () => { S.opts.cross_modes = {}; analyze(); };
  $("#monitoring-only").onchange = (e) => { S.opts.monitoring_only = e.target.checked; analyze(); };
  $("#in-stock-only").onchange = (e) => { S.opts.in_stock_only = e.target.checked; analyze(); };
  $("#group-search").oninput = () => renderGroup();
  seg($("#money-base"), S.opts.money_base === "potential" ? "potential" : "sales", (v) => {
    S.opts.money_base = v === "potential" ? "potential" : "sales";
    persist();
    analyze();
  });
  $("#rec-toggle").onclick = () => setRecOpen(!S.recOpen);
  $("#rec-reset").onclick = resetToBundled;
  $("#rec-reset-all").onclick = resetAllBundled;
  $("#build").onclick = build;
  $("#build-all").onclick = buildAll;
  $("#stock-build-all").onclick = () => buildStock();
  seg($("#stock-slice"), S.stockSlice || "both", (v) => {
    S.stockSlice = v;
    if (v === "sale" || v === "overstock") S.stockEdit = v;
    renderStockForm();
  });
  seg($("#stock-edit"), S.stockEdit || "sale", (v) => {
    S.stockEdit = v;
    renderStockForm();
  });
  $("#check-folder").onclick = checkFolder;
  bindSettings();
}

function applyTemplate() {
  S.opts.base = "pst";
  S.opts.sources = {};
  S.opts.filters = {};
  S.opts.abc = [];
  S.opts.price_source = "both";
  S.opts.price_pick = "min";
  S.opts.monitoring_only = true;
  S.opts.in_stock_only = false;
  S.opts.cross_split = "none";
  S.opts.cross_modes = {};
  S.opts.green_when = "above";
}

function paintUi() {
  S.opts.ui = "custom";
  document.body.classList.add("detail");
}

function paintRepTabs() {
  const tab = S.repTab || "price";
  const bar = $("#rep-tabs");
  if (bar) bar.querySelectorAll("button").forEach((b) => {
    b.classList.toggle("on", b.dataset.v === tab);
    b.onclick = () => { S.repTab = b.dataset.v; paintRepTabs(); };
  });
  for (const name of ["price", "coverage", "table", "settings"]) {
    const pane = $(`#pane-${name}`);
    if (pane) pane.classList.toggle("hidden", name !== tab);
  }
}

function paintReportArea() {
  const stock = S.reportArea === "stock";
  seg($("#report-area"), stock ? "stock" : "group", (v) => { S.reportArea = v; paintReportArea(); });
  document.body.classList.toggle("area-stock", stock);
  const area = $("#stock-area");
  if (area) area.classList.toggle("hidden", !stock);
  const flow = $("#flow");
  if (flow) flow.classList.toggle("area-off", stock);
  document.querySelectorAll(".group-only").forEach((el) => el.classList.toggle("hidden", stock));
  $("#actionbar").classList.toggle("hidden", stock || S.tab !== "report");
  if (stock) renderStockForm();
}

function showTab(name) {
  S.tab = name;
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
  for (const page of ["report", "prices", "data", "settings"]) {
    const el = $(`#page-${page}`);
    if (el) el.classList.toggle("hidden", name !== page);
  }
  $("#actionbar").classList.toggle("hidden", name !== "report" || S.reportArea === "stock");
  if (name === "prices") loadAnalytics();
  if (name === "data") renderPriceAudit();
  if (name === "settings") { if (!S.settings) loadSettings(); loadDiagnose(); }
}

function renderDiagnose(d) {
  $("#diag-build").textContent = d.build_label || "";
  $("#diag-marks").innerHTML = (d.marks || []).map((m) => `
    <div class="note ${m.ok ? "" : "bad"}">${m.ok ? "Да" : "Нет"} · <b>${esc(m.name)}</b>. ${esc(m.detail)}</div>`).join("");
  $("#diag-output").textContent = d.output || "";
  const open = $("#diag-open");
  if (open) open.dataset.path = d.output || "";
  $("#diag-brief").innerHTML = d.brief_exists
    ? `<div class="note">Разбор.txt есть, изменён ${esc(d.brief_mtime)}, ${fmtInt(d.brief_size)} байт. ${esc(d.brief_path)}</div>`
    : `<div class="note">Разбор.txt в этой папке ещё нет. Он появится после «Сформировать отчёт». Путь: ${esc(d.brief_path)}</div>`;
  const col = d.column || {};
  const extra = d.brand ? ` ОЕМ прочитано: ${fmtInt(d.oem_rows)}. Цен этой марки в ПСТ: ${fmtInt(d.sister_prices)}.` : "";
  $("#diag-column").innerHTML = `<div class="note">${esc(col.text || "")}${esc(extra)}</div>`;
  const files = d.files || [];
  $("#diag-files").innerHTML = files.length
    ? files.map((f) => `<tr><td>${esc(f.name)}</td><td>${esc(f.mtime)}</td><td class="num">${fmtInt(f.size)}</td></tr>`).join("")
    : `<tr><td colspan="3" class="muted">В папке пока нет файлов</td></tr>`;
}

async function loadDiagnose() {
  try {
    const d = await api("/api/diagnose", { folder: S.folder, brand: S.opts.brand, group: S.opts.group });
    renderDiagnose(d);
  } catch (e) {
    toast(e.message, "bad");
  }
}

async function writeDiagnose() {
  try {
    const d = await api("/api/diagnose/write", { folder: S.folder, brand: S.opts.brand, group: S.opts.group });
    renderDiagnose(d);
    toast(`Записан ${d.probe_name}`, "good");
  } catch (e) {
    toast(e.message, "bad");
  }
}

// ------------------------------------------------------------------ шаг 1: папка

async function rescan(manual = false) {
  showBoot(S.init && S.scan ? "Читаю папки" : "Открываю программу");
  const folder = $("#folder").value.trim();
  $("#docs").innerHTML = `<div class="docs">${Array.from({ length: 9 }, () => `<div class="doc"><div class="skeleton" style="width:60%"></div><div class="skeleton" style="width:90%;margin-top:6px"></div></div>`).join("")}</div>`;
  $("#rescan").disabled = true;
  const pxRescan = $("#px-rescan");
  if (pxRescan) pxRescan.disabled = true;
  try {
    S.scan = await api("/api/scan", { folder });
    S.folder = S.scan.folder;
    $("#folder").value = S.folder;
    S.priceFiles = S.scan.price_files || [];
    S.editing = null;
    renderDirs();
    renderDocs();
    renderPriceFiles();
    chooseGroup();
    paintBuildAll();
    renderRatesNote();
    if (manual) toast(S.scan.reread ? `Заново прочитано файлов: ${S.scan.reread}` : "Даты файлов те же, разбор уже был сохранён");
    S.analytics = null;
    if (S.tab === "prices") loadAnalytics();
    if (S.tab === "data") renderPriceAudit();
    hideBoot();
    return true;
  } catch (e) {
    $("#docs").innerHTML = `<div class="notes"><div class="note bad">${esc(e.message)}</div></div>`;
    $("#flow").classList.add("hidden");
    updateSummary();
    showBootError(e.message);
    return false;
  } finally {
    $("#rescan").disabled = false;
    const again = $("#px-rescan");
    if (again) again.disabled = false;
  }
}

const OPTIONAL = new Set(["discounts", "gz", "positioning", "crosses", "sales", "sbs", "transit"]);

function renderDirs() {
  const dirs = (S.scan && S.scan.dirs) || [];
  const el = $("#dirs");
  if (!dirs.length) { el.innerHTML = ""; return; }
  el.innerHTML = dirs.map((d) => `<div class="dir">
      <div class="name">${esc(d.title)} <span class="pill ${d.files ? "good" : ""}">${fmtInt(d.files)}</span></div>
      <div class="hint">${esc(d.hint)}</div>
      <button class="btn small" data-open="${esc(d.path)}">Открыть</button>
    </div>`).join("");
  el.querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => api("/api/open", { path: b.dataset.open }).catch((e) => toast(e.message, "bad"))));
}

function renderDocs() {
  const sc = S.scan;
  const docs = sc.documents.map((d) => `
    <div class="doc ${d.label ? "found" : ""} ${OPTIONAL.has(d.kind) ? "optional" : ""}">
      <div class="name"><span class="dot"></span>${esc(d.title)}${d.rows != null ? `<span class="muted" style="font-weight:500">· ${fmtInt(d.rows)}</span>` : ""}</div>
      <div class="where" title="${esc(d.label)}">${d.label ? esc(d.label) : "не найден"}</div>
      <div class="hint">${esc(d.hint)}</div>
    </div>`).join("");
  const pf = sc.price_files || [];
  const priceTile = `
    <div class="doc optional ${pf.length ? "found" : ""}">
      <div class="name"><span class="dot"></span>Прайсы поставщиков${pf.length ? `<span class="muted" style="font-weight:500">· ${pf.length}</span>` : ""}</div>
      <div class="where" title="${esc(pf.map((f) => f.file).join(", "))}">${pf.length ? esc(pf.map(pfLabel).join(", ")) : "не найдены"}</div>
      <div class="hint">календарь и распознавание — ниже</div>
    </div>`;
  const notes = [...sc.warnings.map((w) => `<div class="note">${esc(w)}</div>`), ...sc.skipped.map((w) => `<div class="note info">${esc(w)}</div>`)];
  const hasBase = sc.groups.pst.length || (sc.queue || []).length;
  let empty = "";
  if (!sc.files) {
    empty = `<div class="empty"><h3>Папка пуста</h3><p>Положите выгрузки из 1С во «Входящие» или в папку группы: ПСТ, СБС, Продажи, АЦПКД. Кроссы и прайсы — в свои папки.</p><p class="mono">${esc(sc.folder)}</p></div>`;
  } else if (!hasBase) {
    empty = `<div class="empty"><h3>ПСТ не найден</h3><p>Отчёт строится от ПСТ. Проверьте во вкладке «Данные», что выгрузка лежит во «Входящие» или в папке группы и сохранена как .xlsx или .csv с шапкой.</p></div>`;
  }
  $("#docs").innerHTML = `<div class="docs">${docs}${priceTile}</div>`;
  const scanNote = $("#scan-note");
  // Замечания по файлам не закрывают экран: свёрнуты в одну строку, раскрываются по клику.
  const fold = notes.length ? `<details class="note-fold"><summary>Замечания по файлам: ${notes.length}. ${esc((sc.warnings[0] || sc.skipped[0] || "").slice(0, 90))}${notes.length > 1 ? "…" : ""}</summary><div class="notes">${notes.join("")}</div></details>` : "";
  if (scanNote) scanNote.innerHTML = `${fold}${empty}`;
  if (!sc.files || !hasBase) $("#card-folder").open = true;
  if (!sc.files && S.tab !== "data") showTab("data");
  $("#flow").classList.toggle("hidden", !hasBase);
}

// ------------------------------------------------------------------ шаг 2: группа и основа

function listedGroups() {
  if (!S.scan) return [];
  if ((S.scan.queue || []).length) return S.scan.queue;
  return (S.scan.groups && S.scan.groups.pst) || [];
}

function currentGroup() {
  return listedGroups().find((g) => g.brand === S.opts.brand && g.group === S.opts.group) || null;
}

function currentKey() {
  const g = currentGroup();
  return (g && g.key) || "";
}

function thresholdSnapshot(rec) {
  const snap = {};
  for (const field of REC_NUMS) if (rec && rec[field] != null && rec[field] !== "") snap[field] = rec[field];
  snap.up = clone((rec && rec.up) || {});
  return snap;
}

function mergeThresholds(base, overlay) {
  return {
    ...clone(base || {}),
    ...clone(overlay || {}),
    up: { ...((base && base.up) || {}), ...((overlay && overlay.up) || {}) },
  };
}

function bundledNow() {
  const key = S.recScope === "group" ? currentKey() : "";
  return mergeThresholds(S.recBundled, (S.recBundledGroups || {})[key]);
}

function currentSource() {
  if (S.recScope === "all") return S.recSource === "mine" ? "mine" : "yours";
  const key = currentKey();
  const src = (S.recGroupSource || {})[key];
  if (src === "mine" || src === "yours") return src;
  return key && S.recGroups[key] ? "mine" : "yours";
}

function setSource(value) {
  if (S.recScope === "all") S.recSource = value === "mine" ? "mine" : "yours";
  else if (currentKey()) S.recGroupSource[currentKey()] = value === "mine" ? "mine" : "yours";
}

function applyGroupRec() {
  const enabled = !!(S.opts.rec && S.opts.rec.enabled);
  const target = (S.opts.rec && S.opts.rec.target) || (S.recCommon && S.recCommon.target) || "main";
  const bundled = bundledNow();
  const base = currentSource() === "yours"
    ? bundled
    : (S.recScope === "all" ? mergeThresholds(S.recBundled, S.recMine) : mergeThresholds(bundled, S.recGroups[currentKey()]));
  S.opts.rec = { ...base, enabled, target };
}

function resetToBundled() {
  if (S.recScope === "all") {
    S.recMine = thresholdSnapshot(S.recBundled);
    S.recSource = "mine";
  } else if (currentKey()) {
    delete S.recGroups[currentKey()];
    S.recGroupSource[currentKey()] = "yours";
  }
  S.recDirty = false;
  applyGroupRec();
  renderRecForm();
  persist(true);
  analyze();
  toast("Правила группы — стандартные", "good");
}

function resetAllBundled() {
  S.recGroups = {};
  S.recGroupSource = {};
  S.recMine = thresholdSnapshot(S.recBundled);
  S.recSource = "yours";
  S.recDirty = false;
  applyGroupRec();
  renderRecForm();
  renderGroup();
  persist(true);
  analyze();
  toast("У всех групп — стандартные правила", "good");
}

function sameNum(a, b) {
  const x = Number(a);
  const y = Number(b);
  if (!Number.isFinite(x) && !Number.isFinite(y)) return true;
  return x === y;
}

let saveMineTimer = null;
function rememberGroupRec() {
  if (!S.opts.rec) return;
  const snap = thresholdSnapshot(S.opts.rec);
  if (S.recScope === "all") {
    S.recMine = snap;
    S.recSource = "mine";
  } else if (currentKey()) {
    S.recGroups[currentKey()] = snap;
    S.recGroupSource[currentKey()] = "mine";
  }
  paintRecMarks();
  renderRecScope();
  // Свои правила запоминаются сами — без кнопки «Сохранить».
  clearTimeout(saveMineTimer);
  saveMineTimer = setTimeout(() => persist(true), 700);
}

function renderRecScope() {
  const el = $("#rec-scope");
  const own = currentSource() === "mine";
  const r = S.opts.rec || {};
  if (el) el.textContent = r.enabled ? (own ? "свои правила этой группы" : "стандартные правила группы") : "выключены — в отчёт не попадут";
  const reset = $("#rec-reset");
  if (reset) reset.classList.toggle("hidden", !own);
}

function setRecOpen(open) {
  S.recOpen = !!open;
  try { localStorage.setItem("recOpen", S.recOpen ? "1" : ""); } catch (e) { /* без памяти */ }
  const panel = $("#card-recs");
  if (panel) panel.classList.toggle("hidden", !S.recOpen);
  const lab = $("#price-lab");
  if (lab) lab.classList.toggle("open", S.recOpen);
  const btn = $("#rec-toggle");
  if (btn) {
    btn.textContent = S.recOpen ? "Скрыть правила" : "Правила";
    btn.classList.toggle("primary", !S.recOpen);
  }
}

function paintRecMarks() {
  document.querySelectorAll("#groups .group-card").forEach((el) => {
    const g = (S.listed || [])[+el.dataset.i];
    const own = !!(g && g.key && ((S.recGroupSource || {})[g.key] === "mine" || (!(S.recGroupSource || {})[g.key] && S.recGroups[g.key])));
    const mark = el.querySelector("[data-ownrec]");
    if (mark) mark.classList.toggle("hidden", !own);
  });
}

function analysisFresh() {
  return !!(S.analysis && S.analysisFor === `${S.opts.brand}\t${S.opts.group}`);
}

function renderContext() {
  const g = currentGroup();
  const title = $("#context-title");
  if (!title) return;
  const a = analysisFresh() ? S.analysis : null;
  const d = a && a.dashboard;
  title.textContent = S.opts.brand || S.opts.group ? `${S.opts.brand} · ${S.opts.group || "без группы"}` : "Группа не выбрана";
  const bits = [];
  if (a && a.kept != null) bits.push(`${fmtInt(a.kept)} товаров в отчёте`);
  else if (g) bits.push(`${fmtInt(g.count)} в ПСТ`);
  if (g && g.own) bits.push("свой АЦПКД");
  else if (g && g.acp) bits.push("АЦПКД из входящих");
  const sister = a && a.sister;
  if (sister && sister.brand && !sister.oem_rows) bits.push(`${sister.brand} после цен конкурентов: файл ОЕМ не прочитан`);
  else if (sister && sister.brand) bits.push(`${sister.brand} после цен конкурентов: цена у ${fmtInt(sister.priced)} из ${fmtInt(sister.rows)}`);
  $("#context-meta").textContent = bits.join(" · ");
  const flags = g ? [["ПСТ", g.pst !== false], ["СБС", !!g.sbs], ["Продажи", !!g.sales], ["АЦПКД", !!g.acp]] : [];
  const docs = flags.map(([name, ok]) => `<span class="pill ${ok ? "good" : ""}">${esc(name)}</span>`).join("");
  const status = !g ? "" : g.ready ? `<span class="pill good">полный набор</span>` : `<span class="pill">${esc(g.reason || "не всё нашлось")}</span>`;
  $("#context-docs").innerHTML = docs + status;
  const vr = d && d.vr != null ? fmtPct(d.vr) : "—";
  const after = d && d.rec_enabled && d.vr_rec != null;
  $("#context-vr").innerHTML = after
    ? `<div class="l">ВР сейчас → после</div><div class="v">${vr} <span>→ ${fmtPct(d.vr_rec)}</span></div>`
    : `<div class="l">ВР</div><div class="v">${vr}</div>`;
}

function chooseGroup() {
  S.opts.base = "pst";
  const groups = listedGroups();
  if (!groups.some((g) => g.brand === S.opts.brand && g.group === S.opts.group) && groups.length) {
    const first = groups.find((g) => g.ready) || groups[0];
    S.opts.brand = first.brand;
    S.opts.group = first.group;
  }
  S.analysisFor = "";
  applyGroupRec();
  renderGroup();
  renderFilters();
  renderRecForm();
  updateSummary();
  analyze();
}

function renderGroup() {
  const all = listedGroups();
  const q = ($("#group-search").value || "").trim().toLowerCase();
  const groups = q ? all.filter((g) => `${g.brand} ${g.group} ${g.reason || ""}`.toLowerCase().includes(q)) : all;
  S.listed = groups;
  const bucketName = { target: "Целевые группы", small: "Мелкие группы", noprice: "Нет цен" };
  const bucketOrder = ["target", "small", "noprice"];
  const mode = S.groupBy === "brand" ? "brand" : "bucket";
  const byHead = new Map();
  groups.forEach((g, i) => {
    const head = mode === "brand" ? (g.brand || "без ТМ") : (bucketName[g.bucket] || "Целевые группы");
    if (!byHead.has(head)) byHead.set(head, []);
    byHead.get(head).push({ g, i });
  });
  const heads = mode === "bucket"
    ? bucketOrder.map((key) => bucketName[key]).filter((name) => byHead.has(name))
    : [...byHead.keys()];
  const body = groups.length ? heads.map((head) => `
    <div class="group-brand">${esc(head)}</div>
    ${byHead.get(head).map(({ g, i }) => {
      const mark = g.ready === true ? "ready" : g.ready === false ? "wait" : "";
      const ownRec = (S.recGroupSource || {})[g.key] === "mine" || (!((S.recGroupSource || {})[g.key]) && !!(g.key && S.recGroups[g.key]));
      const bits = [fmtInt(g.count)];
      if (g.own) bits.push("свой файл");
      if (g.price_gap === "cross") bits.push("кроссов нет");
      else if (g.price_gap === "price") bits.push("цен нет");
      else if (g.priced_rows) bits.push(`с ценой ${fmtInt(g.priced_rows)}`);
      if (g.reason) bits.push(g.reason);
      return `<button type="button" class="group-card ${mark} ${g.brand === S.opts.brand && g.group === S.opts.group ? "on" : ""}" data-i="${i}">
        <span class="dot"></span>
        <span class="t">${esc(g.group || "без группы")}</span>
        <span class="s">${esc(bits.join(" · "))}<span data-ownrec class="${ownRec ? "" : "hidden"}"> · свои правила</span></span>
      </button>`;
    }).join("")}`).join("")
    : `<div class="empty"><p>${q ? "Ничего не нашлось" : "В позиционировании нет групп из этой выгрузки"}</p></div>`;
  $("#groups").innerHTML = body;
  const readyN = all.filter((g) => g.ready).length;
  const hasReady = all.some((g) => "ready" in g);
  $("#group-count").textContent = q
    ? `${fmtInt(groups.length)} из ${fmtInt(all.length)}`
    : hasReady
      ? `${fmtInt(readyN)} с полным набором · ${fmtInt(all.length)} ${plural(all.length, "группа", "группы", "групп")}`
      : `${fmtInt(all.length)} ${plural(all.length, "группа", "группы", "групп")} из позиционирования`;
  $("#groups").querySelectorAll(".group-card").forEach((b) => (b.onclick = () => {
    const g = groups[+b.dataset.i];
    S.opts.brand = g.brand; S.opts.group = g.group; S.opts.cross_modes = {};
    S.analysisFor = "";
    applyGroupRec();
    renderGroup();
    renderRecForm();
    updateSummary();
    analyze();
  }));
  const groupBy = $("#group-by");
  if (groupBy && !groupBy.dataset.bound) {
    groupBy.dataset.bound = "1";
    seg(groupBy, mode, (v) => { S.groupBy = v; renderGroup(); });
  } else if (groupBy) {
    groupBy.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.v === mode));
  }
  const selected = $("#groups").querySelector(".group-card.on");
  if (selected && !q) selected.scrollIntoView({ block: "nearest" });
  const g = all.find((x) => x.brand === S.opts.brand && x.group === S.opts.group);
  const subs = g ? g.subgroups : [];
  $("#subgroups").innerHTML = subs.map((s) => `<span class="tag">${esc(s.name || "без подгруппы")} <b>${fmtInt(s.count)}</b></span>`).join("");
  $("#subgroup-note").textContent = subs.length > 1
    ? "На одном листе у каждой подгруппы свой % позиционирования — диски не получают процент колодок. «Лист на подгруппу» только раскладывает те же строки по вкладкам."
    : "В этой группе одна подгруппа, делить пока не на что. Когда появятся ещё — здесь выбирается лист на подгруппу или один общий. Процент в обоих случаях берётся по подгруппе строки.";
  seg($("#subgroup-mode"), S.opts.subgroup_mode, (v) => { S.opts.subgroup_mode = v; renderGroup(); persist(); });
  renderPosLines();
  renderContext();
}

function renderPosLines() {
  const el = $("#pos-lines");
  if (!el) return;
  const rows = (S.analysis && S.analysis.positioning_rows) || [];
  const by = new Map();
  rows.forEach((r) => {
    const sign = Number(r.percent) > 0 ? "+" : "";
    const text = `${r.name} ${sign}${String(r.percent).replace(".", ",")}%`;
    if (!by.has(r.subgroup)) by.set(r.subgroup, []);
    by.get(r.subgroup).push(text);
  });
  el.innerHTML = [...by.entries()].map(([sg, parts]) =>
    `<div class="pos-line"><b>${esc(sg || "без подгруппы")}</b> — ${esc(parts.join(" · "))}</div>`).join("");
}

function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
  return many;
}

// ------------------------------------------------------------------ шаг 3: фильтры

function renderSheetFilters() {
  const choices = S.init.sheet_filter_choices || [];
  const titles = Object.fromEntries(choices.map((c) => [c.key, c.title]));
  const selected = S.opts.sheet_filters || [];
  const chips = selected.map((key) => `<span class="tag">${esc(titles[key] || key)} <button type="button" class="btn small ghost" data-off="${esc(key)}" title="Убрать">✕</button></span>`).join("");
  const rest = choices.filter((c) => !selected.includes(c.key));
  const add = rest.length ? `<select class="input" id="sheet-filter-add"><option value="">Добавить…</option>${rest.map((c) => `<option value="${esc(c.key)}">${esc(c.title)}</option>`).join("")}</select>` : "";
  const el = $("#sheet-filters");
  if (!el) return;
  el.innerHTML = chips + add;
  el.querySelectorAll("[data-off]").forEach((b) => (b.onclick = () => {
    S.opts.sheet_filters = selected.filter((k) => k !== b.dataset.off);
    renderSheetFilters();
    persist();
  }));
  const menu = $("#sheet-filter-add");
  if (menu) menu.onchange = () => {
    if (!menu.value) return;
    S.opts.sheet_filters = [...selected, menu.value];
    renderSheetFilters();
    persist();
  };
}

function renderFilters() {
  const f = S.init.filters;
  $("#filters").innerHTML = Object.entries(f).map(([k, def]) => `
    <div class="filter"><div class="t">${esc(def.title)}</div>
      ${segHtml(def.choices, S.opts.filters[k] || def.default, `data-filter="${k}"`)}
    </div>`).join("");
  $("#filters").querySelectorAll("[data-filter]").forEach((el) => {
    const k = el.dataset.filter;
    seg(el, S.opts.filters[k] || f[k].default, (v) => {
      if (v === f[k].default) delete S.opts.filters[k]; else S.opts.filters[k] = v;
      renderFilters(); analyze();
    });
  });
  renderAbc();
}

function renderAbc() {
  const counts = S.analysis ? S.analysis.abc : {};
  const values = S.init.abc_values.filter((v) => counts[v] || S.opts.abc.includes(v) || !S.analysis);
  const all = !S.opts.abc.length;
  $("#abc").innerHTML = [`<button class="chip ${all ? "on" : ""}" data-v="*">Все</button>`,
    ...values.map((v) => `<button class="chip ${!all && S.opts.abc.includes(v) ? "on" : ""}" data-v="${esc(v)}">${v || "пусто"}${counts[v] != null ? ` <span class="count">${fmtInt(counts[v])}</span>` : ""}</button>`)].join("");
  $("#abc").querySelectorAll(".chip").forEach((b) => (b.onclick = () => {
    const v = b.dataset.v;
    if (v === "*") S.opts.abc = [];
    else if (S.opts.abc.includes(v)) S.opts.abc = S.opts.abc.filter((x) => x !== v);
    else S.opts.abc = [...S.opts.abc, v];
    renderAbc(); analyze();
  }));
}

function renderCounter() {
  const a = S.analysis;
  const reasons = a.reasons.map(([r, n]) => `<span class="tag">${esc(r)} <b>${fmtInt(n)}</b></span>`).join("");
  $("#counter").innerHTML = `
    <div><div class="big">${fmtInt(a.kept)}</div><div class="muted" style="font-size:12.5px">из ${fmtInt(a.total)} в отчёт</div></div>
    <div><div class="big" style="color:var(--amber)">${fmtInt(a.priced)}</div><div class="muted" style="font-size:12.5px">с ценой конкурента</div></div>
    ${a.reasons.length ? `<div class="stack" style="gap:4px"><div class="muted" style="font-size:12.5px">Не попадут:</div><div class="reasons">${reasons}</div></div>` : `<div class="muted">Фильтры ничего не убирают</div>`}
    ${a.no_positioning ? `<div class="note" style="margin-left:auto">${fmtInt(a.no_positioning)} товаров без позиционирования — для них конкурентов не будет</div>` : ""}`;
}

// ------------------------------------------------------------------ шаг 4: кроссы

function renderSplit() {
  const el = $("#cross-split");
  el.innerHTML = Object.entries(S.init.cross_splits).map(([k, t]) => `<button data-v="${k}">${esc(t)}</button>`).join("");
  seg(el, S.opts.cross_split, (v) => { S.opts.cross_split = v; renderSplit(); analyze(); });
}

function stackBar(s) {
  const w = (n) => `${s.items ? (n / s.items) * 100 : 0}%`;
  return `<div class="stackbar" title="«Да»: ${s.yes} · только «Нет»: ${s.no_only} · без кроссов: ${s.none}">
    <i class="yes" style="width:${w(s.yes)}"></i><i class="no" style="width:${w(s.no_only)}"></i><i class="none" style="width:${w(s.none)}"></i></div>`;
}

function crossCounts(s) {
  return `<span class="cnt yes">${fmtInt(s.yes)}</span><span class="cnt no">${fmtInt(s.no_only)}</span><span class="cnt none">${fmtInt(s.none)}</span>`;
}

function modeAvailable(mode) {
  return S.analysis.has_svod || mode === "yes" || mode === "off";
}

function modePicker(s, key) {
  const btns = Object.keys(S.init.cross_modes).filter(modeAvailable).map((m) => {
    const label = !S.analysis.has_svod && m === "yes" ? "Брать из АЦПКД" : S.init.cross_modes[m];
    const n = m === "off" ? "" : `<small>${fmtInt(s.priced[m] ?? 0)}</small>`;
    const tip = `${S.init.cross_mode_hints[m]}${m !== "off" ? ` — цена будет у ${s.priced[m] ?? 0} из ${s.items}` : ""}`;
    return `<button data-v="${m}" class="${m === s.mode ? "on" : ""} ${m === s.suggested ? "sugg" : ""}" title="${esc(tip)}">${esc(label)}${n}</button>`;
  }).join("");
  return `<div class="seg modes" data-key="${esc(key)}">${btns}</div>`;
}

function coverage(priced, items) {
  return `<div class="cov"><span class="v">${fmtInt(priced)} <span class="muted" style="font-weight:500">из ${fmtInt(items)}</span></span><div class="bar"><i style="width:${pct(priced, items)}%"></i></div></div>`;
}

function segLabel(s, split) {
  if (split === "abc") return `<span class="abc-badge">${esc(s.abc || "пусто")}</span>`;
  if (split === "subgroup") return `<span class="seg-name">${esc(s.subgroup || "без подгруппы")}</span>`;
  if (split === "subgroup_abc") return `<span class="abc-badge">${esc(s.abc || "пусто")}</span>`;
  return "";
}

function setMode(comp, segKey, mode) {
  if (segKey === "*") {
    Object.keys(S.opts.cross_modes).filter((k) => k.startsWith(`${comp.key}|`)).forEach((k) => delete S.opts.cross_modes[k]);
    S.opts.cross_modes[`${comp.key}|*`] = mode;
    comp.segments.forEach((s) => (s.mode = mode));
  } else {
    S.opts.cross_modes[`${comp.key}|${segKey}`] = mode;
    comp.segments.find((s) => s.seg === segKey).mode = mode;
  }
  comp.priced = comp.segments.reduce((sum, s) => sum + (s.priced[s.mode] ?? 0), 0);
  renderCrosses();
  analyze();
}

function renderCrosses() {
  const a = S.analysis;
  if (!a.crosses.length) {
    $("#crosses").innerHTML = `<div class="empty"><h3>Конкурентов нет</h3><p>Для выбранной группы в позиционировании (${esc(a.positioning)}) не указаны бренды-конкуренты.</p></div>`;
    return;
  }
  const split = a.cross_split;
  const notes = [];
  if (!a.has_svod) notes.push(`<div class="note">Файла «Кроссы свод» нет — кроссы берутся из АЦПКД, признака «Да»/«Нет» у них нет.</div>`);
  const html = a.crosses.map((c) => {
    const modes = new Set(c.segments.map((s) => s.mode));
    const common = modes.size === 1 ? [...modes][0] : null;
    const allPriced = Object.fromEntries(Object.keys(S.init.cross_modes).map((m) => [m, c.segments.reduce((n, s) => n + (s.priced[m] ?? 0), 0)]));
    const suggested = new Set(c.segments.map((s) => s.suggested));
    const head = { ...c, mode: common, priced: allPriced, suggested: suggested.size === 1 ? [...suggested][0] : null };
    const collapsed = split !== "none" && S.collapsed[c.key];
    const manual = c.segments.filter((s) => s.mode !== s.suggested).length;
    let body = "";
    if (split === "none") {
      const s = c.segments[0];
      body = `<div class="verdict-line">${esc(s.verdict)}</div>`;
    } else if (!collapsed) {
      let lastSub = null;
      body = `<div class="seg-rows">${c.segments.map((s) => {
        let sub = "";
        if (split === "subgroup_abc" && s.subgroup !== lastSub) {
          lastSub = s.subgroup;
          sub = `<div class="seg-sub">${esc(s.subgroup || "без подгруппы")}</div>`;
        }
        return `${sub}<div class="seg-row ${s.mode !== s.suggested ? "changed" : ""}">
          <div class="seg-label">${segLabel(s, split)}<span class="muted">${fmtInt(s.items)} тов.</span>
            <span class="info" title="${esc(s.verdict)}">?</span></div>
          <div class="seg-mix">${stackBar(s)}<div class="mix-nums">${crossCounts(s)}</div></div>
          <div>${modePicker(s, `${c.key}|${s.seg}`)}</div>
          ${coverage(s.priced[s.mode] ?? 0, s.items)}
        </div>`;
      }).join("")}</div>`;
    }
    return `<div class="comp-block">
      <div class="comp-head">
        <div class="comp-title">
          ${split !== "none" ? `<button class="caret ${collapsed ? "" : "open"}" data-toggle="${esc(c.key)}" title="Свернуть / развернуть">▸</button>` : ""}
          <span class="ord">${c.order}</span><b>${esc(c.name)}</b>
          <span class="muted">${fmtInt(c.items)} тов.</span>
          ${manual && split !== "none" ? `<span class="pill violet" title="Частей, где выбор отличается от подсказки">изменено: ${manual}</span>` : ""}
        </div>
        <div class="seg-mix">${stackBar(c)}<div class="mix-nums">${crossCounts(c)}${c.yes_articles ? `<span class="muted">· ${fmtInt(c.yes_articles)} разных «Да»</span>` : ""}</div></div>
        <div>${split !== "none" ? `<div class="all-label">${common ? "для всех частей" : "разные режимы — выбрать для всех"}</div>` : ""}${modePicker(head, `${c.key}|*`)}</div>
        ${coverage(c.priced, c.items)}
      </div>
      ${body}
    </div>`;
  }).join("");
  $("#crosses").innerHTML = `${notes.join("")}<div class="comp-list">${html}</div>`;
  $("#crosses").querySelectorAll(".seg.modes").forEach((el) => {
    const [compKey, segKey] = el.dataset.key.split("|");
    const comp = a.crosses.find((c) => c.key === compKey);
    el.querySelectorAll("button").forEach((b) => (b.onclick = () => setMode(comp, segKey, b.dataset.v)));
  });
  $("#crosses").querySelectorAll("[data-toggle]").forEach((b) => (b.onclick = () => {
    S.collapsed[b.dataset.toggle] = !S.collapsed[b.dataset.toggle];
    renderCrosses();
  }));
}

// ------------------------------------------------------------------ шаг 4: цены и прайсы

const pfLabel = (f) => (f.city ? `${f.supplier} · ${f.city}` : f.supplier);
const STATUS = {
  auto: ["Распознан сам", "good"],
  saved: ["По ручной настройке", "good"],
  manual: ["Настроен вручную", "violet"],
  none: ["Не распознан", "bad"],
};

function renderPriceControls() {
  const el = $("#price-source");
  el.innerHTML = Object.entries(S.init.price_sources).map(([k, t]) => `<button data-v="${k}">${esc(t)}</button>`).join("");
  seg(el, S.opts.price_source, (v) => { S.opts.price_source = v; renderPriceControls(); renderPriceFiles(); analyze(); });
  const pick = $("#price-pick");
  if (pick) {
    pick.innerHTML = Object.entries(S.init.price_picks || { min: "Самая низкая" }).map(([k, t]) => `<button type="button" data-v="${k}">${esc(t)}</button>`).join("");
    seg(pick, S.opts.price_pick || "min", (v) => { S.opts.price_pick = v; analyze(); });
  }
  $("#monitoring-only").checked = S.opts.monitoring_only;
  $("#in-stock-only").checked = S.opts.in_stock_only;
  const a = S.analysis;
  const offers = a ? a.offers || {} : {};
  const parts = [];
  if (S.opts.price_source !== "prices") parts.push(`АЦПКД: ${fmtInt(offers["АЦПКД"] || 0)} цен`);
  if (S.opts.price_source !== "acp") parts.push(`прайсы: ${fmtInt(offers["Прайсы"] || 0)} цен`);
  $("#price-source-note").textContent = a ? parts.join(" · ") : "";
  const notes = [];
  if (a && a.skipped_by_monitoring && S.opts.price_source !== "acp") {
    notes.push(`${fmtInt(a.skipped_by_monitoring)} ${plural(a.skipped_by_monitoring, "цена из прайсов отброшена", "цены из прайсов отброшены", "цен из прайсов отброшено")}: поставщик не мониторит этот бренд по позиционированию`);
  }
  if (a && a.skipped_low) {
    notes.push(`${fmtInt(a.skipped_low)} ${plural(a.skipped_low, "аномально низкая цена не берётся", "аномально низкие цены не берутся", "аномально низких цен не берутся")}: 1–2 ₽ и цены в разы ниже остальных по этому артикулу`);
  }
  $("#monitoring-note").textContent = notes.join(" ");
}

function latestPriceDate(files) {
  const dates = [...new Set((files || []).map((f) => f.date).filter(Boolean))];
  dates.sort((a, b) => dateKey(a) - dateKey(b));
  return dates.length ? dates[dates.length - 1] : "";
}

function dateKey(text) {
  const m = String(text || "").match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  return m ? Number(`${m[3]}${m[2]}${m[1]}`) : 99999999;
}

function renderPriceFiles() {
  const latest = latestPriceDate(S.priceFiles);
  const files = latest ? S.priceFiles.filter((f) => !f.date || f.date === latest) : S.priceFiles;
  const off = S.opts.price_source === "acp";
  if (!files.length) {
    $("#price-files").innerHTML = `<div class="empty compact"><h3>Прайсов в папке нет</h3>
      <p>Положите прайсы в Данные/Прайсы/Поставщик — папка с именем поставщика, внутри его файлы. Можно и прямо в «Прайсы»: тогда поставщик угадывается по имени файла. Столбцы программа ищет сама.</p></div>`;
    return;
  }
  const bySupplier = new Map();
  files.forEach((f) => {
    const key = f.supplier || "Без поставщика";
    if (!bySupplier.has(key)) bySupplier.set(key, []);
    bySupplier.get(key).push(f);
  });
  const card = (f) => {
    const [st, kind] = STATUS[f.status] || STATUS.none;
    const brands = f.competitor_brands.slice(0, 4).map((b) => `<span class="tag">${esc(b.brand)} <b>${fmtInt(b.offers)}</b></span>`).join("");
    const editing = S.editing && S.editWhere !== "prices" && S.editing.id === f.id;
    return `<div class="pf ${f.enabled ? "" : "disabled"} ${editing ? "editing" : ""}">
      <div class="pf-row">
        <label class="switch" title="Использовать прайс"><input type="checkbox" data-enable="${esc(f.id)}" ${f.enabled ? "checked" : ""}><span></span></label>
        <div class="pf-main">
          <div class="pf-title"><b>${esc(pfLabel(f))}</b>
            ${f.monitored ? "" : `<span class="pill warn" title="Поставщика нет в мониторинге позиционирования">нет в позиционировании</span>`}</div>
          <div class="pf-file" title="${esc(f.id)}">${esc(f.id)}</div>
        </div>
        <div class="pf-stat"><span class="pill ${kind}">${st}</span>
          <span class="muted">${fmtInt(f.offers)} цен${f.header_row ? ` · шапка в строке ${f.header_row}` : ""}</span></div>
        <div class="pf-brands">${brands || `<span class="muted">брендов конкурентов не найдено</span>`}</div>
        <button class="btn small" data-edit="${esc(f.id)}">${editing ? "Закрыть" : "Настроить"}</button>
      </div>
      ${f.problem ? `<div class="pf-problem">${esc(f.problem)}</div>` : ""}
      ${editing ? editorHtml() : ""}
    </div>`;
  };
  if (!S.priceOpen) S.priceOpen = {};
  const blocks = [...bySupplier.entries()].map(([supplier, list]) => {
    const offers = list.reduce((sum, f) => sum + (f.offers || 0), 0);
    const open = list.length === 1 || S.priceOpen[supplier] || list.some((f) => S.editing && S.editing.id === f.id);
    if (list.length === 1) return card(list[0]);
    return `<button type="button" class="pf-supplier" data-supplier="${esc(supplier)}">${open ? "−" : "+"} ${esc(supplier)} <span>${fmtInt(list.length)} прайсов · ${fmtInt(offers)} цен</span></button>
      ${open ? `<div class="pf-list">${list.map(card).join("")}</div>` : ""}`;
  }).join("");
  $("#price-files").innerHTML = `${off ? `<div class="note info">Сейчас цены берутся только из АЦПКД — прайсы не используются.</div>` : ""}
    ${latest ? `<div class="muted small-text" style="margin-bottom:8px">Прайсы на <b>${esc(latest)}</b> — последнюю дату выгрузки.</div>` : ""}
    <div class="pf-list ${off ? "dim" : ""}">${blocks}</div>
    <div class="muted small-text" style="margin-top:8px">Прайсы одного поставщика свёрнуты. В отчёт попадает минимальная цена конкурента.</div>`;
  bindPriceFiles();
  $("#price-files").querySelectorAll("[data-supplier]").forEach((b) => (b.onclick = () => {
    S.priceOpen[b.dataset.supplier] = !S.priceOpen[b.dataset.supplier];
    renderPriceFiles();
  }));
  if (S.tab === "data") renderPriceAudit();
}

const PX_FROM = { folder: "папка", file: "имя файла", manual: "вручную" };

function supplierKey(v) {
  return String(v ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().toLowerCase().replace(/ё/g, "е").replace(/[^0-9a-zа-я]/g, "");
}

function brandKey(v) {
  return String(v ?? "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().toUpperCase().replace(/Ё/g, "Е").replace(/[^0-9A-ZА-Я]/g, "");
}

function priceBad(f) {
  return !!(f.problem || f.doubt || f.status === "none" || !f.offers || f.mismatch || (f.quality && f.quality.length));
}

function priceGroups(files, monBy) {
  const groups = new Map();
  files.forEach((f) => {
    const key = supplierKey(f.supplier) || "—";
    if (!groups.has(key)) {
      const known = monBy.get(key);
      groups.set(key, { key, supplier: (known && known.supplier) || f.supplier || "Без поставщика", files: [] });
    }
    groups.get(key).files.push(f);
  });
  return [...groups.values()].map((g) => {
    const expected = (monBy.get(g.key) && monBy.get(g.key).brands) || g.files[0].expected_brands || [];
    const found = new Map();
    g.files.forEach((f) => (f.competitor_brands || []).forEach((b) => {
      const k = brandKey(b.brand);
      found.set(k, (found.get(k) || 0) + (b.offers || 0));
    }));
    g.found = found;
    g.brands = expected.map((name) => ({ name, offers: found.get(brandKey(name)) || 0, hit: found.has(brandKey(name)) }));
    g.expected = expected;
    g.bad = g.files.filter(priceBad).length;
    g.offers = g.files.reduce((sum, f) => sum + (f.offers || 0), 0);
    g.monitored = monBy.has(g.key);
    return g;
  });
}

function fileState(f) {
  if (f.status === "none" || !f.offers || f.problem) return "bad";
  if (f.doubt || f.mismatch || (f.quality && f.quality.length)) return "warn";
  return "good";
}

function renderCalendar(files, waiting) {
  const el = $("#dt-calendar");
  if (!el) return;
  const dates = [...new Set([...files.map((f) => f.date || "без даты"), ...waiting.map((w) => w.date || "без даты")])]
    .sort((a, b) => dateKey(a) - dateKey(b)).slice(-21);
  if (!dates.length) { el.innerHTML = ""; return; }
  const rows = new Map();
  const put = (supplier, date, item) => {
    const key = supplierKey(supplier) || "—";
    if (!rows.has(key)) rows.set(key, { name: supplier || "—", cells: {} });
    const cell = rows.get(key).cells[date] || (rows.get(key).cells[date] = { files: 0, bad: 0, warn: 0, wait: 0, offers: 0 });
    cell.files += 1;
    if (item === "wait") cell.wait += 1;
    else {
      const st = fileState(item);
      if (st === "bad") cell.bad += 1;
      if (st === "warn") cell.warn += 1;
      cell.offers += item.offers || 0;
    }
  };
  files.forEach((f) => put(f.supplier, f.date || "без даты", f));
  waiting.forEach((w) => put(w.supplier, w.date || "без даты", "wait"));
  const names = [...rows.values()].sort((a, b) => a.name.localeCompare(b.name, "ru"));
  const short = (d) => d.length === 10 ? d.slice(0, 5) : d;
  const head = `<tr><th>Поставщик</th>${dates.map((d) => `<th class="num" title="${esc(d)}">${esc(short(d))}</th>`).join("")}</tr>`;
  const body = names.map((r) => `<tr><td class="nowrap"><b>${esc(r.name)}</b></td>${dates.map((d) => {
    const c = r.cells[d];
    if (!c) return `<td class="cal-cell empty" title="${esc(d)}: файлов нет">·</td>`;
    const kind = c.bad ? "bad" : c.warn ? "warn" : c.wait === c.files ? "wait" : "good";
    const tip = `${d}: ${c.files} ${plural(c.files, "файл", "файла", "файлов")}${c.wait ? `, не разобрано ${c.wait}` : ""}${c.bad ? `, с проблемой ${c.bad}` : ""}${c.warn ? `, с замечанием ${c.warn}` : ""}${c.offers ? ` · ${fmtInt(c.offers)} цен` : ""}`;
    return `<td class="cal-cell ${kind}" data-cal-date="${esc(d)}" data-cal-sup="${esc(r.name)}" title="${esc(tip)}">${c.files}</td>`;
  }).join("")}</tr>`).join("");
  const anyWait = waiting.length;
  el.innerHTML = `<div class="cal-legend"><span><i class="good"></i>распознаны</span><span><i class="warn"></i>есть замечание</span><span><i class="bad"></i>есть проблема</span><span><i class="wait"></i>прошлая дата, ещё не разобрана</span>
      ${anyWait ? `<button type="button" class="btn small" id="dt-wake">Разобрать прошлые даты · ${fmtInt(anyWait)}</button>` : ""}</div>
    <div class="table-wrap cal"><table class="mini"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
  el.querySelectorAll("[data-cal-date]").forEach((td) => (td.onclick = () => {
    S.pxDateOpen = { [td.dataset.calDate]: true };
    S.pxOpen = S.pxOpen || {};
    S.pxOpen[`${td.dataset.calDate}|${supplierKey(td.dataset.calSup)}`] = true;
    S.pxQuery = td.dataset.calSup;
    const search = $("#px-search");
    if (search) search.value = td.dataset.calSup;
    renderPriceAudit();
    const block = document.querySelector(`[data-date-block="${CSS.escape(td.dataset.calDate)}"]`);
    if (block) block.scrollIntoView({ behavior: "smooth", block: "start" });
  }));
  const wake = $("#dt-wake");
  if (wake) wake.onclick = async () => { await loadAnalytics(); renderPriceAudit(); };
}

function renderPriceAudit() {
  const kpis = $("#px-kpis");
  const list = $("#px-list");
  if (!kpis || !list) return;
  renderBrandHints();
  const filterEl = $("#px-filter");
  const filter = S.pxFilter || "all";
  if (filterEl) filterEl.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.v === filter));
  const files = S.priceFiles || [];
  const waiting = ((S.scan && S.scan.price_waiting) || []).filter((w) => !files.some((f) => f.date === w.date && f.file === w.file));
  const mon = (S.scan && S.scan.monitoring_suppliers) || [];
  const monBy = new Map(mon.map((m) => [supplierKey(m.supplier), m]));
  if (!S.scan) {
    kpis.innerHTML = "";
    list.innerHTML = `<div class="card"><div class="empty compact"><h3>Папка ещё не прочитана</h3><p>Нажмите «Обновить данные».</p></div></div>`;
    return;
  }
  renderCalendar(files, waiting);
  const latest = latestPriceDate(files);
  const latestFiles = latest ? files.filter((f) => !f.date || f.date === latest) : files;
  const latestGroups = new Map(priceGroups(latestFiles, monBy).map((g) => [g.key, g]));
  let covered = 0;
  let pairs = 0;
  mon.forEach((m) => {
    const g = latestGroups.get(supplierKey(m.supplier));
    m.brands.forEach((brand) => {
      pairs += 1;
      if (g && g.found.has(brandKey(brand))) covered += 1;
    });
  });
  const badFiles = files.filter(priceBad).length;
  const recognized = files.filter((f) => f.status !== "none" && f.offers > 0).length;
  const suppliers = new Set(files.map((f) => supplierKey(f.supplier)));
  kpis.innerHTML = [
    ["Файлов", fmtInt(files.length + waiting.length), waiting.length ? `${fmtInt(waiting.length)} прошлых дат ещё не разобрано` : "все разобраны"],
    ["Распознано", fmtInt(recognized), badFiles ? `${fmtInt(badFiles)} с проблемой или замечанием` : "без замечаний"],
    ["Поставщиков", fmtInt(suppliers.size), "по папке или имени файла"],
    ["Цен на последнюю дату", fmtInt(latestFiles.reduce((s, f) => s + (f.offers || 0), 0)), latest || "—"],
    ["Пары поставщик–бренд", pairs ? `${fmtInt(covered)} из ${fmtInt(pairs)}` : "—", "из позиционирования есть в прайсах последней даты"],
  ].map(([l, v, sub]) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${esc(sub)}</div></div>`).join("");
  if (!files.length && !waiting.length) {
    list.innerHTML = `<div class="card"><div class="empty compact"><h3>Прайсов пока нет</h3>
      <p>Создайте в Данные/Прайсы папку-дату, внутри — папку с именем поставщика, как в позиционировании, и положите туда его файлы. Например: Данные/Прайсы/14.10.2026/TISS/TissEKB.xls.</p></div></div>`;
    return;
  }
  if (!S.pxOpen) S.pxOpen = {};
  if (!S.pxDateOpen) S.pxDateOpen = {};
  const q = String(S.pxQuery || "").trim().toLowerCase();
  const hitFile = (f) => !q || [f.file, f.id, f.city, f.problem, f.folder, f.name_supplier, f.supplier,
    ...(f.competitor_brands || []).map((b) => b.brand)].join(" ").toLowerCase().includes(q);
  const byDate = new Map();
  files.forEach((f) => {
    const d = f.date || "без даты";
    if (!byDate.has(d)) byDate.set(d, []);
    byDate.get(d).push(f);
  });
  const waitBy = new Map();
  waiting.forEach((w) => {
    const d = w.date || "без даты";
    if (!waitBy.has(d)) waitBy.set(d, []);
    waitBy.get(d).push(w);
  });
  const dates = [...new Set([...byDate.keys(), ...waitBy.keys()])].sort((a, b) => dateKey(b) - dateKey(a));
  const blocks = dates.map((d, n) => {
    const dayFiles = byDate.get(d) || [];
    const dayWait = waitBy.get(d) || [];
    const groups = priceGroups(dayFiles, monBy).map((g) => {
      const hits = g.files.filter((f) => hitFile(f) || (q && g.supplier.toLowerCase().includes(q)));
      return { ...g, files: hits.length ? hits : g.files, show: (!q || hits.length || g.brands.some((b) => b.name.toLowerCase().includes(q))) && matchPxFilter(g, filter) };
    }).filter((g) => g.show);
    groups.sort((a, b) => a.supplier.localeCompare(b.supplier, "ru"));
    const waitShown = dayWait.filter((w) => !q || `${w.supplier} ${w.file}`.toLowerCase().includes(q));
    if (!groups.length && !waitShown.length) return "";
    const bad = dayFiles.filter(priceBad).length;
    const open = Object.prototype.hasOwnProperty.call(S.pxDateOpen, d) ? S.pxDateOpen[d] : (n === 0 || (!!q && groups.length > 0));
    const sup = new Set(dayFiles.map((f) => supplierKey(f.supplier))).size;
    const head = `<button type="button" class="date-head" data-date="${esc(d)}" data-open="${open ? "1" : "0"}">
      <span>${open ? "▾" : "▸"}</span><b>${esc(d)}</b>
      <span class="muted">${fmtInt(dayFiles.length + dayWait.length)} ${plural(dayFiles.length + dayWait.length, "файл", "файла", "файлов")} · ${fmtInt(sup || new Set(dayWait.map((w) => w.supplier)).size)} ${plural(sup, "поставщик", "поставщика", "поставщиков")}${dayFiles.length ? ` · ${fmtInt(dayFiles.reduce((s, f) => s + (f.offers || 0), 0))} цен` : ""}</span>
      ${bad ? `<span class="pill bad">${fmtInt(bad)} с замечанием</span>` : ""}
      ${dayWait.length ? `<span class="pill">не разобрано ${fmtInt(dayWait.length)}</span>` : ""}
      ${d === latest ? `<span class="pill good">для отчётов</span>` : ""}
    </button>`;
    const body = open ? `${groups.map((g) => pxCard(g, d)).join("")}
      ${waitShown.length ? `<div class="card"><div class="muted">Ещё не разобраны (прошлая дата читается, когда нужна динамика): ${waitShown.map((w) => `${esc(w.supplier)} — ${esc(w.file)}`).join(", ")}</div></div>` : ""}` : "";
    return `<div class="date-block" data-date-block="${esc(d)}">${head}${body}</div>`;
  }).join("");
  const present = new Set(files.map((f) => supplierKey(f.supplier)));
  const gaps = mon.filter((m) => !present.has(supplierKey(m.supplier))).filter((m) => {
    if (filter !== "all" && filter !== "gap") return false;
    return !q || `${m.supplier} ${m.brands.join(" ")}`.toLowerCase().includes(q);
  });
  const gapCard = gaps.length ? `<div class="card">
    <div class="card-head"><h2>В позиционировании, прайса нет</h2><span class="sub">${fmtInt(gaps.length)}</span></div>
    <div class="px-brands" style="margin-bottom:10px">Эти поставщики мониторят бренды по позиционированию, а их папки или файла программа не увидела ни в одной дате.</div>
    ${gaps.map((m) => `<div style="margin-bottom:8px"><b>${esc(m.supplier)}</b>
      <span class="px-brands">${m.brands.map((b) => `<span class="tag miss">${esc(b)}</span>`).join(" ")}</span></div>`).join("")}
  </div>` : "";
  const shown = filter === "gap" ? "" : blocks;
  list.innerHTML = (shown || gapCard) ? shown + gapCard
    : `<div class="card"><div class="empty compact"><h3>По этому отбору пусто</h3><p>Смените фильтр или очистите поиск.</p></div></div>`;
  list.querySelectorAll("[data-date]").forEach((b) => (b.onclick = () => {
    S.pxDateOpen[b.dataset.date] = b.dataset.open !== "1";
    renderPriceAudit();
  }));
  list.querySelectorAll("[data-px]").forEach((b) => (b.onclick = () => {
    S.pxOpen[b.dataset.px] = b.dataset.open !== "1";
    renderPriceAudit();
  }));
  list.querySelectorAll("[data-px-edit]").forEach((b) => (b.onclick = (e) => {
    e.stopPropagation();
    const f = (S.priceFiles || []).find((x) => x.id === b.dataset.pxEdit);
    if (!f) return;
    if (S.editing && S.editWhere === "prices" && S.editing.id === f.id) {
      S.editing = null;
      renderPriceFiles();
    } else openEditor(f, "prices");
  }));
  if (S.editing && S.editWhere === "prices") bindEditor(list);
}

function matchPxFilter(g, filter) {
  if (filter === "gap") return false;
  if (filter === "problem") return g.bad > 0;
  if (filter === "missing") return g.brands.some((b) => !b.hit);
  if (filter === "nomon") return !g.monitored;
  return true;
}

function pxCard(g, date) {
  const key = `${date}|${g.key}`;
  const trouble = g.bad > 0 || g.brands.some((b) => !b.hit);
  const editingHere = !!(S.editWhere === "prices" && S.editing && g.files.some((f) => f.id === S.editing.id));
  const open = editingHere || (Object.prototype.hasOwnProperty.call(S.pxOpen, key) ? S.pxOpen[key] : trouble);
  const brands = g.brands.length
    ? g.brands.map((b) => `<span class="tag ${b.hit ? "hit" : "miss"}">${esc(b.name)}${b.hit ? ` <b>${fmtInt(b.offers)}</b>` : ""}</span>`).join("")
    : `<span class="muted">этого поставщика нет в позиционировании</span>`;
  const body = open ? `<div class="table-wrap"><table>
    <thead><tr><th>Файл</th><th>Город</th><th>Как читается</th><th>Статус</th><th class="num">Цен</th><th>Бренды в файле</th><th>Замечание</th><th></th></tr></thead>
    <tbody>${g.files.map(pxRow).join("")}</tbody></table></div>` : "";
  return `<div class="card px-card">
    <button type="button" class="px-head" data-px="${esc(key)}" data-open="${open ? "1" : "0"}">
      <span>${open ? "▾" : "▸"}</span>
      <b>${esc(g.supplier)}</b>
      ${g.monitored ? `<span class="pill good">в мониторинге</span>` : `<span class="pill warn">нет в позиционировании</span>`}
      <span class="muted">${fmtInt(g.files.length)} ${plural(g.files.length, "файл", "файла", "файлов")} · ${fmtInt(g.offers)} цен</span>
      ${g.bad ? `<span class="pill bad">${fmtInt(g.bad)} с замечанием</span>` : ""}
    </button>
    <div class="px-brands">${brands}</div>
    ${body}
  </div>`;
}

function pxRow(f) {
  const [st, kind] = STATUS[f.status] || STATUS.none;
  const fields = f.fields || {};
  const cols = [["article", "артикул"], ["price", "цена"], ["brand", "бренд"], ["qty", "наличие"]]
    .filter(([k]) => fields[k]).map(([, t]) => t).join(", ");
  const brands = (f.competitor_brands || []).slice(0, 6).map((b) => `${esc(b.brand)} ${fmtInt(b.offers)}`).join(", ");
  const from = PX_FROM[f.supplier_from] || f.supplier_from || "";
  const notes = [];
  if (f.mismatch) notes.push(`имя файла похоже на ${f.name_supplier}`);
  if (f.supplier_from === "folder" && f.folder && supplierKey(f.folder) !== supplierKey(f.supplier)) notes.push(`папка «${f.folder}»`);
  if (f.problem) notes.push(f.problem);
  else if (f.why) notes.push(f.why);
  if (f.doubt) notes.push(f.doubt);
  (f.quality || []).forEach((q) => notes.push(q));
  if (!f.enabled) notes.push("выключен");
  const info = (f.info || []).map((q) => `<div class="info-note">${esc(q)}</div>`).join("");
  const where = f.shape || (f.header_row ? `шапка, строка ${f.header_row}${cols ? ` · ${cols}` : ""}` : (f.problem ? "" : "шапка не найдена"));
  const editing = S.editing && S.editWhere === "prices" && S.editing.id === f.id;
  return `<tr class="${priceBad(f) ? "px-bad" : ""}">
    <td><b>${esc(f.file)}</b><div class="muted">${esc([f.date, f.sheet].filter(Boolean).join(" · "))}</div></td>
    <td>${esc(f.city) || "—"}</td>
    <td>${esc(from)}${where ? `<div class="muted">${esc(where)}</div>` : ""}</td>
    <td><span class="pill ${kind}">${st}</span></td>
    <td class="num">${fmtInt(f.offers)}</td>
    <td>${brands || `<span class="muted">${f.offers ? "брендов из позиционирования нет" : "—"}</span>`}</td>
    <td>${notes.length ? esc(notes.join(". ")) : (info ? "" : "—")}${info}</td>
    <td><button type="button" class="btn small" data-px-edit="${esc(f.id)}">${editing ? "Закрыть" : "Править"}</button></td>
  </tr>${editing ? `<tr class="ed-row"><td colspan="8">${editorHtml()}</td></tr>` : ""}`;
}

function columnSense(file) {
  const checks = file.checks || {};
  const bit = (key, title) => {
    const row = checks[key];
    if (!row) return "";
    return `${title} ${row.letter}: ${Math.round((row.numeric || 0) * 100)}% чисел`;
  };
  const parts = [bit("price", "цена"), bit("name", "наименование"), bit("article", "артикул")].filter(Boolean);
  const head = parts.join(" · ");
  if (file.doubt) return head ? `${head}. ${file.doubt}` : file.doubt;
  return head || "Цена берётся только из столбца, где почти все значения — числа.";
}

function editorHtml() {
  const { draft, file } = S.editing;
  const fields = S.init.price_fields;
  const colOptions = (fld) => [`<option value="">— нет —</option>`, ...file.columns.map((c) =>
    `<option value="${c.index}" ${String(draft.cols[fld] ?? "") === String(c.index) ? "selected" : ""}>${esc(c.letter)} · ${esc(c.title || "без названия")}${c.sample.length ? ` — ${esc(c.sample.slice(0, 2).join(", "))}` : ""}</option>`)].join("");
  const headRows = file.head_rows.map((r) => `<option value="${r.row}" ${r.row === draft.header_row ? "selected" : ""}>${r.row}: ${esc(r.text.slice(0, 90))}</option>`).join("");
  const preview = file.preview.length
    ? `<table class="mini"><thead><tr><th>Бренд</th><th>Артикул</th><th class="num">Цена</th><th class="num">Наличие</th></tr></thead><tbody>
      ${file.preview.map((o) => `<tr><td>${esc(o.brand) || '<span class="muted">—</span>'}</td><td><b>${esc(o.article)}</b></td><td class="num price">${fmtMoney(o.price)}</td><td class="num">${o.qty == null ? "—" : fmtInt(o.qty)}</td></tr>`).join("")}
      </tbody></table>`
    : `<div class="muted">Строк с артикулом и ценой не найдено</div>`;
  return `<div class="pf-editor">
    <div class="ed-grid">
      <div class="ed-col">
        <div class="ed-pair"><label>Поставщик</label><input class="input" data-d="supplier" value="${esc(draft.supplier)}"></div>
        <div class="ed-pair"><label>Город / склад</label><input class="input" data-d="city" value="${esc(draft.city)}"></div>
        <div class="ed-pair"><label>Строка шапки</label><select class="input" data-d="header_row">${headRows}</select></div>
        ${Object.entries(fields).map(([k, t]) => `<div class="ed-pair"><label>${esc(t)}${["article", "price"].includes(k) ? " *" : ""}</label><select class="input" data-col="${k}">${colOptions(k)}</select></div>`).join("")}
        <div class="ed-pair"><label>Бренд для всего прайса</label><input class="input" data-d="fixed_brand" value="${esc(draft.fixed_brand)}" placeholder="если в прайсе нет столбца бренда"></div>
      </div>
      <div class="ed-col">
        <div class="ed-preview-head"><b>Что получится</b><span class="muted">${S.editing.pending ? "Пересчитываю по первым строкам…" : file.preview_only ? `${fmtInt(file.offers)} цен в первых ${fmtInt(Math.min(400, file.rows || 400))} строках из ${fmtInt(file.rows)} · ${fmtInt(file.brands)} брендов` : `${fmtInt(file.offers)} цен · ${fmtInt(file.brands)} брендов`}</span></div>
        <div class="muted">${esc(columnSense(file))}</div>
        ${file.problem ? `<div class="note">${esc(file.problem)}</div>` : ""}
        ${preview}
      </div>
    </div>
    <div class="ed-actions">
      <label class="check" title="${file.can_remember ? "Похожие прайсы (с такими же названиями столбцов) будут читаться так же" : "Нужны названия столбцов артикула и цены в строке шапки"}">
        <input type="checkbox" id="ed-remember" ${draft.remember && file.can_remember ? "checked" : ""} ${file.can_remember ? "" : "disabled"}> запомнить формат и после сброса этого файла</label>
      <div class="spacer"></div>
      <button class="btn small ghost danger" id="ed-reset">Сбросить на автоопределение</button>
      <button class="btn small" id="ed-cancel">Отмена</button>
      <button class="btn small primary" id="ed-save">Сохранить</button>
    </div>
  </div>`;
}

function openEditor(f, where) {
  S.editWhere = where === "prices" ? "prices" : "report";
  S.editing = {
    id: f.id,
    file: f,
    draft: {
      supplier: f.supplier, city: f.city,
      header_row: f.header_row || (f.head_rows && f.head_rows[0] && f.head_rows[0].row) || 1,
      cols: { ...f.cols }, fixed_brand: f.fixed_brand, remember: true,
    },
  };
  if (S.editWhere === "prices") S.pxOpen[supplierKey(f.supplier) || "—"] = true;
  renderPriceFiles();
}

let dryTimer = null;
function dryRun(resetCols) {
  const ed = S.editing;
  ed.pending = true;
  const head = document.querySelector(".ed-preview-head .muted");
  if (head) head.textContent = "Пересчитываю по первым строкам…";
  clearTimeout(dryTimer);
  dryTimer = setTimeout(async () => {
    const seq = (ed.seq = (ed.seq || 0) + 1);
    try {
      // Новая строка шапки — столбцы ищутся в ней заново; до ответа в списках остаются прежние.
      const body = { folder: S.folder, id: ed.id, ...ed.draft, dry_run: true };
      if (resetCols) delete body.cols;
      const r = await api("/api/prices/save", body);
      if (S.editing !== ed || seq !== ed.seq) return;
      ed.file = r.file;
      if (resetCols) ed.draft.cols = { ...r.file.cols };
      ed.pending = false;
      renderPriceFiles();
    } catch (e) { ed.pending = false; toast(e.message, "bad"); }
  }, 120);
}

async function savePrice(body) {
  const heavy = !body.dry_run && (body.header_row != null || body.cols || body.reset || body.remember);
  if (heavy) showBoot("Сохраняю настройку прайса");
  try {
    const r = await api("/api/prices/save", { folder: S.folder, ...body });
    S.priceFiles = r.files;
    if (S.scan) S.scan.price_files = r.files;
    if (r.message) toast(r.message, r.message.startsWith("Формат запомнен") || r.message.startsWith("Столбцы") ? "good" : "");
    renderDocs();
    renderPriceFiles();
    if (S.editWhere !== "prices") analyze();
    if (heavy) hideBoot();
    return true;
  } catch (e) {
    if (heavy) showBootError(e.message);
    else toast(e.message, "bad");
    return false;
  }
}

function bindPriceFiles() {
  const root = $("#price-files");
  if (!root) return;
  root.querySelectorAll("[data-enable]").forEach((cb) => (cb.onchange = () => savePrice({ id: cb.dataset.enable, enabled: cb.checked })));
  root.querySelectorAll("[data-edit]").forEach((b) => (b.onclick = () => {
    const f = S.priceFiles.find((x) => x.id === b.dataset.edit);
    if (S.editing && S.editWhere !== "prices" && S.editing.id === f.id) { S.editing = null; renderPriceFiles(); } else openEditor(f, "report");
  }));
  if (!S.editing || S.editWhere === "prices") return;
  bindEditor(root);
}

function bindEditor(root) {
  if (!root || !S.editing) return;
  const ed = S.editing;
  root.querySelectorAll("[data-d]").forEach((inp) => {
    const key = inp.dataset.d;
    if (key === "header_row") inp.onchange = () => { ed.draft.header_row = +inp.value; dryRun(true); };
    else if (key === "fixed_brand") inp.onchange = () => { ed.draft.fixed_brand = inp.value.trim(); dryRun(false); };
    else inp.oninput = () => { ed.draft[key] = inp.value; };
  });
  root.querySelectorAll("[data-col]").forEach((sel) => (sel.onchange = () => {
    if (sel.value === "") delete ed.draft.cols[sel.dataset.col]; else ed.draft.cols[sel.dataset.col] = +sel.value;
    dryRun(false);
  }));
  const remember = root.querySelector("#ed-remember");
  if (remember) remember.onchange = (e) => (ed.draft.remember = e.target.checked);
  const cancel = root.querySelector("#ed-cancel");
  if (cancel) cancel.onclick = () => { S.editing = null; renderPriceFiles(); };
  const reset = root.querySelector("#ed-reset");
  if (reset) reset.onclick = async () => {
    if (await savePrice({ id: ed.id, reset: true })) { S.editing = null; renderPriceFiles(); }
  };
  const save = root.querySelector("#ed-save");
  if (save) save.onclick = async () => {
    const d = ed.draft;
    if (!("article" in d.cols) || !("price" in d.cols)) { toast("Укажите столбцы артикула и цены", "bad"); return; }
    const ok = await savePrice({ id: ed.id, supplier: d.supplier, city: d.city, fixed_brand: d.fixed_brand, header_row: d.header_row, cols: d.cols, remember: d.remember && ed.file.can_remember });
    if (ok) { S.editing = null; renderPriceFiles(); }
  };
}

// ------------------------------------------------------------------ прайсы: аналитика и графики

const UP = "#b42318";
const DOWN = "#1d4ed8";
const SAME_C = "#98a2b3";
const PALETTE = ["#0f4c5c", "#d97706", "#7c3aed", "#059669", "#db2777", "#2563eb", "#65a30d", "#dc2626", "#0891b2", "#9333ea", "#ca8a04", "#475569"];
const PRESETS = [["prev", "Предыдущая выгрузка"], ["week", "Неделю назад"], ["month", "Месяц назад"], ["first", "Первая дата"]];

async function loadAnalytics(body = {}) {
  if (S.paLoading) { S.paAgain = body; return; }
  S.paLoading = true;
  const note = $("#pa-note");
  const slow = !S.analytics || (S.scan && (S.scan.price_waiting || []).length);
  if (slow) showBoot("Считаю изменения цен");
  else if (note) note.textContent = "Считаю…";
  try {
    const req = { folder: S.folder };
    const to = body.date_to !== undefined ? body.date_to : (S.analytics && S.analytics.date_to);
    const from = body.date_from !== undefined ? body.date_from : (S.analytics && S.analytics.date_from);
    if (to) req.date_to = to;
    if (from) req.date_from = from;
    const a = await api("/api/prices/analytics", req);
    S.analytics = a;
    if (a.price_files && !S.editing) {
      // Прошлые даты прочитаны — теперь они есть и в «Данных».
      S.priceFiles = a.price_files;
      if (S.scan) { S.scan.price_files = a.price_files; S.scan.price_waiting = []; }
      if (S.tab === "data") renderPriceAudit();
    }
    if (S.pa.brand && !(a.brands || []).some((b) => b.key === S.pa.brand)) { S.pa.brand = ""; S.pa.path = []; }
    paintAnalytics();
    if (slow) hideBoot();
  } catch (e) {
    if (slow) showBootError(e.message);
    else toast(e.message, "bad");
  } finally {
    S.paLoading = false;
    if (S.paAgain) { const next = S.paAgain; S.paAgain = null; loadAnalytics(next); }
  }
}

const signPct = (v, digits = 1) => {
  if (v == null || Number.isNaN(Number(v))) return "—";
  const n = Number(v) * 100;
  const body = Math.abs(n).toLocaleString("ru-RU", { minimumFractionDigits: digits, maximumFractionDigits: digits }) + "%";
  return n > 0.05 ? `+${body}` : n < -0.05 ? `−${body}` : `0%`;
};
const share = (a, b) => (b ? `${Math.round((a / b) * 100)}%` : "—");
const tint = (v) => {
  if (v == null) return "transparent";
  const k = Math.min(1, Math.abs(v) / 0.08);
  if (Math.abs(v) < 0.0005) return "#f2f4f7";
  return v > 0 ? `rgba(180,35,24,${0.12 + 0.6 * k})` : `rgba(29,78,216,${0.12 + 0.6 * k})`;
};

function paintAnalytics() {
  const a = S.analytics;
  if (!a) return;
  const note = $("#pa-note");
  const grid = $("#pa-grid");
  const dates = a.dates || [];
  const fill = (el, list, value) => {
    if (!el) return;
    el.innerHTML = list.map((d) => `<option value="${esc(d)}">${esc(d)}</option>`).join("");
    el.value = value || "";
  };
  fill($("#pa-to"), dates, a.date_to);
  fill($("#pa-from"), dates.filter((d) => dateKey(d) < dateKey(a.date_to)), a.date_from);
  const preset = $("#pa-preset");
  if (preset) {
    const pr = a.presets || {};
    const current = PRESETS.find(([k]) => pr[k] && pr[k] === a.date_from);
    preset.innerHTML = PRESETS.map(([k, t]) => `<button type="button" data-v="${k}" ${pr[k] ? "" : "disabled"} title="${pr[k] ? esc(pr[k]) : "дат пока мало"}">${esc(t)}</button>`).join("");
    preset.querySelectorAll("button").forEach((b) => {
      b.classList.toggle("on", !!current && current[0] === b.dataset.v);
      b.onclick = () => { if (pr[b.dataset.v]) loadAnalytics({ date_from: pr[b.dataset.v] }); };
    });
  }
  const brandSel = $("#pa-brand");
  if (brandSel) {
    const list = (a.brands || []).slice().sort((x, y) => x.brand.localeCompare(y.brand, "ru"));
    brandSel.innerHTML = `<option value="">Все бренды</option>` + list.map((b) => `<option value="${esc(b.key)}">${esc(b.brand)} · ${fmtInt(b.n)}</option>`).join("");
    brandSel.value = S.pa.brand || "";
  }
  const problems = [];
  if (!a.has_crosses) problems.push("Нет файла кроссов: аналитика берёт только позиции из наших кроссов. Положите «Кроссы свод» в Данные/Кроссы.");
  if (dates.length < 2) problems.push("Нужны хотя бы две даты выгрузки: Данные/Прайсы/12.10.2026/Поставщик и Данные/Прайсы/13.10.2026/Поставщик.");
  if (note) note.textContent = problems.length ? problems.join(" ")
    : `Сравнение ${a.date_from} → ${a.date_to}: только позиции, которые есть в обе даты у того же поставщика и города. Пар «бренд + артикул» в кроссах: ${fmtInt(a.crossed)}.`;
  if (grid) grid.classList.toggle("hidden", dates.length < 2 || !a.has_crosses);
  paintKpis(a);
  if (dates.length < 2 || !a.has_crosses) return;
  paintIndex(a);
  paintBrandBars(a);
  paintDrill(a);
  paintWhere(a);
  paintHist(a);
  paintFeed(a);
  paintArticles(a);
  paintCorr(a);
}

function paintKpis(a) {
  const el = $("#pa-kpis");
  if (!el) return;
  const s = a.summary || {};
  const rate = Object.entries((a.rates && a.rates.change) || {}).filter(([, v]) => v != null);
  el.innerHTML = [
    ["Сравнено позиций", fmtInt(s.n), `${fmtInt(s.places || 0)} пар поставщик + город · ${fmtInt(s.brands || 0)} брендов`],
    ["Подорожали", `<span style="color:${UP}">${fmtInt(s.up)}</span>`, share(s.up, s.n)],
    ["Подешевели", `<span style="color:${DOWN}">${fmtInt(s.down)}</span>`, share(s.down, s.n)],
    ["Без изменений", fmtInt(s.same), share(s.same, s.n)],
    ["Медиана изменения", signPct(s.median), `среднее ${signPct(s.mean)}`],
    ["Пропали / новые", `${fmtInt(s.gone)} / ${fmtInt(s.new)}`, "у тех же поставщика и города"],
    ["Курс", rate.length ? rate.map(([c, v]) => `${esc(c)} ${signPct(v)}`).join("<br>") : "—", rate.length ? "за тот же период" : "файла «валюта» нет"],
  ].map(([l, v, sub]) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${esc(sub)}</div></div>`).join("");
}

function lineChart(labels, series, opts = {}) {
  const W = 900;
  const H = opts.height || 280;
  const pad = { l: 46, r: 46, t: 14, b: 30 };
  const shown = series.filter((x) => !x.hidden && x.values.some((v) => v != null));
  const left = shown.filter((x) => !x.right);
  const right = shown.filter((x) => x.right);
  const range = (list) => {
    const vals = list.flatMap((x) => x.values.filter((v) => v != null));
    if (!vals.length) return [99, 101];
    let lo = Math.min(...vals);
    let hi = Math.max(...vals);
    if (hi - lo < 1) { lo -= 0.5; hi += 0.5; }
    const m = (hi - lo) * 0.08;
    return [lo - m, hi + m];
  };
  const [l0, l1] = range(left);
  const [r0, r1] = range(right);
  const n = labels.length;
  const x = (i) => pad.l + (n <= 1 ? 0 : (i * (W - pad.l - pad.r)) / (n - 1));
  const y = (v, lo, hi) => pad.t + (1 - (v - lo) / (hi - lo)) * (H - pad.t - pad.b);
  const ticks = (lo, hi) => {
    const out = [];
    const step = (hi - lo) / 4;
    for (let k = 0; k <= 4; k += 1) out.push(lo + step * k);
    return out;
  };
  const grid = ticks(l0, l1).map((v) => `<line x1="${pad.l}" x2="${W - pad.r}" y1="${y(v, l0, l1)}" y2="${y(v, l0, l1)}" class="ch-grid"/>
    <text x="${pad.l - 6}" y="${y(v, l0, l1) + 4}" class="ch-tick" text-anchor="end">${v.toLocaleString("ru-RU", { maximumFractionDigits: 1 })}</text>`).join("");
  const rticks = right.length ? ticks(r0, r1).map((v) => `<text x="${W - pad.r + 6}" y="${y(v, r0, r1) + 4}" class="ch-tick">${v.toLocaleString("ru-RU", { maximumFractionDigits: 1 })}</text>`).join("") : "";
  const every = Math.max(1, Math.ceil(n / 12));
  const xl = labels.map((t, i) => (i % every === 0 || i === n - 1) ? `<text x="${x(i)}" y="${H - 8}" class="ch-tick" text-anchor="middle">${esc(String(t).slice(0, 5))}</text>` : "").join("");
  const base = left.length && l0 < 100 && l1 > 100 && opts.base100 ? `<line x1="${pad.l}" x2="${W - pad.r}" y1="${y(100, l0, l1)}" y2="${y(100, l0, l1)}" class="ch-base"/>` : "";
  const lines = shown.map((sr) => {
    const [lo, hi] = sr.right ? [r0, r1] : [l0, l1];
    let d = "";
    let pen = false;
    sr.values.forEach((v, i) => {
      if (v == null) { pen = false; return; }
      d += `${pen ? "L" : "M"}${x(i).toFixed(1)},${y(v, lo, hi).toFixed(1)}`;
      pen = true;
    });
    const dots = sr.values.map((v, i) => v == null ? "" : `<circle cx="${x(i)}" cy="${y(v, lo, hi)}" r="${n > 30 ? 2 : 3.2}" fill="${sr.color}"><title>${esc(sr.name)} · ${esc(labels[i])}: ${v.toLocaleString("ru-RU", { maximumFractionDigits: 2 })}${sr.unit || ""}</title></circle>`).join("");
    return `<path d="${d}" fill="none" stroke="${sr.color}" stroke-width="${sr.width || 2}" ${sr.dash ? `stroke-dasharray="6 4"` : ""}/>${dots}`;
  }).join("");
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img">${grid}${rticks}${base}${xl}${lines}</svg>`;
}

function paintIndex(a) {
  const el = $("#pa-index");
  if (!el) return;
  const idx = a.index || {};
  const lines = idx.lines || {};
  const keys = Object.keys(lines).filter((k) => k);
  keys.sort((p, q) => (idx.labels[p] || p).localeCompare(idx.labels[q] || q, "ru"));
  const focus = S.pa.brand;
  const series = [{ key: "", name: "Все бренды", values: lines[""] || [], color: "#17212b", width: 3 }];
  keys.forEach((k, i) => series.push({ key: k, name: idx.labels[k] || k, values: lines[k], color: PALETTE[i % PALETTE.length], width: focus === k ? 3.5 : 1.8 }));
  const rates = (a.rates && a.rates.index) || {};
  Object.keys(rates).forEach((code, i) => series.push({ key: `rate:${code}`, name: `Курс ${code}`, values: rates[code], color: ["#94a3b8", "#a16207", "#be185d", "#0e7490"][i % 4], dash: true, width: 1.6 }));
  series.forEach((sr) => {
    sr.hidden = S.pa.hidden[sr.key] === true || (focus && sr.key && !sr.key.startsWith("rate:") && sr.key !== focus && S.pa.hidden[sr.key] !== false);
  });
  const legend = series.map((sr) => `<button type="button" class="lg ${sr.hidden ? "off" : ""}" data-lg="${esc(sr.key)}"><i style="background:${sr.color}${sr.dash ? ";height:2px" : ""}"></i>${esc(sr.name)}</button>`).join("");
  el.innerHTML = `${lineChart(idx.dates || [], series, { base100: true })}<div class="ch-legend">${legend}</div>
    <div class="muted small-text">Нажмите на бренд в легенде — линия скроется или появится. Индекс: от даты к дате медиана изменения по позициям, которые есть в обе даты, затем перемножение.</div>`;
  el.querySelectorAll("[data-lg]").forEach((b) => (b.onclick = () => {
    const k = b.dataset.lg;
    const sr = series.find((x) => x.key === k);
    S.pa.hidden[k] = !(sr && sr.hidden) ? true : false;
    paintIndex(a);
  }));
}

function paintBrandBars(a) {
  const el = $("#pa-brands");
  if (!el) return;
  const rows = (a.brands || []).slice();
  if (!rows.length) { el.innerHTML = `<div class="muted">Общих позиций в эти даты нет.</div>`; return; }
  el.innerHTML = `<div class="dv-head"><span style="color:${DOWN}">← подешевели</span><span style="color:${UP}">подорожали →</span></div>` + rows.map((r) => {
    const down = r.n ? (r.down / r.n) * 50 : 0;
    const up = r.n ? (r.up / r.n) * 50 : 0;
    return `<button type="button" class="dv-row ${S.pa.brand === r.key ? "on" : ""}" data-brand="${esc(r.key)}" title="${esc(`${r.brand}: вверх ${r.up}, вниз ${r.down}, без изменений ${r.same} из ${r.n}. Пропало ${r.gone}, новых ${r.new}`)}">
      <span class="dv-name">${esc(r.brand)}</span>
      <span class="dv-bar"><i class="dn" style="width:${down}%;right:50%"></i><i class="up" style="width:${up}%;left:50%"></i><em></em></span>
      <span class="dv-num"><b style="color:${(r.median || 0) > 0.0005 ? UP : (r.median || 0) < -0.0005 ? DOWN : "inherit"}">${signPct(r.median)}</b> <span class="muted">${fmtInt(r.n)} поз.</span></span>
    </button>`;
  }).join("");
  el.querySelectorAll("[data-brand]").forEach((b) => (b.onclick = () => {
    S.pa.brand = S.pa.brand === b.dataset.brand ? "" : b.dataset.brand;
    S.pa.path = [];
    paintAnalytics();
  }));
}

function placeRows(a, brand, path) {
  const rows = (a.places || []).filter((p) => p.key === brand);
  const [cat, group] = path;
  if (!cat) return rows.filter((p) => !p.group);
  if (!group) return rows.filter((p) => p.category === cat && p.group && !p.subgroup);
  return rows.filter((p) => p.category === cat && p.group === group && p.subgroup);
}

function paintDrill(a) {
  const el = $("#pa-drill");
  const sub = $("#pa-drill-sub");
  if (!el) return;
  const brand = S.pa.brand || ((a.brands || [])[0] || {}).key;
  if (!brand) { el.innerHTML = `<div class="muted">Нет данных.</div>`; return; }
  const name = ((a.brands || []).find((b) => b.key === brand) || {}).brand || brand;
  const path = S.pa.path || [];
  const rows = placeRows(a, brand, path).sort((x, y) => Math.abs(y.median || 0) * Math.sqrt(y.n) - Math.abs(x.median || 0) * Math.sqrt(x.n));
  if (sub) sub.textContent = S.pa.brand ? name : `${name} — больше всего изменений; выберите бренд слева`;
  const crumbs = [`<button type="button" class="crumb" data-crumb="0">${esc(name)}</button>`,
    ...path.map((p, i) => `<button type="button" class="crumb" data-crumb="${i + 1}">${esc(p)}</button>`)].join(" › ");
  const max = Math.max(0.01, ...rows.map((r) => Math.abs(r.median || 0)));
  const deeper = path.length < 2;
  const body = rows.length ? rows.map((r) => {
    const label = path.length === 0 ? r.category : path.length === 1 ? r.group : r.subgroup;
    const w = (Math.abs(r.median || 0) / max) * 48;
    const side = (r.median || 0) >= 0 ? `left:50%;width:${w}%;background:${UP}` : `right:50%;width:${w}%;background:${DOWN}`;
    return `<button type="button" class="dr-row" data-place="${esc(label)}" ${deeper ? "" : "data-leaf=1"}>
      <span class="dr-name">${esc(label || "—")}</span>
      <span class="dr-bar"><i style="${side}"></i><em></em></span>
      <span class="dr-num"><b>${signPct(r.median)}</b> <span class="muted">↑${fmtInt(r.up)} ↓${fmtInt(r.down)} из ${fmtInt(r.n)}</span></span>
    </button>`;
  }).join("") : `<div class="muted">Позиции этого бренда не привязаны к нашим группам: нет кроссов «Да» к товарам ПСТ.</div>`;
  el.innerHTML = `<div class="crumbs">${crumbs}</div>${body}`;
  el.querySelectorAll("[data-crumb]").forEach((b) => (b.onclick = () => {
    S.pa.brand = brand;
    S.pa.path = path.slice(0, Number(b.dataset.crumb));
    paintDrill(a); paintArticles(a);
  }));
  el.querySelectorAll("[data-place]").forEach((b) => (b.onclick = () => {
    S.pa.brand = brand;
    if (b.dataset.leaf) S.pa.path = [...path.slice(0, 2), b.dataset.place];
    else S.pa.path = [...path, b.dataset.place];
    if (S.pa.path.length > 2 && !b.dataset.leaf) S.pa.path = S.pa.path.slice(0, 2);
    paintDrill(a); paintArticles(a);
  }));
}

function paintWhere(a) {
  const el = $("#pa-where");
  if (!el) return;
  const brand = S.pa.brand || "";
  const rows = (a.where || []).filter((w) => w.key === brand);
  if (!rows.length) { el.innerHTML = `<div class="muted">Нет общих позиций.</div>`; return; }
  const cities = [...new Set(rows.map((r) => r.city))].sort((x, y) => x.localeCompare(y, "ru"));
  const sups = [...new Set(rows.map((r) => r.supplier))].sort((x, y) => x.localeCompare(y, "ru"));
  const at = new Map(rows.map((r) => [`${r.supplier}|${r.city}`, r]));
  const head = `<tr><th></th>${cities.map((c) => `<th class="hm-city">${esc(c === "—" ? "город не указан" : c)}</th>`).join("")}</tr>`;
  const body = sups.map((sup) => `<tr><td class="nowrap"><b>${esc(sup)}</b></td>${cities.map((c) => {
    const r = at.get(`${sup}|${c}`);
    if (!r) return `<td class="hm-cell none"></td>`;
    const dark = Math.abs(r.median || 0) > 0.05;
    return `<td class="hm-cell" style="background:${tint(r.median)};${dark ? "color:#fff" : ""}" title="${esc(`${sup} · ${c}: медиана ${signPct(r.median)}, вверх ${r.up}, вниз ${r.down}, без изменений ${r.same} из ${r.n}`)}">${signPct(r.median)}<small>${fmtInt(r.n)}</small></td>`;
  }).join("")}</tr>`).join("");
  el.innerHTML = `<div class="table-wrap"><table class="heat">${head}${body}</table></div>
    <div class="muted small-text">${brand ? esc(((a.brands || []).find((b) => b.key === brand) || {}).brand || brand) : "Все бренды"}. Цвет — медиана: красный — подорожало, синий — подешевело. Мелкая цифра — позиций.</div>`;
}

function paintHist(a) {
  const el = $("#pa-hist");
  if (!el) return;
  const labels = (a.hist && a.hist.labels) || [];
  const values = ((a.hist && a.hist.values) || {})[S.pa.brand || ""] || labels.map(() => 0);
  const W = 520;
  const H = 220;
  const max = Math.max(1, ...values);
  const bw = (W - 20) / labels.length;
  const mid = Math.floor(labels.length / 2);
  const bars = values.map((v, i) => {
    const h = (v / max) * (H - 60);
    const color = i < mid ? DOWN : i > mid ? UP : SAME_C;
    const xx = 10 + i * bw;
    return `<rect x="${xx + 4}" y="${H - 40 - h}" width="${bw - 8}" height="${h}" fill="${color}" opacity="${i === mid ? 0.6 : 0.85}"><title>${esc(labels[i])}: ${v}</title></rect>
      <text x="${xx + bw / 2}" y="${H - 44 - h}" class="ch-tick" text-anchor="middle">${v ? fmtInt(v) : ""}</text>
      <text x="${xx + bw / 2}" y="${H - 22}" class="ch-tick" text-anchor="middle">${esc(labels[i].replace("без изменений", "0"))}</text>`;
  }).join("");
  el.innerHTML = `<svg class="chart" viewBox="0 0 ${W} ${H}">${bars}</svg>`;
}

function paintFeed(a) {
  const el = $("#pa-feed");
  if (!el) return;
  const rows = (a.feed || []).filter((r) => !S.pa.brand || r.key === S.pa.brand).slice(0, 24);
  el.innerHTML = rows.length ? `<div class="feed">${rows.map((r) => {
    const dir = (r.median || 0) > 0.0005 ? "up" : (r.median || 0) < -0.0005 ? "down" : "mix";
    const word = dir === "up" ? "подняли" : dir === "down" ? "снизили" : "двигали в обе стороны";
    return `<div class="feed-row ${dir}"><b>${esc(r.brand)}</b> · ${esc(r.supplier)}${r.city && r.city !== "—" ? ` · ${esc(r.city)}` : ""} — ${word}: ${fmtInt(r.up + r.down)} из ${fmtInt(r.n)} позиций,
      медиана ${signPct(r.median)}, среднее ${signPct(r.mean)}</div>`;
  }).join("")}</div>` : `<div class="muted">Заметных сдвигов нет: меньше 10% позиций поменяли цену.</div>`;
}

function paintArticles(a) {
  const el = $("#pa-articles");
  const sub = $("#pa-articles-sub");
  if (!el) return;
  const brand = S.pa.brand || ((a.brands || [])[0] || {}).key;
  const [cat, group, subgroup] = S.pa.path || [];
  let rows = ((a.articles || {})[brand] || []).filter((r) => (!cat || (r.category || "Без категории") === cat) && (!group || r.group === group) && (!subgroup || r.subgroup === subgroup));
  const name = ((a.brands || []).find((b) => b.key === brand) || {}).brand || brand || "";
  if (sub) sub.textContent = `${name}${cat ? ` · ${[cat, group, subgroup].filter(Boolean).join(" › ")}` : ""} · самые большие изменения`;
  rows = rows.slice(0, 60);
  el.innerHTML = rows.length ? `<div class="table-wrap scroll"><table class="mini"><thead><tr><th>Артикул</th><th>Поставщик</th><th>Город</th><th class="num">Было</th><th class="num">Стало</th><th class="num">Изм.</th><th>Подгруппа</th></tr></thead><tbody>
    ${rows.map((r) => `<tr><td class="nowrap"><b>${esc(r.article)}</b></td><td>${esc(r.supplier)}</td><td>${esc(r.city === "—" ? "" : r.city)}</td>
      <td class="num">${fmtMoney(r.old)}</td><td class="num">${fmtMoney(r.new)}</td>
      <td class="num" style="color:${r.change > 0 ? UP : DOWN}"><b>${signPct(r.change)}</b></td><td class="muted">${esc(r.subgroup || r.group || "")}</td></tr>`).join("")}
  </tbody></table></div>` : `<div class="muted">Изменений цен нет.</div>`;
}

function paintCorr(a) {
  const el = $("#pa-corr");
  if (!el) return;
  const rows = (a.corr || []).filter((r) => !S.pa.brand || r.key === S.pa.brand);
  const codes = (a.rates && a.rates.codes) || [];
  if (!codes.length) { el.innerHTML = `<div class="muted">Файла «валюта» нет: положите его прямо в Данные/Прайсы — в шапке даты, в строках USD, EUR, JPY, CNY.</div>`; return; }
  if (a.corr_note || !rows.length) {
    el.innerHTML = `<div class="muted">${esc(a.corr_note || "Мало общих шагов цены и курса.")} Пока видно только направление: сравните пунктир курса с линиями брендов на графике индекса.</div>`;
    return;
  }
  const word = (r) => Math.abs(r) >= 0.7 ? "сильная" : Math.abs(r) >= 0.4 ? "заметная" : "слабая";
  el.innerHTML = `<div class="table-wrap"><table class="mini"><thead><tr><th>Бренд</th><th>Валюта</th><th class="num">Связь</th><th>Как читать</th></tr></thead><tbody>
    ${rows.sort((x, y) => Math.abs(y.r) - Math.abs(x.r)).slice(0, 40).map((r) => `<tr><td><b>${esc(r.brand)}</b></td><td>${esc(r.code)}</td>
      <td class="num">${r.r.toLocaleString("ru-RU", { maximumFractionDigits: 2 })}</td>
      <td class="muted">${word(r.r)} ${r.r > 0 ? "прямая" : "обратная"} связь${r.lag ? `, цена идёт за курсом через ${r.lag} ${plural(r.lag, "дату", "даты", "дат")}` : ", в ту же дату"} · шагов ${r.n}</td></tr>`).join("")}
  </tbody></table></div>`;
}

async function writePriceExport() {
  const btn = $("#pa-export");
  if (btn) btn.disabled = true;
  showBoot("Пишу выгрузку цен");
  try {
    const r = await api("/api/prices/export", { folder: S.folder });
    S.pxFolder = r.folder;
    hideBoot();
    toast(`Записан ${r.name}: ${fmtInt(r.rows)} строк, ${fmtInt(r.dates)} ${plural(r.dates, "дата", "даты", "дат")}, ${fmtInt(r.files)} файлов`, "good");
    api("/api/open", { path: r.path }).catch(() => {});
  } catch (e) {
    showBootError(e.message);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function writeCrossPrices() {
  const btn = $("#px-crosses");
  if (btn) btn.disabled = true;
  try {
    const r = await api("/api/prices/crosses", { folder: S.folder });
    S.pxFolder = r.folder;
    toast(`Записан ${r.name}: ${fmtInt(r.matched)} цен, кроссов «Да» без цены ${fmtInt(r.missing)}.`, "good");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function writePriceLayout() {
  const btn = $("#px-layout");
  if (btn) btn.disabled = true;
  try {
    const r = await api("/api/prices/layout", { folder: S.folder });
    S.pxFolder = r.folder;
    toast(`Записан ${r.name}. Его можно прислать целиком.`, "good");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ------------------------------------------------------------------ курсы

function allMissingRates() {
  const known = new Set();
  for (const [k, v] of (S.settings && S.settings.rates) || []) {
    if (String(k || "").trim() && v !== "" && v != null) known.add(String(k).trim().toUpperCase());
  }
  const wide = (S.scan && S.scan.missing_rates) || [];
  const one = (S.analysis && S.analysis.missing_rates) || [];
  const raw = wide.length ? wide : one;
  return raw.filter((m) => m && m.currency && !known.has(String(m.currency).trim().toUpperCase()));
}

function renderRatesNote() {
  const missing = allMissingRates();
  $("#rates-note").classList.toggle("hidden", !missing.length);
  document.querySelector(".rates-k").classList.toggle("hidden", !missing.length);
  if (!missing.length) return;
  const list = missing.map((m) => `<b>${esc(m.currency)}</b> <span class="muted">(${fmtInt(m.items)} тов.)</span>`).join(", ");
  $("#rates-note").innerHTML = `<div class="note">Нет курса для ${list}. Это по всем группам, не только по открытой. Без курса не посчитаются закупка в рублях, СБС план и цена смежной марки.
    <button class="btn small ghost" id="goto-rates">Задать курсы</button></div>`;
  $("#goto-rates").onclick = () => showTab("settings");
}

// ------------------------------------------------------------------ шаг 5: источники

function renderSources() {
  const a = S.analysis;
  const choices = S.init.sources;
  $("#sources").innerHTML = a.compare.map((c) => {
    const parts = [];
    if (c.both) parts.push(`в обоих ${fmtInt(c.both)}, расходится <b>${fmtInt(c.differ)}</b>`);
    if (c.only_own) parts.push(`только у нас ${fmtInt(c.only_own)}`);
    if (c.only_acp) parts.push(`только в АЦПКД ${fmtInt(c.only_acp)}`);
    return `<div class="src-row">
      <div><div class="t">${esc(c.title)}</div></div>
      <div>${segHtml(choices, c.source, `data-src="${c.key}"`)}<div class="opts">Наши: ${esc(c.own)} · АЦПКД: ${esc(c.acp)}</div></div>
      <div class="stat">${parts.length ? parts.join(" · ") : '<span class="muted">нет данных ни там, ни там</span>'}</div>
    </div>`;
  }).join("");
  $("#sources").querySelectorAll("[data-src]").forEach((el) => {
    const key = el.dataset.src;
    const cur = a.compare.find((c) => c.key === key).source;
    seg(el, cur, (v) => { S.opts.sources[key] = v; analyze(); });
  });
}

// ------------------------------------------------------------------ шаг 6: предпросмотр

const KNOBS = [
  { key: "tolerance", title: "Коридор: не трогаем, если до цели меньше", unit: "%", min: 0, max: 10, step: 0.5,
    hint: "Цена уже рядом с целью — оставляем как есть" },
  { key: "down", title: "Не понижать больше чем на", unit: "%", min: 0, max: 30, step: 0.5,
    hint: "Самый большой шаг вниз за один раз" },
  { key: "up", title: "Не повышать больше чем на", unit: "%", min: 0, max: 30, step: 0.5, segments: true,
    hint: "По ценовому сегменту автопарка: премиум терпит рост лучше" },
  { key: "min_vr", title: "Не опускать ВР ниже", unit: "%", min: 0, max: 60, step: 1,
    hint: "Защита маржи: ниже — цену поднимаем, даже если рынок дешевле" },
  { key: "max_lift", title: "Снижать, только если окупится ростом продаж до", unit: "%", min: 0, max: 60, step: 1,
    hint: "Снижение на 8% при ВР 39% окупается, только если продажи вырастут на 26%. Глубже этого порога цену не режем. 0 — не проверять" },
];
const ROUNDING = [1, 5, 10, 50, 100];

function knobHtml(k, value, label = k.title, attr = `data-knob="${k.key}"`) {
  const v = value ?? "";
  return `<div class="knob">
    <div class="knob-head"><span>${esc(label)}</span><span class="knob-val"><input class="input knob-num" type="number" ${attr} min="${k.min}" max="${k.max}" step="${k.step}" value="${esc(v)}"> ${esc(k.unit)}</span></div>
    <input type="range" class="knob-range" ${attr} min="${k.min}" max="${k.max}" step="${k.step}" value="${esc(v === "" ? k.min : v)}">
  </div>`;
}

function renderRecForm() {
  const r = S.opts.rec;
  $("#rec-enabled").checked = !!r.enabled;
  $("#rec-body").classList.toggle("off", !r.enabled);
  const targetButtons = [
    ["first", "Только первый", "Только первый в списке позиционирования. Нет его цены — на остальных не ориентируемся."],
    ["main", "По списку", "Сначала первый в списке. Нет цены — следующий, со своим процентом."],
    ["all", "Все с ценой", "Среднее или самая низкая цель среди всех, у кого есть цена."],
  ];
  const group = r.target === "avg" || r.target === "min" ? "all" : (r.target === "first" ? "first" : "main");
  $("#rec-target").innerHTML = targetButtons.map(([k, t, tip]) =>
    `<button type="button" data-v="${k}" title="${esc(tip)}">${esc(t)}</button>`).join("");
  seg($("#rec-target"), group, (v) => {
    r.target = v === "all" ? (r.target === "min" ? "min" : "avg") : v;
    rememberGroupRec();
    renderRecForm();
    analyze(120);
  });
  const all = $("#rec-all");
  all.classList.toggle("hidden", group !== "all");
  seg(all, r.target === "min" ? "min" : "avg", (v) => { r.target = v; rememberGroupRec(); analyze(120); });
  const note = targetButtons.find(([k]) => k === group);
  $("#rec-target-note").textContent = note ? note[2] : "";
  seg($("#green-when"), S.opts.green_when, (v) => { S.opts.green_when = v; persist(); if (S.analysis) renderDash(); });
  seg($("#money-base"), S.opts.money_base === "potential" ? "potential" : "sales", (v) => {
    S.opts.money_base = v === "potential" ? "potential" : "sales";
    persist();
    analyze(120);
  });
  const segments = S.init.segments || [];
  $("#rec-knobs").innerHTML = KNOBS.map((k) => {
    if (k.segments) {
      return `<div class="rp-sec"><div class="rp-t">${esc(k.title)}</div><div class="muted small-text">${esc(k.hint)}</div>
        ${segments.map((sg) => knobHtml(k, (r.up || {})[sg], sg, `data-up="${esc(sg)}"`)).join("")}</div>`;
    }
    const pair = k.pair ? knobHtml({ key: k.pair, min: 0.1, max: 1, step: 0.05, unit: "× лимиты" }, r[k.pair], "лимиты умножаются на") : "";
    return `<div class="rp-sec">${knobHtml(k, r[k.key])}${pair}<div class="muted small-text">${esc(k.hint)}</div></div>`;
  }).join("") + `<div class="rp-sec"><div class="rp-t">Округление Кат. А</div>
      <div class="seg" id="rec-round">${ROUNDING.map((v) => `<button type="button" data-v="${v}">${v} ₽</button>`).join("")}</div></div>`;
  seg($("#rec-round"), String(Math.round(Number(r.rounding) || 1)), (v) => { r.rounding = Number(v); rememberGroupRec(); analyze(120); });
  const knobs = $("#rec-knobs");
  const apply = (el) => {
    const n = toNum(el.value);
    if (n === "") return;
    if (el.dataset.up) {
      r.up = r.up || {};
      r.up[el.dataset.up] = n;
    } else r[el.dataset.knob] = n;
    // Ползунок и число рядом — одно значение.
    const sel = el.dataset.up ? `[data-up="${CSS.escape(el.dataset.up)}"]` : `[data-knob="${el.dataset.knob}"]`;
    knobs.querySelectorAll(sel).forEach((other) => { if (other !== el) other.value = n; });
    rememberGroupRec();
    analyze(150);
  };
  knobs.querySelectorAll("input").forEach((el) => (el.oninput = () => apply(el)));
  $("#rec-enabled").onchange = (e) => {
    r.enabled = e.target.checked;
    $("#rec-body").classList.toggle("off", !r.enabled);
    renderRecScope();
    persist();
    analyze(120);
  };
  renderRecScope();
}

function tone(label) {
  if (String(label).startsWith("В пределах")) return "mid";
  const dearer = String(label).startsWith("Дороже");
  return dearer === (S.opts.green_when !== "below") ? "good" : "bad";
}

const fmtRub = (v) => (v == null || v === "" || Number.isNaN(Number(v)) ? "—" : `${Math.round(Number(v)).toLocaleString("ru-RU")} ₽`);
const fmtSignedRub = (v) => {
  if (v == null || Number.isNaN(Number(v))) return "—";
  const n = Math.round(Number(v));
  const body = `${Math.abs(n).toLocaleString("ru-RU")} ₽`;
  return n > 0 ? `+${body}` : n < 0 ? `−${body}` : body;
};

function deltaKind(now, after, better) {
  if (now == null || after == null || Number.isNaN(Number(now)) || Number.isNaN(Number(after))) return "";
  const d = Number(after) - Number(now);
  if (Math.abs(d) < 1e-9) return "flat";
  if (better === "closer") return Math.abs(Number(after)) < Math.abs(Number(now)) - 1e-9 ? "good" : "bad";
  if (better === "down") return d < 0 ? "good" : "bad";
  return d > 0 ? "good" : "bad";
}

function cmpRow(label, now, after, delta, kind) {
  return `<tr><td>${esc(label)}</td><td class="num now">${now}</td><td class="num after">${after}</td><td class="num delta ${kind}">${delta}</td></tr>`;
}

function cutTable(title, rows, head, body) {
  if (!rows.length) return "";
  return `<div class="dash-block"><div class="dash-title">${esc(title)}</div>
    <div class="table-wrap scroll"><table class="mini"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div></div>`;
}

function groupedBars(labels, a, b, opts = {}) {
  const W = 400;
  const H = 230;
  const max = Math.max(1, ...a, ...(b || []));
  const n = labels.length;
  const slot = (W - 20) / n;
  const bw = b ? (slot - 12) / 2 : slot - 14;
  const out = labels.map((label, i) => {
    const x = 10 + i * slot;
    const h1 = (a[i] / max) * (H - 70);
    const bar1 = `<rect x="${x + 6}" y="${H - 44 - h1}" width="${bw}" height="${h1}" rx="3" fill="${(opts.colorA && opts.colorA(i)) || "#98a2b3"}"><title>${esc(label)}: ${fmtInt(a[i])}${opts.unit || ""}</title></rect>
      <text x="${x + 6 + bw / 2}" y="${H - 48 - h1}" class="ch-tick" text-anchor="middle">${a[i] ? fmtInt(a[i]) : ""}</text>`;
    let bar2 = "";
    if (b) {
      const h2 = (b[i] / max) * (H - 70);
      bar2 = `<rect x="${x + 8 + bw}" y="${H - 44 - h2}" width="${bw}" height="${h2}" rx="3" fill="#0f4c5c"><title>${esc(label)}, после: ${fmtInt(b[i])}</title></rect>
        <text x="${x + 8 + bw + bw / 2}" y="${H - 48 - h2}" class="ch-tick" text-anchor="middle" style="fill:#0f4c5c">${b[i] ? fmtInt(b[i]) : ""}</text>`;
    }
    const text = [];
    String(label).split(" ").forEach((w) => {
      const last = text[text.length - 1];
      if (last != null && (last + " " + w).length <= 11) text[text.length - 1] = `${last} ${w}`;
      else text.push(w);
    });
    const lab = text.slice(0, 3).map((t, k) => `<text x="${x + slot / 2}" y="${H - 30 + k * 11}" class="ch-tick ch-small" text-anchor="middle">${esc(t)}</text>`).join("");
    return bar1 + bar2 + lab;
  }).join("");
  return `<svg class="chart" viewBox="0 0 ${W} ${H}">${out}</svg>`;
}

function hbars(rows, total) {
  const max = Math.max(1, ...rows.map((r) => r.value));
  return `<div class="hb">${rows.map((r) => `<div class="hb-row" title="${esc(r.tip || "")}">
    <span class="hb-name">${esc(r.label)}</span>
    <span class="hb-track"><i style="width:${(r.value / max) * 100}%;background:${r.color || "#0f4c5c"}"></i></span>
    <span class="hb-num"><b>${fmtInt(r.value)}</b>${total ? ` <span class="muted">${pct(r.value, total)}%</span>` : ""}</span>
  </div>`).join("")}</div>`;
}

// Где наша цена относительно цели: расходящаяся шкала, «нет цены» — отдельно, серым со штриховкой.
const POS_COLORS = { no_price: "url(#hatch)", d20: "#b3541e", d10: "#e08a4a", d: "#f2c49b", band: "#c9d1d9", c: "#a9cbe8", c10: "#5b9bd5", c20: "#1f5f99" };
const POS_CSS = { no_price: "repeating-linear-gradient(135deg,#e5e9ee 0 4px,#f5f7f9 4px 8px)", d20: "#b3541e", d10: "#e08a4a", d: "#f2c49b", band: "#c9d1d9", c: "#a9cbe8", c10: "#5b9bd5", c20: "#1f5f99" };

function positionMix(d) {
  const pos = d.position;
  if (!pos || !pos.now) return "";
  const byRub = d.weight_by !== "items";
  const line = (title, parts) => {
    const total = parts.reduce((s, p) => s + (byRub ? p.sale : p.items), 0) || 1;
    const segs = parts.filter((p) => (byRub ? p.sale : p.items) > 0).map((p) => {
      const share = (byRub ? p.sale : p.items) / total;
      const dark = ["d20", "c20", "c10"].includes(p.key);
      return `<i style="flex:${share};background:${POS_CSS[p.key]};${dark ? "color:#fff" : ""}" title="${esc(`${p.label}: ${Math.round(share * 100)}% выручки · ${fmtInt(p.items)} тов. · ${fmtRub(p.sale)}`)}">${share >= 0.06 ? `${Math.round(share * 100)}%` : ""}</i>`;
    }).join("");
    return `<div class="mix-row"><span class="mix-t">${esc(title)}</span><div class="mix-bar">${segs}</div></div>`;
  };
  const legend = pos.now.map((p) => `<span><i style="background:${POS_CSS[p.key]}"></i>${esc(p.label)}</span>`).join("");
  const total = pos.now.reduce((s, p) => s + p.sale, 0) || 1;
  const none = pos.now.find((p) => p.key === "no_price");
  const dear = pos.now.filter((p) => p.key.startsWith("d")).reduce((s, p) => s + p.sale, 0);
  const cheap = pos.now.filter((p) => p.key.startsWith("c")).reduce((s, p) => s + p.sale, 0);
  const band = (pos.now.find((p) => p.key === "band") || {}).sale || 0;
  const priced = total - (none ? none.sale : 0);
  const say = priced
    ? `Из выручки с ценой конкурента: дороже цели — ${Math.round(dear / priced * 100)}%, в коридоре — ${Math.round(band / priced * 100)}%, дешевле — ${Math.round(cheap / priced * 100)}%. Без цены конкурента — ${Math.round((none ? none.sale : 0) / total * 100)}% выручки: по ним рекомендация видит только ВР.`
    : "Цены конкурента нет ни у одного товара — рекомендация видит только ВР.";
  return `<div class="card flat wide"><div class="dash-title">Где наша цена относительно цели <span class="muted">· доля выручки группы</span></div>
    ${line("Сейчас", pos.now)}${pos.after ? line("После", pos.after) : ""}
    <div class="mix-legend">${legend}</div><div class="muted small-text">${esc(say)}</div></div>`;
}

function bridgeBlock(d) {
  const b = d.bridge;
  if (!b) return "";
  const up = b.up || {};
  const dn = b.down || {};
  const net = (up.delta || 0) + (dn.delta || 0);
  const max = Math.max(1, Math.abs(up.delta || 0), Math.abs(dn.delta || 0), Math.abs(net));
  const bar = (label, v, color, sub) => `<div class="br-row"><span class="br-t">${esc(label)}<small>${esc(sub || "")}</small></span>
      <span class="br-track"><em></em><i style="${v >= 0 ? `left:50%` : `right:50%`};width:${Math.abs(v) / max * 50}%;background:${color}"></i></span>
      <b class="br-v ${v > 0 ? "pos" : v < 0 ? "neg" : ""}">${fmtSignedRub(v)}</b></div>`;
  const pct = b.profit ? ` (${fmtPct(net / b.profit, true)} прибыли)` : "";
  const lines = [];
  if (up.items) {
    lines.push(`<div class="pb up"><b>↑ ${fmtInt(up.items)} ${plural(up.items, "товар", "товара", "товаров")}</b>, медиана ${fmtPct(up.median_change, true)}.
      ${up.may_lose != null ? `Прибыль не упадёт, даже если их продажи снизятся на <b>${fmtPct(up.may_lose)}</b>.` : ""}</div>`);
  }
  if (dn.items) {
    lines.push(`<div class="pb down"><b>↓ ${fmtInt(dn.items)} ${plural(dn.items, "товар", "товара", "товаров")}</b>, медиана ${fmtPct(dn.median_change, true)}.
      ${dn.need_lift != null ? `Окупится, только если их продажи вырастут на <b>${fmtPct(dn.need_lift)}</b> — иначе это минус прибыли.` : ""}</div>`);
  }
  return `<div class="card flat"><div class="dash-title">Откуда меняется прибыль <span class="muted">· при тех же продажах</span></div>
    ${bar("Повышения", up.delta || 0, "#1f5f99", up.items ? `${fmtInt(up.items)} тов.` : "")}
    ${bar("Снижения", dn.delta || 0, "#b3541e", dn.items ? `${fmtInt(dn.items)} тов.` : "")}
    ${bar("Итого", net, "#17212b", pct)}
    <div class="pb-list">${lines.join("") || `<div class="muted">Рекомендация ничего не меняет.</div>`}</div></div>`;
}

function compBars(d, recOn) {
  const comps = (d.slices || []).filter((s) => s.kind === "comp" && s.n);
  if (!comps.length) return "";
  const val = (s, key) => (s[key] != null ? s[key] : s.median);
  const vals = comps.flatMap((s) => [val(s, "weighted"), s.weighted_rec, s.target_dev]).filter((v) => v != null);
  const span = Math.min(0.6, Math.max(0.1, ...vals.map((v) => Math.abs(v)))) * 1.1;
  const x = (v) => 50 + Math.max(-50, Math.min(50, (v / span) * 50));
  const rows = comps.map((s) => {
    const now = val(s, "weighted");
    const after = s.weighted_rec;
    const name = String(s.title).replace(" (без учёта % позиц.)", "").replace("Против ", "");
    const seg = (v, cls) => v == null ? "" : `<i class="${cls}" style="${v >= 0 ? `left:50%;width:${x(v) - 50}%` : `left:${x(v)}%;width:${50 - x(v)}%`}"></i>`;
    const goal = s.target_dev != null ? `<b class="goal" style="left:${x(s.target_dev)}%" title="Цель по позиционированию: ${esc(fmtPct(s.target_dev, true))}"></b>` : "";
    return `<div class="cp-row" title="${esc(`${name}: ${fmtInt(s.n)} тов. с ценой. Сейчас ${fmtPct(now, true)}${after != null ? `, после ${fmtPct(after, true)}` : ""}${s.target_dev != null ? `, цель ${fmtPct(s.target_dev, true)}` : ""}. Плюс — мы дешевле`)}">
      <span class="cp-name">${esc(name)}<small>${fmtInt(s.n)} тов.</small></span>
      <span class="cp-track"><em></em>${seg(now, "now")}${recOn ? seg(after, "after") : ""}${goal}</span>
      <span class="cp-num">${fmtPct(now, true)}${recOn && after != null ? ` <span class="to">→ ${fmtPct(after, true)}</span>` : ""}</span></div>`;
  }).join("");
  return `<div class="card flat${recOn ? "" : " wide"}"><div class="dash-title">Мы против каждого конкурента <span class="muted">· средняя по выручке разница в цене</span></div>
    <div class="cp-axis"><span>← мы дороже</span><span>вровень</span><span>дешевле →</span></div>${rows}
    <div class="mix-legend"><span><i style="background:#9fb3c8"></i>сейчас</span>${recOn ? `<span><i style="background:#0f4c5c"></i>после рекомендаций</span>` : ""}<span><i class="goal-i"></i>цель по позиционированию</span></div></div>`;
}

function tile(label, value, sub, kind = "") {
  return `<div class="kpi ${kind}"><div class="l">${label}</div><div class="v">${value}</div><div class="s">${sub || ""}</div></div>`;
}

// На что обратить внимание: две-четыре строки, только то, что требует действия.
function attention(d) {
  const out = [];
  const pos = (d.position && d.position.now) || [];
  const total = pos.reduce((s, p) => s + p.sale, 0);
  const none = (pos.find((p) => p.key === "no_price") || {}).sale || 0;
  const priced = total - none;
  const dear = pos.filter((p) => p.key.startsWith("d")).reduce((s, p) => s + p.sale, 0);
  const cheap = pos.filter((p) => p.key.startsWith("c")).reduce((s, p) => s + p.sale, 0);
  if (total && none / total >= 0.3) {
    out.push({ kind: "warn", text: `${Math.round(none / total * 100)}% выручки — без цены конкурента. По этим товарам рекомендация видит только ВР. Самые продаваемые из них — на листе «Итоги», блок «Проценить вручную». Почему так — вкладка «Почему мало цен».` });
  }
  if (priced && dear / priced >= 0.3) out.push({ kind: "dn", text: `Дороже цели ${Math.round(dear / priced * 100)}% выручки с ценой конкурента. Снижать стоит адресно: топ, где разрыв большой, — и смотреть, окупится ли.` });
  if (priced && cheap / priced >= 0.3) out.push({ kind: "up", text: `Дешевле цели ${Math.round(cheap / priced * 100)}% выручки с ценой конкурента — это запас для повышения.` });
  const comps = (d.slices || []).filter((x) => x.kind === "comp" && x.n);
  const weak = comps.filter((x) => d.items && x.n / d.items < 0.05);
  if (weak.length) {
    const names = weak.map((x) => `${String(x.title).replace(" (без учёта % позиц.)", "").replace("Против ", "")} (${fmtInt(x.n)})`).join(", ");
    out.push({ kind: "info", text: `Слабые ориентиры: ${names} — цена меньше чем у 5% товаров.` });
  }
  const dn = d.rec_enabled && d.bridge && d.bridge.down;
  if (dn && dn.items && dn.need_lift != null && dn.need_lift > 0.15) {
    out.push({ kind: "dn", text: `Снижения (${fmtInt(dn.items)} тов.) окупятся только при росте их продаж на ${fmtPct(dn.need_lift)}.` });
  }
  const low = d.rec_enabled ? d.low_vr_rec : d.low_vr;
  if (low) out.push({ kind: "info", text: `${fmtInt(low)} ${plural(low, "товар", "товара", "товаров")} с ВР ниже ${fmtInt(d.min_vr)}%${d.rec_enabled ? " и после рекомендаций" : ""}.` });
  return out.slice(0, 4);
}

function renderDash() {
  const d = S.analysis && S.analysis.dashboard;
  const el = $("#dash");
  if (!d) { el.innerHTML = ""; return; }
  const recOn = !!d.rec_enabled;
  const k = d.kinds || {};
  const arrow = (a, b, fmt) => `${fmt(a)} <span class="to">→ ${fmt(b)}</span>`;
  const dk = (a, b, better) => deltaKind(a, b, better);
  const pos = (d.position && d.position.now) || [];
  const total = pos.reduce((s, p) => s + p.sale, 0);
  const none = (pos.find((p) => p.key === "no_price") || {}).sale || 0;
  const share = total ? (total - none) / total : null;
  const priceTile = tile("Выручка с ценой конкурента", share == null ? "—" : `${Math.round(share * 100)}%`,
    `${fmtInt(d.priced)} из ${fmtInt(d.items)} товаров`, share != null && share < 0.5 ? "bad" : "");
  const tiles = recOn ? [
    tile("ВР группы", arrow(d.vr, d.vr_rec, fmtPct), d.vr != null && d.vr_rec != null ? fmtPp(d.vr_rec - d.vr) : "", dk(d.vr, d.vr_rec, "up")),
    tile("Прибыль", d.profit_rec != null && d.profit != null ? fmtSignedRub(d.profit_rec - d.profit) : "—", "при тех же продажах", dk(d.profit, d.profit_rec, "up")),
    priceTile,
    tile("Двигаем цены", `<span class="up-t">↑${fmtInt((k.up || 0) + (k.vr || 0))}</span> <span class="dn-t">↓${fmtInt(k.down || 0)}</span>`, `не трогаем ${fmtInt(k.keep || 0)}`),
  ] : [
    tile("ВР группы", fmtPct(d.vr), `ниже ${fmtInt(d.min_vr)}% — ${fmtInt(d.low_vr)} шт.`),
    tile("Товаров", fmtInt(d.items), `${fmtInt(d.priced)} с ценой конкурента`),
    priceTile,
    tile(`В коридоре ±${esc(d.tolerance)}%`, fmtInt(d.in_band), `из ${fmtInt(d.with_target)} с целью`),
  ];
  const notes = attention(d);
  const icon = { warn: "!", dn: "↓", up: "↑", info: "i" };
  const limitColors = { free: "#0f4c5c", capped: "#d97706", band: "#98a2b3", vr: "#7c3aed", top: "#0891b2", payback: "#b3541e", no_price: "#cbd5e1" };
  const limitRows = (d.limits || []).filter((x) => x.key !== "no_price");
  const limits = recOn && limitRows.length ? hbars(limitRows.map((x) => ({ label: x.title, value: x.items, color: limitColors[x.key] })), limitRows.reduce((n, x) => n + x.items, 0)) : "";
  const groupHead = (first) => (recOn
    ? `<th>${esc(first)}</th><th class="num">Шт.</th><th class="num">ВР</th><th class="num">После</th><th class="num">Δ ВР</th><th class="num">Ср. откл.</th><th class="num">После</th>`
    : `<th>${esc(first)}</th><th class="num">Шт.</th><th class="num">ВР</th><th class="num">Ср. откл.</th>`);
  const groupBody = (rows) => rows.map((x) => `<tr>
      <td>${esc(x.name || "пусто")}</td><td class="num">${fmtInt(x.items)}</td>
      <td class="num">${fmtPct(x.vr)}</td>
      ${recOn ? `<td class="num after">${fmtPct(x.vr_rec)}</td><td class="num delta ${deltaKind(x.vr, x.vr_rec, "up")}">${x.vr != null && x.vr_rec != null ? fmtPp(x.vr_rec - x.vr) : "—"}</td>` : ""}
      <td class="num">${fmtPct(x.mean_gap, true)}</td>
      ${recOn ? `<td class="num after">${fmtPct(x.mean_gap_rec, true)}</td>` : ""}
    </tr>`).join("");
  const movers = (d.top_changes || []).slice(0, 10);
  const moverRows = movers.map((x) => `<tr>
      <td class="nowrap"><b>${esc(x.article)}</b></td>
      <td class="muted ellip" title="${esc(x.name)}">${esc(x.name)}</td>
      <td class="num">${fmtMoney(x.cat_a)}</td>
      <td class="num after">${fmtMoney(x.rec)}</td>
      <td class="num">${x.change == null ? "—" : fmtPct(x.change, true)}</td>
      <td class="num delta ${deltaKind(0, x.profit_delta, "up")}">${fmtSignedRub(x.profit_delta)}</td>
      <td class="muted ellip" title="${esc(x.reason || "")}">${esc(x.reason || "")}</td>
    </tr>`).join("");
  el.innerHTML = `
    <div class="kpis live">${tiles.join("")}</div>
    ${notes.length ? `<div class="attn">${notes.map((n) => `<div class="attn-row k-${n.kind}"><i>${icon[n.kind]}</i><span>${esc(n.text)}</span></div>`).join("")}</div>` : ""}
    <div class="live-grid">
      ${positionMix(d)}
      ${recOn ? bridgeBlock(d) : ""}
      ${compBars(d, recOn)}
    </div>
    <details class="card flat more"><summary class="dash-title">Подробнее: выводы, что решило рекомендацию, крупные изменения, ABC и подгруппы</summary>
      ${(d.conclusions || []).length ? `<div class="dash-block"><div class="dash-title">Выводы</div><ul class="conclusions">${d.conclusions.map((t) => `<li>${esc(t)}</li>`).join("")}</ul></div>` : ""}
      ${limits ? `<div class="dash-block"><div class="dash-title">Что решило рекомендацию <span class="muted">· товары с ценой конкурента</span></div>${limits}
        <div class="muted small-text">«Упёрлись в лимит» — до цели не дошли из-за шага. «Окупаемость» — снижение глубже не окупится ростом продаж.</div></div>` : ""}
      ${recOn && movers.length ? `<div class="dash-block"><div class="dash-title">Что сильнее меняет прибыль</div>
        <div class="table-wrap scroll"><table class="mini"><thead><tr><th>Артикул</th><th>Наименование</th><th class="num">Кат. А</th><th class="num">Рек.</th><th class="num">Изм.</th><th class="num">Δ прибыли</th><th>Почему</th></tr></thead><tbody>${moverRows}</tbody></table></div></div>` : ""}
      ${cutTable("ВР по ABC", d.abc || [], groupHead("ABC"), groupBody(d.abc || []))}
      ${cutTable("По подгруппам", (d.subgroups || []).length > 1 ? d.subgroups : [], groupHead("Подгруппа"), groupBody(d.subgroups || []))}
      ${cutTable("По сегментам автопарка", d.segments || [], groupHead("Сегмент"), groupBody(d.segments || []))}
    </details>`;
}

// ------------------------------------------------------------------ покрытие ценами конкурентов

const isoDate = (v) => String(v || "").split("-").reverse().join(".");

const FUNNEL = [
  ["items", "Товаров в отчёте", "#475569"],
  ["yes", "Есть кросс «Да»", "#0f4c5c"],
  ["in_prices", "Артикул кросса есть в прайсах", "#146b80"],
  ["monitored", "У поставщиков из позиционирования", "#2b8a9e"],
  ["kept", "После отсева несопоставимого", "#4aa7b8"],
  ["final", "С ценой в отчёте", "#16a34a"],
];
const LOSS = {
  yes: "нет кросса «Да»", in_prices: "кросса нет ни в одном прайсе", monitored: "цена только у поставщиков вне позиционирования",
  kept: "отсеяно: комплект или цена в разы", final: "",
};

function renderCoverage() {
  const el = $("#coverage");
  if (!el) return;
  const a = S.analysis;
  const f = (a && a.funnel) || [];
  if (!a || !f.length) { el.innerHTML = `<div class="muted">Конкурентов нет: для группы в позиционировании не указаны бренды.</div>`; return; }
  const sc = a.screen || { enabled: true, ratio: 3 };
  const screened = Object.entries(a.screened || {}).map(([k, n]) => `${esc(k)} — ${fmtInt(n)}`).join("; ");
  const total = { items: 0, yes: 0, in_prices: 0, monitored: 0, kept: 0, final: 0 };
  const losses = {};
  const cards = f.map((c) => {
    const vals = { items: c.items, yes: c.yes, in_prices: c.in_prices, monitored: c.monitored, kept: Math.max(0, c.monitored - (c.excluded || 0)), final: c.final };
    Object.keys(total).forEach((k) => (total[k] += vals[k] || 0));
    let prev = null;
    const steps = FUNNEL.map(([key, title, color]) => {
      const v = vals[key] || 0;
      const lost = prev != null && key !== "final" ? Math.max(0, prev - v) : 0;
      if (lost && LOSS[key]) losses[LOSS[key]] = (losses[LOSS[key]] || 0) + lost;
      prev = key === "final" ? prev : v;
      return `<div class="fn-row">
        <span class="fn-name">${esc(title)}</span>
        <span class="fn-track"><i style="width:${pct(v, c.items)}%;background:${color}"></i></span>
        <span class="fn-num"><b>${fmtInt(v)}</b> <span class="muted">${pct(v, c.items)}%</span>${lost ? `<span class="fn-lost">−${fmtInt(lost)}</span>` : ""}</span>
      </div>`;
    }).join("");
    const r = c.reliability || {};
    const lead = (r.leaders || []).map((x) => `${esc(x.supplier)} ${fmtInt(x.rows)}`).join(", ");
    const extra = [
      c.by_no ? `из них по «Нет» ${fmtInt(c.by_no)}` : "",
      c.sold_share != null ? `доля продаж с ценой ${fmtPct(c.sold_share)}` : "",
      r.priced ? `цена только у одного поставщика: ${fmtInt(r.single)}` : "",
      r.second_gap != null ? `второй поставщик дороже на ${fmtPct(r.second_gap, true)}` : "",
      lead ? `чаще даёт минимум: ${lead}` : "",
    ].filter(Boolean).join(" · ");
    return `<div class="fn-card">
      <div class="fn-head"><b>${esc(c.name)}</b><span class="muted">${extra}</span></div>
      ${steps}
      ${(c.actions || []).length ? `<ul class="conclusions">${c.actions.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>` : ""}
    </div>`;
  }).join("");
  const worst = Object.entries(losses).sort((x, y) => y[1] - x[1]);
  const why = worst.length ? `<div class="why-box"><div class="dash-title">Где теряются цены — по всем конкурентам группы</div>
      ${hbars(worst.map(([label, value]) => ({ label, value, color: "#d97706" })), total.items)}
      <div class="muted small-text">Самая большая потеря — <b>${esc(worst[0][0])}</b>. ${
        worst[0][0].startsWith("нет кросса") ? "Цен станет больше, если добавить кроссы «Да»: «Выгрузить дыры в кроссах» даёт список товаров по продажам."
        : worst[0][0].startsWith("кросса нет") ? "Кроссы есть, но таких артикулов нет в прайсах: проверьте написание брендов в «Данных» и прайсы нужных поставщиков."
        : worst[0][0].startsWith("цена только") ? "Цены есть у поставщиков, которые по позиционированию этот бренд не мониторят. Снимите галочку «только поставщики из позиционирования» в настройках отчёта, если им можно верить."
        : "Отсев убирает комплекты и цены, отличающиеся от нашей в разы, — ниже можно поменять порог."
      }</div></div>` : "";
  const h = (a.history || []).slice().reverse();
  const history = h.length ? `<details class="card flat"><summary class="dash-title">Как менялось по сборкам отчёта</summary>
    <div class="table-wrap"><table class="mini"><thead><tr>
      <th>Дата</th><th class="num">Товаров</th><th class="num">С ценой</th><th class="num">Доля продаж с ценой</th>
      <th class="num">В позиционировании</th><th class="num">Медиана откл.</th><th class="num">Откл. с весом</th><th class="num">ВР</th>
    </tr></thead><tbody>${h.map((x) => `<tr>
      <td>${esc(isoDate(x.date))}</td><td class="num">${fmtInt(x.items)}</td><td class="num">${fmtInt(x.priced)}</td>
      <td class="num">${x.sold_share == null ? "—" : fmtPct(x.sold_share)}</td>
      <td class="num">${x.with_target ? `${fmtInt(x.in_band)} из ${fmtInt(x.with_target)}` : "—"}</td>
      <td class="num">${fmtPct(x.median, true)}</td><td class="num">${fmtPct(x.weighted, true)}</td><td class="num">${fmtPct(x.vr)}</td>
    </tr>`).join("")}</tbody></table></div></details>` : "";
  el.innerHTML = `
    <div class="card-head"><h2>Почему мало цен</h2><span class="sub">путь от товара до цены конкурента: на каком шаге теряются позиции</span>
      <div class="right"><button class="btn small" id="gaps-export" type="button">Выгрузить дыры в кроссах</button></div></div>
    <div class="kpis">
      ${tile("С ценой хотя бы одного конкурента", `${fmtInt(a.priced)} <span class="to">из ${fmtInt(a.kept)}</span>`, `${pct(a.priced, a.kept)}% товаров отчёта`)}
      ${tile("Кроссов «Да»", fmtInt(total.yes), `из ${fmtInt(total.items)} пар товар–конкурент`)}
      ${tile("Есть в прайсах", fmtInt(total.in_prices), `${pct(total.in_prices, total.yes)}% кроссов «Да»`)}
      ${tile("Цен в отчёте", fmtInt(total.final), `${pct(total.final, total.items)}% пар`)}
    </div>
    ${why}
    <div class="fn-grid">${cards}</div>
    <div class="row tight" style="flex-wrap:wrap; gap:12px; margin:12px 0 4px">
      <label><input type="checkbox" id="screen-on" ${sc.enabled ? "checked" : ""}> Отсеивать несопоставимое: комплекты и цены, отличающиеся от нашей в</label>
      <input class="input num-input" id="screen-ratio" type="number" min="1.5" step="0.5" value="${esc(sc.ratio)}" style="width:64px"><span class="muted">раза и больше</span>
    </div>
    ${screened ? `<div class="muted small-text">Отсеяно пар: ${screened}. Список — в выгрузке дыр, лист «Несопоставимо».</div>` : ""}
    ${history}`;
  const on = $("#screen-on");
  const ratio = $("#screen-ratio");
  if (on) on.onchange = saveScreen;
  if (ratio) ratio.onchange = saveScreen;
  const gaps = $("#gaps-export");
  if (gaps) gaps.onclick = exportGaps;
}

async function saveScreen() {
  const enabled = !!($("#screen-on") && $("#screen-on").checked);
  const ratio = Math.max(1.5, Number(($("#screen-ratio") || {}).value) || 3);
  try {
    await api("/api/options", { options: { screen: { enabled, kits: true, ratio } } });
  } catch (e) {
    toast(e.message, "bad");
    return;
  }
  analyze();
}

async function exportGaps() {
  const b = $("#gaps-export");
  if (b) b.disabled = true;
  try {
    const r = await api("/api/crosses-gaps", { folder: S.folder, options: S.opts });
    toast(`Готово: ${r.name} — без «Да» ${fmtInt(r.no_yes)}, «Да» без цены ${fmtInt(r.yes_no_price)}, отсеяно ${fmtInt(r.excluded)}`, "good");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    if (b) b.disabled = false;
  }
}

// ------------------------------------------------------------------ написание брендов

function renderBrandHints() {
  const el = $("#px-hints");
  if (!el) return;
  const hints = (S.scan && S.scan.brand_hints) || [];
  const aliases = (S.scan && S.scan.brand_aliases) || [];
  const ignored = (S.scan && S.scan.brand_ignore) || [];
  if (!hints.length && !aliases.length && !ignored.length) {
    el.classList.add("hidden");
    el.innerHTML = "";
    return;
  }
  if (!S.hintPick) S.hintPick = {};
  el.classList.remove("hidden");
  const picked = hints.filter((h) => S.hintPick[h.brand]);
  el.innerHTML = `
    <div class="card-head"><h2>Написание брендов</h2>
      <span class="sub">Бренд в прайсе или кроссах похож на конкурента, но не совпал. «Считать как» — и его цены пойдут в сравнение. «Не тот бренд» — подсказка больше не появится</span></div>
    ${hints.length ? `<div class="row tight" style="margin-bottom:8px; flex-wrap:wrap">
        <label class="check"><input type="checkbox" id="hint-all" ${picked.length === hints.length ? "checked" : ""}> выбрать все · ${fmtInt(hints.length)}</label>
        <button type="button" class="btn small primary" id="hint-apply" ${picked.length ? "" : "disabled"}>Считать как конкурента · ${fmtInt(picked.length)}</button>
        <button type="button" class="btn small" id="hint-ignore" ${picked.length ? "" : "disabled"}>Не тот бренд · ${fmtInt(picked.length)}</button>
      </div>
      <div class="table-wrap"><table class="mini"><thead><tr>
      <th></th><th>Как написано</th><th class="num">Строк</th><th>Где</th><th>Похож на</th><th></th></tr></thead><tbody>
      ${hints.map((h, i) => `<tr><td><input type="checkbox" data-pick="${i}" ${S.hintPick[h.brand] ? "checked" : ""}></td>
        <td><b>${esc(h.brand)}</b></td><td class="num">${fmtInt(h.rows)}</td>
        <td class="muted">${esc((h.where || []).join(", "))}</td><td>${esc(h.suggest)}</td>
        <td class="nowrap"><button type="button" class="btn small" data-alias="${i}">Считать как ${esc(h.suggest)}</button>
          <button type="button" class="btn small ghost" data-ignore="${i}">Не тот бренд</button></td></tr>`).join("")}
    </tbody></table></div>` : `<div class="muted">Похожих несовпавших брендов нет.</div>`}
    ${aliases.length ? `<div class="px-brands" style="margin-top:10px">Считаются как конкурент:
      ${aliases.map((x) => `<span class="tag hit">${esc(x[0])} → ${esc(x[1])}
        <button type="button" class="btn small ghost" data-unalias="${esc(x[0])}">убрать</button></span>`).join(" ")}</div>` : ""}
    ${ignored.length ? `<details style="margin-top:10px"><summary class="muted">Отклонённые подсказки · ${fmtInt(ignored.length)}</summary>
      <div class="px-brands" style="margin-top:6px">${ignored.map((b) => `<span class="tag">${esc(b)}
        <button type="button" class="btn small ghost" data-unignore="${esc(b)}">вернуть</button></span>`).join(" ")}</div></details>` : ""}`;
  el.querySelectorAll("[data-pick]").forEach((cb) => (cb.onchange = () => {
    const h = hints[Number(cb.dataset.pick)];
    if (h) S.hintPick[h.brand] = cb.checked;
    renderBrandHints();
  }));
  const all = $("#hint-all");
  if (all) all.onchange = () => { hints.forEach((h) => (S.hintPick[h.brand] = all.checked)); renderBrandHints(); };
  const apply = $("#hint-apply");
  if (apply) apply.onclick = () => saveAlias({ add_many: picked.map((h) => [h.brand, h.suggest]) });
  const ignore = $("#hint-ignore");
  if (ignore) ignore.onclick = () => saveAlias({ ignore_many: picked.map((h) => h.brand) });
  el.querySelectorAll("[data-alias]").forEach((b) => (b.onclick = () => {
    const h = hints[Number(b.dataset.alias)];
    if (h) saveAlias({ add: [h.brand, h.suggest] });
  }));
  el.querySelectorAll("[data-ignore]").forEach((b) => (b.onclick = () => {
    const h = hints[Number(b.dataset.ignore)];
    if (h) saveAlias({ ignore: h.brand });
  }));
  el.querySelectorAll("[data-unalias]").forEach((b) => (b.onclick = () => saveAlias({ remove: b.dataset.unalias })));
  el.querySelectorAll("[data-unignore]").forEach((b) => (b.onclick = () => saveAlias({ unignore: b.dataset.unignore })));
}

async function saveAlias(body) {
  let r;
  try {
    r = await api("/api/brand-aliases", body);
  } catch (e) {
    toast(e.message, "bad");
    return;
  }
  S.hintPick = {};
  if (!r.reread) {
    // «Не тот бренд» не меняет цены — прайсы не перечитываем.
    if (S.scan) {
      S.scan.brand_ignore = r.brand_ignore;
      const gone = new Set((r.brand_ignore || []).map((b) => brandKey(b)));
      S.scan.brand_hints = (S.scan.brand_hints || []).filter((h) => !gone.has(brandKey(h.brand)));
    }
    renderBrandHints();
    toast("Подсказка убрана", "good");
    return;
  }
  toast(r.added > 1 ? `Сохранено синонимов: ${r.added} — перечитываю прайсы` : "Сохранено — перечитываю прайсы", "good");
  if (await rescan(false)) analyze();
}

function fmtSignedInt(v) {
  const n = Math.round(Number(v) || 0);
  if (!n) return "0";
  return (n > 0 ? "+" : "−") + Math.abs(n).toLocaleString("ru-RU");
}

function renderPreview() {
  const a = S.analysis;
  if (!a.preview.length) {
    $("#preview").innerHTML = `<div class="empty"><h3>Нет строк</h3><p>После фильтров не осталось товаров — ослабьте фильтры в настройках группы.</p></div>`;
    return;
  }
  const comps = a.competitors;
  const recOn = !!(S.opts.rec && S.opts.rec.enabled);
  const transit = !!a.has_transit;
  const sales = a.sales_period ? `Продажи (${esc(a.sales_period)}), шт` : "Продажи, шт";
  $("#preview-note").textContent = "Первые строки отчёта по убыванию продаж в штуках";
  const head = `<tr><th>Код</th><th>Артикул</th><th>Наименование</th><th>Сегмент</th><th>Автопарк</th>
    <th class="num">${sales}</th><th>ABC</th><th class="num">Кат. А</th>
    ${recOn ? `<th class="num">Рек. кат. А</th><th class="num">Изм.</th>` : ""}
    ${comps.map((c) => `<th style="color:var(--violet)">${esc(c.name)}</th>`).join("")}
    <th class="num">Остаток</th>${transit ? `<th class="num">В пути</th><th class="num">Потенциал</th>` : ""}<th class="num">СБС</th>${comps.map((c) => `<th class="num" style="color:var(--amber)">${esc(c.name)} цена</th>`).join("")}</tr>`;
  const rows = a.preview.map((r) => `<tr>
    <td class="nowrap">${esc(r.code)}</td><td class="nowrap"><b>${esc(r.article)}</b></td><td class="muted" style="max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap" title="${esc(r.makes || "")}">${esc(r.name)}</td>
    <td class="nowrap">${esc(r.segment || "")}</td><td class="nowrap" title="${esc(r.makes || "")}">${esc(r.fleet || "")}</td>
    <td class="num">${fmtInt(r.sales_qty)}</td><td>${esc(r.abc || "")}</td>
    <td class="num" style="background:#fffde6">${fmtMoney(r.cat_a)}</td>
    ${recOn ? `<td class="num" style="background:#eaf2fd" title="${esc(r.reason || "")}">${fmtMoney(r.rec)}</td><td class="num">${r.rec_change == null ? "—" : fmtPct(r.rec_change, true)}</td>` : ""}
    ${comps.map((c) => `<td class="nowrap">${(r.crosses[c.key] || []).map((x) => `<span class="${x.yes === null ? "x-acp" : x.yes ? "x-yes" : "x-no"}">${esc(x.a)}</span>`).join(", ") || '<span class="muted">—</span>'}</td>`).join("")}
    <td class="num">${fmtInt(r.stock)}</td>${transit ? `<td class="num">${fmtInt(r.in_transit)}</td><td class="num">${fmtInt(r.stock == null && r.in_transit == null ? null : (Number(r.stock) || 0) + (Number(r.in_transit) || 0))}</td>` : ""}<td class="num">${fmtMoney(r.sbs)}</td>
    ${comps.map((c) => `<td class="num price" title="${esc(r.best[c.key] || "")}">${r.prices[c.key] ? fmtMoney(r.prices[c.key]) : '<span class="muted" style="font-weight:400">—</span>'}</td>`).join("")}
  </tr>`).join("");
  $("#preview").innerHTML = `<div class="table-wrap"><table><thead>${head}</thead><tbody>${rows}</tbody></table></div>
    <div class="muted" style="font-size:12.5px; margin-top:8px">Кроссы: обычные — «Да», <span class="x-no">курсив — «Нет»</span>, <span class="x-acp">серые — из АЦПКД</span>. Наведите на цену — увидите поставщика.</div>`;
}

// ------------------------------------------------------------------ анализ

let analyzeTimer = null;
function analyze(delay = 450) {
  persist();
  clearTimeout(analyzeTimer);
  const state = $("#analyze-state");
  if (state) state.textContent = "Пересчитываю…";
  analyzeTimer = setTimeout(runAnalyze, delay);
}

let analyzeRunning = false;
let analyzeAgain = false;
async function runAnalyze() {
  if (!S.scan || (!S.opts.brand && !S.opts.group && !listedGroups().length)) return;
  if (analyzeRunning) { analyzeAgain = true; return; }
  analyzeRunning = true;
  const seq = ++S.analyzeSeq;
  const state = $("#analyze-state");
  if (state) state.textContent = "Считаю отчёт… фильтры можно менять дальше";
  document.querySelectorAll("#dash, #coverage, #preview").forEach((el) => el.classList.add("loading-veil"));
  try {
    const a = await api("/api/analyze", { folder: S.folder, options: S.opts });
    if (seq !== S.analyzeSeq) return;
    S.analysis = a;
    S.analysisFor = `${S.opts.brand}\t${S.opts.group}`;
    if (!S.editing && a.price_files) {
      S.priceFiles = a.price_files;
      renderPriceFiles();
    }
    if (a.price_date) {
      const state = $("#analyze-state");
      if (state) state.textContent = `Прайсы от ${a.price_date}`;
    }
    renderCounter(); renderAbc(); renderPriceControls(); renderCrosses(); renderSources(); renderDash(); renderCoverage(); renderPreview(); renderPosLines(); renderRatesNote(); renderContext();
    if (S.settings && !S.settingsDirty) renderRates();
    updateSummary();
  } catch (e) {
    if (seq === S.analyzeSeq) toast(e.message, "bad");
  } finally {
    analyzeRunning = false;
    if (seq === S.analyzeSeq) document.querySelectorAll(".loading-veil").forEach((el) => el.classList.remove("loading-veil"));
    if (analyzeAgain) {
      analyzeAgain = false;
      runAnalyze();
    }
  }
}

function paintBuildAll() {
  const all = $("#build-all");
  if (!all || all.dataset.busy) return;
  const n = ((S.scan && S.scan.batch) || []).reduce((sum, p) => sum + ((p.groups || []).length), 0);
  all.disabled = !S.scan || !S.scan.workspace || !n;
  all.textContent = n ? `Все группы · ${n}` : "Все группы";
}

function updateSummary() {
  const fresh = analysisFresh();
  const a = fresh ? S.analysis : null;
  const visible = S.scan && !$("#flow").classList.contains("hidden");
  const ok = !!(a && visible);
  $("#build").disabled = !ok || !a.kept;
  paintBuildAll();
  if (!visible) { $("#summary").innerHTML = `<span>Выберите папку с выгрузками</span>`; return; }
  if (!ok) {
    $("#summary").innerHTML = `<div><b>${esc(S.opts.brand)} · ${esc(S.opts.group)}</b></div><span>считаю…</span>`;
    return;
  }
  const sheets = S.opts.subgroup_mode === "sheets" && a.subgroups.length > 1 ? a.subgroups.length : 1;
  const d = a.dashboard;
  const vr = d && d.vr != null
    ? `<div><b>${fmtPct(d.vr)}</b> <span>ВР${d.rec_enabled && d.vr_rec != null ? ` → ${fmtPct(d.vr_rec)}` : ""}</span></div>`
    : "";
  $("#summary").innerHTML = `
    <div><b>${esc(S.opts.brand)} · ${esc(S.opts.group)}</b></div>
    <div><b>${fmtInt(a.kept)}</b> <span>товаров</span></div>
    <div><b>${fmtInt(a.priced)}</b> <span>с ценой</span></div>
    ${vr}
    <div><b>${sheets}</b> <span>${plural(sheets, "лист", "листа", "листов")}</span></div>`;
}

// ------------------------------------------------------------------ формирование

function recPayload() {
  return {
    ...S.opts,
    custom: customSnapshot(),
    rec_scope: S.recScope,
    rec_source: S.recSource,
    rec_group_source: S.recGroupSource,
    rec_mine: S.recMine,
    rec_groups: S.recGroups,
  };
}

async function build() {
  const btn = $("#build");
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Формирую…`;
  try {
    const r = await api("/api/build", { folder: S.folder, options: recPayload() });
    $("#result").innerHTML = `<div class="result">
      <div><div class="t">Готово: ${esc(r.name)}</div>
      <div class="muted" style="font-size:13px">${fmtInt(r.rows)} товаров, ${fmtInt(r.priced)} с ценой конкурента, исключено ${fmtInt(r.excluded)} · листы: ${r.sheets.map(esc).join(", ")}</div>
      ${r.brief ? `<div class="muted" style="font-size:13px">Разбор группы: ${esc(r.brief)}</div>` : ""}</div>
      <div class="spacer" style="flex:1"></div>
      <button class="btn primary" data-open="${esc(r.path)}">Открыть файл</button>
      ${r.brief ? `<button class="btn" data-open="${esc(r.brief)}">Открыть разбор</button>` : ""}
      <button class="btn" data-open="${esc(r.folder)}">Открыть папку</button></div>`;
    $("#result").querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => api("/api/open", { path: b.dataset.open }).catch((e) => toast(e.message, "bad"))));
    $("#result").scrollIntoView({ behavior: "smooth", block: "nearest" });
    toast("Отчёт сформирован", "good");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    btn.textContent = "Сформировать отчёт";
    updateSummary();
  }
}

async function buildAll() {
  const btn = $("#build-all");
  btn.dataset.busy = "1";
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Собираю…`;
  try {
    const custom = customSnapshot();
    const r = await api("/api/build-all", {
      folder: S.folder,
      open: true,
      options: {
        ...custom,
        report_kind: S.opts.report_kind,
        sheet_filters: S.opts.sheet_filters,
        ...recPayload(),
        custom,
      },
    });
    const made = (r.made || []).map((m) => `<div>${esc(m.brand)} · ${esc(m.group)} — ${esc(m.name)} · ${fmtInt(m.rows)} товаров${m.priced ? `, с ценой ${fmtInt(m.priced)}` : ""}${m.note ? ` · ${esc(m.note)}` : ""}</div>`).join("");
    const skipped = (r.skipped || []).map((s) => `<div>${esc(s.folder)}${s.brand ? ` · ${esc(s.brand)} / ${esc(s.group)}` : ""} — ${esc(s.reason)}</div>`).join("");
    const n = (r.made || []).length;
    $("#result").innerHTML = `<div class="result">
      <div>
        <div class="t">${n ? `Готово: ${fmtInt(n)} ${plural(n, "файл", "файла", "файлов")}` : "Файлы не созданы"}</div>
        <div class="muted" style="font-size:13px">Файлы Excel — только группы с полным набором: ПСТ, СБС, Продажи и АЦПКД. В разборе — каждая группа из списка слева, даже если файл не создан.</div>
        ${r.brief ? `<div class="muted" style="font-size:13px">Текстовый разбор: ${esc(r.brief)}</div>` : ""}
        ${(r.missing_rates || []).length ? `<div class="note">Не хватает курсов, из‑за этого часть цен пустая: ${(r.missing_rates || []).map((m) => `<b>${esc(m.currency)}</b> (${fmtInt(m.items)})`).join(", ")}. <button class="btn small ghost" id="goto-rates-batch">Задать курсы</button></div>` : ""}
        ${made}
        ${skipped ? `<div style="margin-top:8px">${skipped}</div>` : ""}
      </div>
      <div class="spacer" style="flex:1"></div>
      ${r.brief ? `<button class="btn" data-open="${esc(r.brief)}">Открыть разбор</button>` : ""}
      ${n ? `<button class="btn" data-open="${esc(r.folder)}">Открыть папку</button>` : ""}
    </div>`;
    $("#result").querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => api("/api/open", { path: b.dataset.open }).catch((e) => toast(e.message, "bad"))));
    const ratesBtn = $("#goto-rates-batch");
    if (ratesBtn) ratesBtn.onclick = () => showTab("settings");
    $("#result").scrollIntoView({ behavior: "smooth", block: "nearest" });
    toast(r.made && r.made.length ? `Собрано файлов: ${r.made.length}` : "Ни одна группа не собралась", r.made && r.made.length ? "good" : "bad");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    delete btn.dataset.busy;
    btn.innerHTML = "";
    updateSummary();
  }
}

function blankStockProfile() {
  const rec = clone(S.init.rec || {});
  rec.enabled = true;
  rec.up = { ...(S.init.default_up || {}), ...(rec.up || {}) };
  return {
    report_kind: "full", money_base: "sales", filters: {}, abc: [], price_source: "both",
    monitoring_only: true, in_stock_only: false, base: "pst", sources: {}, cross_split: "none",
    cross_modes: {}, subgroup_mode: "sheets", green_when: "above",
    suggest_threshold: 0.3, rec_scope: "all", rec_source: "yours", rec_groups: {},
    rec_group_source: {}, rec_mine: {}, rec,
  };
}

function ensureStockProfiles() {
  const saved = (S.init.options && S.init.options.stock_profiles) || {};
  const merge = (raw) => {
    const base = blankStockProfile();
    if (!raw || typeof raw !== "object") return base;
    return {
      ...base, ...raw,
      rec: { ...base.rec, ...(raw.rec || {}), up: { ...(base.rec.up || {}), ...((raw.rec || {}).up || {}) } },
      filters: foldClearance(raw.filters || {}),
      abc: raw.abc || [],
    };
  };
  S.stockProfiles = { sale: merge(saved.sale), overstock: merge(saved.overstock) };
}

function stockProfile() {
  return S.stockProfiles[S.stockEdit === "overstock" ? "overstock" : "sale"];
}

function stockPayload(slice) {
  const p = S.stockProfiles[slice];
  return { ...p, stock_slice: slice, rec_scope: "all", rec_source: "yours", rec: p.rec };
}

let stockTimer = null;
function persistStock() {
  clearTimeout(stockTimer);
  stockTimer = setTimeout(() => {
    api("/api/options", { options: { stock_profiles: S.stockProfiles } }).catch(() => {});
  }, 300);
}

function renderStockForm() {
  const editWrap = $("#stock-edit-wrap");
  if (editWrap) editWrap.classList.toggle("hidden", S.stockSlice !== "both");
  seg($("#stock-edit"), S.stockEdit === "overstock" ? "overstock" : "sale", (v) => {
    S.stockEdit = v;
    renderStockForm();
  });
  const p = stockProfile();
  if (!p) return;
  p.report_kind = "full";
  seg($("#sk-money"), p.money_base === "potential" ? "potential" : "sales", (v) => { p.money_base = v; persistStock(); });
  const enabled = $("#sk-rec-enabled");
  enabled.checked = !!p.rec.enabled;
  enabled.onchange = () => { p.rec.enabled = enabled.checked; persistStock(); };
  const targetGroup = p.rec.target === "avg" || p.rec.target === "min" ? "all" : (p.rec.target === "first" ? "first" : "main");
  seg($("#sk-target"), targetGroup, (v) => {
    p.rec.target = v === "all" ? (p.rec.target === "min" ? "min" : "avg") : v;
    renderStockForm();
    persistStock();
  });
  const all = $("#sk-target-all");
  all.classList.toggle("hidden", targetGroup !== "all");
  seg(all, p.rec.target === "min" ? "min" : "avg", (v) => { p.rec.target = v; persistStock(); });
  for (const [id, key] of [["#sk-tolerance", "tolerance"], ["#sk-down", "down"], ["#sk-vr", "min_vr"], ["#sk-round", "rounding"]]) {
    const input = $(id);
    input.value = p.rec[key] ?? "";
    input.oninput = () => {
      const n = toNum(input.value);
      if (n === "") return;
      p.rec[key] = n;
      persistStock();
    };
  }
  $("#sk-up").innerHTML = (S.init.segments || []).map((s) => `<label>${esc(s)}
      <input class="input num-input" data-seg="${esc(s)}" type="number" min="0" max="40" step="0.5" value="${esc((p.rec.up || {})[s] ?? "")}"></label>`).join("");
  $("#sk-up").querySelectorAll("input").forEach((inp) => (inp.oninput = () => {
    const n = toNum(inp.value);
    if (n === "") return;
    p.rec.up = p.rec.up || {};
    p.rec.up[inp.dataset.seg] = n;
    persistStock();
  }));
  const skip = new Set(["sale", "overstock", "discontinued"]);
  const filters = S.init.filters || {};
  $("#sk-filters").innerHTML = Object.entries(filters).filter(([k]) => !skip.has(k)).map(([k, def]) => `
    <div class="filter"><div class="t">${esc(def.title)}</div>
      ${segHtml(def.choices, p.filters[k] || def.default, `data-sk-filter="${k}"`)}
    </div>`).join("");
  $("#sk-filters").querySelectorAll("[data-sk-filter]").forEach((el) => {
    const k = el.dataset.skFilter;
    seg(el, p.filters[k] || filters[k].default, (v) => {
      if (v === filters[k].default) delete p.filters[k]; else p.filters[k] = v;
      persistStock();
    });
  });
  const abc = S.init.abc_values || [];
  const allAbc = !p.abc.length;
  $("#sk-abc").innerHTML = [`<button type="button" class="chip ${allAbc ? "on" : ""}" data-v="*">Все</button>`,
    ...abc.map((v) => `<button type="button" class="chip ${!allAbc && p.abc.includes(v) ? "on" : ""}" data-v="${esc(v)}">${v || "пусто"}</button>`)].join("");
  $("#sk-abc").querySelectorAll(".chip").forEach((b) => (b.onclick = () => {
    const v = b.dataset.v;
    if (v === "*") p.abc = [];
    else if (p.abc.includes(v)) p.abc = p.abc.filter((x) => x !== v);
    else p.abc = [...p.abc, v];
    renderStockForm();
    persistStock();
  }));
  const sources = S.init.price_sources || {};
  $("#sk-price-source").innerHTML = Object.entries(sources).map(([k, t]) =>
    `<button type="button" data-v="${esc(k)}">${esc(t)}</button>`).join("");
  seg($("#sk-price-source"), p.price_source, (v) => { p.price_source = v; persistStock(); });
  const mon = $("#sk-monitoring");
  mon.checked = p.monitoring_only !== false;
  mon.onchange = () => { p.monitoring_only = mon.checked; persistStock(); };
  const stocked = $("#sk-instock");
  stocked.checked = !!p.in_stock_only;
  stocked.onchange = () => { p.in_stock_only = stocked.checked; persistStock(); };
  const fieldChoices = S.init.sources || {};
  const fields = S.init.fields || {};
  p.base = "pst";
  const fallback = "own_acp";
  $("#sk-sources").innerHTML = Object.entries(fields).map(([k, def]) => {
    const cur = p.sources[k] || fallback;
    return `<div class="src-row">
      <div><div class="t">${esc(def.title)}</div></div>
      <div>${segHtml(fieldChoices, cur, `data-sk-src="${k}"`)}<div class="opts">Наши: ${esc(def.own)} · АЦПКД: ${esc(def.acp)}</div></div>
    </div>`;
  }).join("");
  $("#sk-sources").querySelectorAll("[data-sk-src]").forEach((el) => {
    const key = el.dataset.skSrc;
    seg(el, p.sources[key] || fallback, (v) => {
      p.sources[key] = v;
      persistStock();
    });
  });
}

async function buildStock() {
  const btn = $("#stock-build-all");
  btn.disabled = true;
  const label = btn.textContent;
  btn.innerHTML = `<span class="spinner"></span> Собираю…`;
  const slices = S.stockSlice === "overstock" ? ["overstock"] : S.stockSlice === "sale" ? ["sale"] : ["sale", "overstock"];
  const names = { sale: "Распродажа", overstock: "Оверсток" };
  const made = [];
  const skipped = [];
  let folder = "";
  try {
    for (const slice of slices) {
      const r = await api("/api/build-all", { folder: S.folder, open: false, options: stockPayload(slice) });
      folder = r.folder || folder;
      for (const item of r.made || []) made.push({ ...item, slice: names[slice] });
      for (const item of r.skipped || []) skipped.push({ ...item, slice: item.slice || names[slice] });
    }
    const madeHtml = made.map((m) => `<div>${esc(m.slice)} · ${esc(m.brand)} · ${esc(m.group)} — ${esc(m.name)} · ${fmtInt(m.rows)} товаров</div>`).join("");
    const skippedHtml = skipped.map((s) => `<div>${esc(s.slice || "")}${s.slice ? " · " : ""}${esc(s.brand || "")}${s.group ? ` / ${esc(s.group)}` : ""} — ${esc(s.reason)}</div>`).join("");
    $("#stock-result").innerHTML = `<div class="result">
      <div>
        <div class="t">${made.length ? `Готово: ${fmtInt(made.length)} ${plural(made.length, "файл", "файла", "файлов")}` : "Файлы не созданы"}</div>
        ${madeHtml}
        ${skippedHtml ? `<div style="margin-top:8px">${skippedHtml}</div>` : ""}
      </div>
      <div class="spacer" style="flex:1"></div>
      ${folder ? `<button class="btn" data-open="${esc(folder)}">Открыть папку</button>` : ""}
    </div>`;
    $("#stock-result").querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => api("/api/open", { path: b.dataset.open }).catch((e) => toast(e.message, "bad"))));
    toast(made.length ? `Собрано файлов: ${made.length}` : "Ни одна группа не собралась", made.length ? "good" : "bad");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}

async function checkFolder() {
  const btn = $("#check-folder");
  btn.disabled = true;
  try {
    const r = await api("/api/check", { folder: S.folder });
    toast(`Проверка папки сохранена: ${r.name}`, "good");
  } catch (e) {
    toast(e.message, "bad");
  } finally {
    btn.disabled = false;
  }
}

// ------------------------------------------------------------------ настройки

function bindSettings() {
  $("#settings-save").onclick = saveSettings;
  $("#settings-reset").onclick = async () => {
    if (!confirm("Вернуть курсы, ГЗ, скидки и марки авто к зашитым в программу значениям?")) return;
    applySettings(await api("/api/settings/reset"));
    toast("Вернул зашитые значения", "good");
  };
  $("#settings-import").onclick = async () => {
    try {
      const r = await api("/api/settings/import", { folder: S.folder });
      applySettings(r);
      toast(`Взял из папки: ${r.imported.join(", ")}`, "good");
    } catch (e) { toast(e.message, "bad"); }
  };
  $("#rate-add").onclick = () => { S.settings.rates.push(["", ""]); markDirty(); renderRates(); };
  $("#city-add").onclick = () => { S.settings.cities.unshift(["", ""]); markDirty(); renderCities(); };
  $("#gz-add").onclick = () => { S.settings.gz.unshift(["", ""]); $("#gz-search").value = ""; markDirty(); renderGz(); };
  $("#disc-add").onclick = () => { S.settings.discounts.unshift(["", "", "", "", "", ""]); $("#disc-search").value = ""; markDirty(); renderDiscounts(); };
  $("#gz-search").oninput = renderGz;
  $("#disc-search").oninput = renderDiscounts;
  $("#cars-search").oninput = renderCars;
  $("#pos-search").oninput = renderPositioning;
  $("#pos-add").onclick = () => {
    if (!S.settings) return;
    S.settings.positioning.unshift(["", "", "", "", "", "", "", "", ""]);
    $("#pos-search").value = "";
    markDirty();
    renderPositioning();
  };
  $("#cars-add").onclick = () => {
    if (!S.settings) return;
    S.settings.cars.unshift(["", "", "", "Японский", "Средний", ""]);
    $("#cars-search").value = "";
    markDirty();
    renderCars();
  };
  ["#threshold", "#output"].forEach((id) => ($(id).oninput = markDirty));
  $("#open-result").onchange = markDirty;
  window.addEventListener("beforeunload", (e) => { if (S.settingsDirty) e.preventDefault(); });
}

async function loadSettings() {
  try { applySettings(await api("/api/settings")); } catch (e) { toast(e.message, "bad"); }
}

function applySettings(data) {
  S.settings = {
    rates: Object.entries(data.currency_rates || {}),
    gz: (data.gz || []).map((x) => [...x]),
    discounts: (data.discounts || []).map((x) => [...x]),
    cars: (data.cars || []).map((x) => [...x]),
    positioning: (data.positioning || []).map((x) => [...x]),
    cities: (data.city_codes || []).map((x) => [...x]),
    fleets: data.fleets || [],
    segments: data.segments || ["Бюджетный", "Средний", "Премиум", "Не определён"],
    raw: data,
  };
  $("#threshold").value = Math.round((data.suggest_threshold ?? 0.3) * 100);
  $("#output").value = data.output || "Результат";
  $("#open-result").checked = data.open_result !== false;
  $("#settings-state").textContent = data.customized ? "Используются ваши сохранённые значения" : "Используются зашитые значения";
  S.settingsDirty = false;
  $("#settings-dirty").classList.add("hidden");
  renderRates(); renderGz(); renderDiscounts(); renderPositioning(); renderCars(); renderCities();
}

function renderCities() {
  const list = S.settings.cities;
  $("#cities").innerHTML = list.map((row, i) => `<tr>
    <td><input class="input" data-r="${i}" data-c="0" value="${esc(row[0])}" placeholder="RZD"></td>
    <td><input class="input" data-r="${i}" data-c="1" value="${esc(row[1])}" placeholder="Раздольное"></td>
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Своих кодов нет — работают встроенные</td></tr>`;
  bindCells($("#cities"), list, renderCities);
}

function markDirty() {
  S.settingsDirty = true;
  $("#settings-dirty").classList.remove("hidden");
}

function bindCells(tbody, list, rerender) {
  tbody.querySelectorAll("input[data-r], select[data-r]").forEach((inp) => {
    const write = () => {
      list[+inp.dataset.r][+inp.dataset.c] = inp.type === "checkbox" ? (inp.checked ? "да" : "") : inp.value;
      markDirty();
    };
    inp.oninput = write;
    inp.onchange = write;
  });
  tbody.querySelectorAll("button[data-del]").forEach((b) => (b.onclick = () => { list.splice(+b.dataset.del, 1); markDirty(); rerender(); }));
}

const delBtn = (i) => `<button class="btn small ghost danger" data-del="${i}" title="Удалить">✕</button>`;

function renderRates() {
  const list = S.settings.rates;
  const missing = Object.fromEntries((S.analysis ? S.analysis.missing_rates : []).map((m) => [m.currency.toUpperCase(), m.items]));
  const order = list.map((row, i) => [row, i]).sort(([a], [b]) => (missing[b[0].toUpperCase()] ?? -1) - (missing[a[0].toUpperCase()] ?? -1));
  $("#rates").innerHTML = order.map(([[k, v], i]) => `<tr class="${missing[k.toUpperCase()] ? "need" : ""}">
    <td><input class="input" data-r="${i}" data-c="0" value="${esc(k)}" placeholder="JPY">
      ${missing[k.toUpperCase()] ? `<span class="pill warn">нужен курс · ${fmtInt(missing[k.toUpperCase()])} тов.</span>` : ""}</td>
    <td class="num"><input class="input num-input" data-r="${i}" data-c="1" value="${esc(v)}" inputmode="decimal" placeholder="—"></td>
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Курсов нет — закупочная цена в рублях не посчитается</td></tr>`;
  bindCells($("#rates"), list, renderRates);
}

function filtered(list, query, cols) {
  const q = query.trim().toLowerCase();
  return list.map((row, i) => [row, i]).filter(([row]) => !q || cols.some((c) => String(row[c] ?? "").toLowerCase().includes(q)));
}

function renderGz() {
  const list = S.settings.gz;
  const rows = filtered(list, $("#gz-search").value, [0]);
  $("#gz-count").textContent = `${list.length} групп`;
  $("#gz").innerHTML = rows.map(([[n, c], i]) => `<tr>
    <td><input class="input" data-r="${i}" data-c="0" value="${esc(n)}"></td>
    <td class="num"><input class="input num-input" data-r="${i}" data-c="1" value="${esc(c)}" inputmode="decimal"></td>
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Ничего не найдено</td></tr>`;
  bindCells($("#gz"), list, renderGz);
}

function renderDiscounts() {
  const list = S.settings.discounts;
  const rows = filtered(list, $("#disc-search").value, [2, 3, 4]);
  $("#disc-count").textContent = `${list.length} строк`;
  $("#discounts").innerHTML = rows.slice(0, 400).map(([row, i]) => `<tr>
    <td><input class="input" data-r="${i}" data-c="2" value="${esc(row[2])}"></td>
    <td><input class="input" data-r="${i}" data-c="3" value="${esc(row[3])}"></td>
    <td><input class="input" data-r="${i}" data-c="4" value="${esc(row[4])}" style="width:110px"></td>
    <td class="num"><input class="input num-input" data-r="${i}" data-c="5" value="${esc(row[5])}" inputmode="decimal"></td>
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="5" class="muted">Ничего не найдено</td></tr>`;
  bindCells($("#discounts"), list, renderDiscounts);
}

function renderPositioning() {
  if (!S.settings) return;
  const list = S.settings.positioning;
  const rows = filtered(list, $("#pos-search").value, [0, 1, 2, 3, 5, 7]);
  $("#pos-count").textContent = `${list.length} строк${S.settings.raw && S.settings.raw.positioning_custom ? " · ваша таблица" : " · зашитое"}`;
  const cell = (i, c, v, ph) => `<td><input class="input" data-r="${i}" data-c="${c}" value="${esc(v ?? "")}" placeholder="${ph || ""}"></td>`;
  const pct = (i, c, v) => `<td class="num"><input class="input num-input" data-r="${i}" data-c="${c}" value="${esc(v ?? "")}" inputmode="decimal"></td>`;
  $("#positioning").innerHTML = rows.slice(0, 200).map(([row, i]) => `<tr>
    ${cell(i, 0, row[0], "ADVICS")}${cell(i, 1, row[1], "Тормозные диски")}${cell(i, 2, row[2], "Стандартные")}
    ${cell(i, 3, row[3], "Brembo")}${pct(i, 4, row[4])}${cell(i, 5, row[5], "")}${pct(i, 6, row[6])}${cell(i, 7, row[7], "")}${pct(i, 8, row[8])}
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="10" class="muted">Ничего не найдено</td></tr>`;
  bindCells($("#positioning"), list, renderPositioning);
}

function renderCars() {
  if (!S.settings) return;
  const list = S.settings.cars;
  const rows = filtered(list, $("#cars-search").value, [0, 1, 2]);
  $("#cars-count").textContent = `${list.length} строк${S.settings.raw && S.settings.raw.cars_custom ? " · ваш справочник" : ""}`;
  const options = (val) => S.settings.segments.map((s) => `<option value="${esc(s)}"${s === val ? " selected" : ""}>${esc(s)}</option>`).join("");
  $("#cars").innerHTML = rows.slice(0, 250).map(([row, i]) => `<tr>
    <td><input class="input" data-r="${i}" data-c="0" value="${esc(row[0])}" placeholder="Toyota"></td>
    <td><input class="input" data-r="${i}" data-c="1" value="${esc(row[1])}" placeholder="Camry"></td>
    <td><input class="input" data-r="${i}" data-c="2" value="${esc(row[2])}" placeholder="CAMRY, КАМРИ"></td>
    <td><input class="input" data-r="${i}" data-c="3" list="fleet-list" value="${esc(row[3])}"></td>
    <td><select data-r="${i}" data-c="4">${options(row[4] || "Средний")}</select></td>
    <td style="text-align:center"><input type="checkbox" data-r="${i}" data-c="5" ${String(row[5]).toLowerCase() === "да" ? "checked" : ""}></td>
    <td>${delBtn(i)}</td></tr>`).join("") || `<tr><td colspan="7" class="muted">Ничего не найдено</td></tr>`;
  let fleets = $("#fleet-list");
  if (!fleets) {
    fleets = document.createElement("datalist");
    fleets.id = "fleet-list";
    document.body.appendChild(fleets);
  }
  fleets.innerHTML = S.settings.fleets.map((f) => `<option value="${esc(f)}"></option>`).join("");
  bindCells($("#cars"), list, renderCars);
}

const toNum = (v) => { const n = parseFloat(String(v).replace(/\s/g, "").replace(",", ".")); return Number.isFinite(n) ? n : ""; };

async function saveSettings() {
  const s = S.settings;
  const body = {
    currency_rates: Object.fromEntries(s.rates.filter(([k]) => k.trim()).map(([k, v]) => [k.trim(), toNum(v)])),
    gz: s.gz.filter(([n, c]) => String(n).trim() && toNum(c) !== "").map(([n, c]) => [n, toNum(c)]),
    discounts: s.discounts.filter((r) => String(r[2]).trim() && toNum(r[5]) !== "").map((r) => [...r.slice(0, 5), toNum(r[5])]),
    cars: s.cars.filter((r) => String(r[0] || "").trim()).map((r) => r.slice(0, 6)),
    positioning: s.positioning.filter((r) => String(r[0] || "").trim() || String(r[2] || "").trim()).map((r) => r.slice(0, 9)),
    city_codes: s.cities.filter((r) => String(r[0] || "").trim() && String(r[1] || "").trim()).map((r) => [String(r[0]).trim(), String(r[1]).trim()]),
    suggest_threshold: Math.max(0, Math.min(100, toNum($("#threshold").value) || 0)) / 100,
    output: $("#output").value.trim() || "Результат",
    open_result: $("#open-result").checked,
  };
  try {
    applySettings(await api("/api/settings/save", body));
    S.opts.suggest_threshold = body.suggest_threshold;
    if (S.custom) S.custom.suggest_threshold = body.suggest_threshold;
    toast("Настройки сохранены", "good");
    if (S.scan) analyze();
  } catch (e) { toast(e.message, "bad"); }
}

start().catch((e) => {
  showBootError(e.message || "Не удалось запустить");
});
