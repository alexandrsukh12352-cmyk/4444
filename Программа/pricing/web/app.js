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
    price_source: "both", monitoring_only: true, in_stock_only: false, subgroup_mode: "sheets", suggest_threshold: 0.3,
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

function showBoot() {
  const boot = $("#boot");
  if (!boot) return;
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
  const stage = $("#boot-stage");
  if (stage && !stage.textContent) stage.textContent = "Остановка";
}

async function pollBoot() {
  if (S.bootHold) return;
  try {
    const res = await fetch("/api/progress");
    const p = await res.json();
    if (p.done && !p.error && p.started && p.started + 0.3 < (S.bootSince || 0)) return;
    const stage = $("#boot-stage");
    const detail = $("#boot-detail");
    const time = $("#boot-time");
    if (stage) stage.textContent = p.stage || "Жду ответ";
    if (detail) detail.textContent = p.detail || "";
    if (time) time.textContent = p.stage ? `${p.stage_elapsed} с на этом шаге · всего ${p.elapsed} с` : "";
    const log = $("#boot-log");
    if (log) {
      const rows = p.log || [];
      log.innerHTML = rows.map((row, i) => `<li class="${i === rows.length - 1 ? "on" : ""}"><b>${esc(row.stage)}</b>${row.detail ? ` · ${esc(row.detail)}` : ""}</li>`).join("");
    }
    if (p.error) showBootError(p.error);
  } catch (e) {
    const stage = $("#boot-stage");
    const detail = $("#boot-detail");
    if (stage) stage.textContent = "Нет ответа от программы";
    if (detail) detail.textContent = e.message || "";
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
    report_kind: S.init.report_kinds[o.report_kind] ? o.report_kind : "full",
    sheet_filters: Array.isArray(o.sheet_filters) ? o.sheet_filters.slice() : (S.init.default_sheet_filters || ["overstock", "sale", "competitor_price"]).slice(),
    green_when: o.green_when === "below" ? "below" : "above",
    money_base: o.money_base === "potential" ? "potential" : "sales",
    rec: { ...(S.init.rec || {}), ...(o.rec || {}), up: { ...(S.init.default_up || {}), ...((o.rec || {}).up || {}) } },
    ui: o.ui === "custom" ? "custom" : "template",
  });
  S.recCommon = clone(S.opts.rec);
  S.recMine = o.rec_mine && typeof o.rec_mine === "object" && Object.keys(o.rec_mine).length
    ? clone(o.rec_mine) : thresholdSnapshot(S.recCommon);
  S.recGroups = o.rec_groups && typeof o.rec_groups === "object" ? clone(o.rec_groups) : {};
  S.recBundled = clone(S.init.bundled_rec || {});
  S.recBundledGroups = clone(S.init.bundled_groups || {});
  S.recScope = o.rec_scope === "all" ? "all" : "group";
  S.recSource = o.rec_source === "yours" ? "yours" : (o.rec_source === "mine" ? "mine" : (Object.keys(S.recGroups).length ? "mine" : "yours"));
  S.recGroupSource = o.rec_group_source && typeof o.rec_group_source === "object" ? clone(o.rec_group_source) : {};
  S.recDirty = false;
  S.custom = seedCustom(o);
  if (S.opts.ui === "template") applyTemplate();
  else applyCustom();
  S.opts.filters = foldClearance(S.opts.filters);
  ensureStockProfiles();
  S.folder = S.init.folder;
  $("#folder").value = S.folder;
  $("#pick-folder").classList.toggle("hidden", !S.init.can_pick_folder);
  bind();
  renderPriceControls();
  renderSplit();
  renderRecForm();
  paintUi(S.opts.ui, false);
  bindReportKind();
  S.ready = true;
  const tag = $("#build-tag");
  if (tag) tag.textContent = S.init.build_label || "";
  renderSheetFilters();
  renderStockForm();
  const opened = await rescan();
  if (opened) hideBoot();
}

const CUSTOM_KEYS = ["base", "sources", "filters", "abc", "cross_modes", "cross_split", "price_source",
  "monitoring_only", "in_stock_only", "subgroup_mode", "suggest_threshold", "green_when"];
const REC_NUMS = ["tolerance", "down", "min_vr", "top_share", "mid_share", "top_factor", "mid_factor", "rounding"];

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
  const dynFrom = $("#px-dyn-from");
  if (dynFrom) dynFrom.onchange = () => { S.dynFrom = dynFrom.value; loadDynamics(); };
  const dynTo = $("#px-dyn-to");
  if (dynTo) dynTo.onchange = () => { S.dynTo = dynTo.value; loadDynamics(); };
  const dynLevel = $("#px-dyn-level");
  if (dynLevel) seg(dynLevel, S.dynLevel || "all", (v) => { S.dynLevel = v; paintDynTable(); });
  const dynSearch = $("#px-dyn-search");
  if (dynSearch) dynSearch.oninput = () => paintDynTable();
  const dynExport = $("#px-dyn-export");
  if (dynExport) dynExport.onclick = () => writeDynamicsExport();
  const pxOpen = $("#px-open");
  if (pxOpen) pxOpen.onclick = () => {
    if (!S.pxFolder) { toast("Сначала запишите устройство — папка результата откроется после этого"); return; }
    api("/api/open", { path: S.pxFolder }).catch((e) => toast(e.message, "bad"));
  };
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
  $("#rec-save").onclick = saveMine;
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
  S.opts.monitoring_only = true;
  S.opts.in_stock_only = false;
  S.opts.cross_split = "none";
  S.opts.cross_modes = {};
  S.opts.green_when = "above";
}

function paintUi(v, fromClick) {
  S.opts.ui = v === "custom" ? "custom" : "template";
  document.body.classList.toggle("detail", S.opts.ui === "custom");
  seg($("#ui-mode"), S.opts.ui, (next) => paintUi(next, true));
  const note = $("#template-note");
  if (note) {
    note.textContent = S.opts.ui === "template"
      ? "Все товары группы. Цены из АЦПКД и прайсов, кроссы — как подскажет программа. Пороги рекомендаций — свои у группы или общие."
      : "Фильтры, кроссы и источники общие для всех групп. Пороги рекомендаций запоминаются отдельно у каждой группы. «По шаблону» не стирает эти настройки.";
  }
  if (fromClick && S.opts.ui === "template") {
    applyTemplate();
    renderRecForm();
    if (S.scan) { renderFilters(); renderPriceControls(); renderSplit(); analyze(); }
  }
  if (fromClick && S.opts.ui === "custom") {
    applyCustom();
    renderRecForm();
    renderPriceControls();
    renderSplit();
    if (S.scan) { renderFilters(); renderGroup(); analyze(); }
  }
  if (fromClick) persist();
}

function bindReportKind() {
  const apply = (v) => {
    S.opts.report_kind = v;
    persist();
    seg($("#report-kind"), v, apply);
    seg($("#report-kind-main"), v, apply);
    updateSummary();
  };
  apply(S.opts.report_kind || "full");
}

function showTab(name) {
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
  $("#page-report").classList.toggle("hidden", name !== "report");
  $("#page-settings").classList.toggle("hidden", name !== "settings");
  const diag = $("#page-diagnose");
  if (diag) diag.classList.toggle("hidden", name !== "diagnose");
  const stock = $("#page-stock");
  if (stock) stock.classList.toggle("hidden", name !== "stock");
  const prices = $("#page-prices");
  if (prices) prices.classList.toggle("hidden", name !== "prices");
  $("#actionbar").classList.toggle("hidden", name !== "report");
  if (name === "stock") renderStockForm();
  if (name === "prices") { renderPriceAudit(); loadDynamics(); }
  if (name === "settings" && !S.settings) loadSettings();
  if (name === "diagnose") loadDiagnose();
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
  showBoot();
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
    const pricesPage = $("#page-prices");
    if (pricesPage && !pricesPage.classList.contains("hidden")) loadDynamics(true);
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
      <div class="hint">разбор распознавания — вкладка «Прайсы»</div>
    </div>`;
  const notes = [...sc.warnings.map((w) => `<div class="note">${esc(w)}</div>`), ...sc.skipped.map((w) => `<div class="note info">${esc(w)}</div>`)];
  const hasBase = sc.groups.pst.length || sc.groups.acp.length || (sc.queue || []).length;
  let empty = "";
  if (!sc.files) {
    empty = `<div class="empty"><h3>Папка пуста</h3><p>Положите выгрузки из 1С во «Входящие» или в папку группы: ПСТ, СБС, Продажи, АЦПКД. Кроссы и прайсы — в свои папки.</p><p class="mono">${esc(sc.folder)}</p></div>`;
  } else if (!hasBase) {
    empty = `<div class="empty"><h3>Не найдено ни ПСТ, ни АЦПКД</h3><p>Без одного из них отчёт не построить. Проверьте, что выгрузка сохранена как .xlsx или .csv и в ней есть шапка с названиями столбцов.</p></div>`;
  }
  $("#docs").innerHTML = `<div class="docs">${docs}${priceTile}</div>`;
  const scanNote = $("#scan-note");
  if (scanNote) scanNote.innerHTML = `${notes.length ? `<div class="notes">${notes.join("")}</div>` : ""}${empty}`;
  if (!sc.files || !hasBase) $("#card-folder").open = true;
  $("#flow").classList.toggle("hidden", !hasBase);
}

// ------------------------------------------------------------------ шаг 2: группа и основа

function availableBases() {
  const pst = (S.scan.groups.pst || []).length || (S.scan.queue || []).length;
  return { pst: pst > 0, acp: (S.scan.groups.acp || []).length > 0 };
}

function listedGroups() {
  if (!S.scan) return [];
  if (S.opts.base === "pst" && (S.scan.queue || []).length) return S.scan.queue;
  return (S.scan.groups && S.scan.groups[S.opts.base]) || [];
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

function saveMine() {
  S.recDirty = false;
  persist(true);
  toast("Мои пороги сохранены", "good");
  renderRecScope();
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
  toast("Пороги сброшены к рекомендованным", "good");
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
  toast("Все пороги сброшены к рекомендованным", "good");
}

function sameNum(a, b) {
  const x = Number(a);
  const y = Number(b);
  if (!Number.isFinite(x) && !Number.isFinite(y)) return true;
  return x === y;
}

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
  S.recDirty = true;
  paintRecMarks();
  renderRecScope();
}

function renderRecScope() {
  const el = $("#rec-scope");
  const who = currentSource() === "mine" ? "Мои" : "Рекомендованные";
  const where = S.recScope === "all" ? "одни на всех" : "эта группа";
  if (el) el.textContent = `${who}, ${where}`;
  const dirty = $("#rec-dirty");
  if (dirty) dirty.classList.toggle("hidden", !S.recDirty);
  seg($("#rec-scope-mode"), S.recScope === "all" ? "all" : "group", (v) => {
    S.recScope = v === "all" ? "all" : "group";
    applyGroupRec();
    renderRecForm();
    persist();
    analyze();
  });
  seg($("#rec-source"), currentSource(), (v) => {
    setSource(v);
    applyGroupRec();
    renderRecForm();
    persist();
    analyze();
  });
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
  const bases = availableBases();
  if (!bases[S.opts.base]) S.opts.base = bases.pst ? "pst" : "acp";
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
  const bases = availableBases();
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
        <span class="s">${esc(bits.join(" · "))}<span data-ownrec class="${ownRec ? "" : "hidden"}"> · мои пороги</span></span>
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
  const baseEl = $("#base");
  if (baseEl) {
    baseEl.querySelectorAll("button").forEach((b) => (b.disabled = !bases[b.dataset.v]));
    seg(baseEl, S.opts.base, (v) => { S.opts.base = v; S.opts.sources = {}; chooseGroup(); });
  }
  const baseNote = $("#base-note");
  if (baseNote) {
    baseNote.textContent = S.opts.base === "pst"
      ? "Товары — из ПСТ. Цены конкурентов и пропуски можно добирать из АЦПКД."
      : "Товары — из АЦПКД. Группа заказа, закупка, разновидность и оверсток — из ПСТ по коду товара.";
  }
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

function renderPriceFiles() {
  const files = S.priceFiles;
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
            ${f.monitored ? "" : `<span class="pill warn" title="Поставщика нет на листе 02 позиционирования">нет в позиционировании</span>`}</div>
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
    <div class="pf-list ${off ? "dim" : ""}">${blocks}</div>
    <div class="muted small-text" style="margin-top:8px">Прайсы одного поставщика свёрнуты. В отчёт попадает минимальная цена конкурента.</div>`;
  bindPriceFiles();
  $("#price-files").querySelectorAll("[data-supplier]").forEach((b) => (b.onclick = () => {
    S.priceOpen[b.dataset.supplier] = !S.priceOpen[b.dataset.supplier];
    renderPriceFiles();
  }));
  const pricesPage = $("#page-prices");
  if (pricesPage && !pricesPage.classList.contains("hidden")) renderPriceAudit();
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

function renderPriceAudit() {
  const kpis = $("#px-kpis");
  const list = $("#px-list");
  if (!kpis || !list) return;
  renderBrandHints();
  const filterEl = $("#px-filter");
  const filter = S.pxFilter || "all";
  if (filterEl) filterEl.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.v === filter));
  const files = S.priceFiles || [];
  const mon = (S.scan && S.scan.monitoring_suppliers) || [];
  const monBy = new Map(mon.map((m) => [supplierKey(m.supplier), m]));
  if (!S.scan) {
    kpis.innerHTML = "";
    list.innerHTML = `<div class="card"><div class="empty compact"><h3>Папка ещё не прочитана</h3><p>Откройте вкладку «Отчёт» — программа посмотрит «Данные/Прайсы».</p></div></div>`;
    return;
  }
  const groups = new Map();
  files.forEach((f) => {
    const key = supplierKey(f.supplier) || "—";
    if (!groups.has(key)) {
      const known = monBy.get(key);
      groups.set(key, { key, supplier: (known && known.supplier) || f.supplier || "Без поставщика", files: [] });
    }
    groups.get(key).files.push(f);
  });
  const rows = [...groups.values()].map((g) => {
    const expected = (monBy.get(g.key) && monBy.get(g.key).brands) || g.files[0].expected_brands || [];
    const found = new Map();
    g.files.forEach((f) => (f.competitor_brands || []).forEach((b) => {
      const k = brandKey(b.brand);
      found.set(k, (found.get(k) || 0) + (b.offers || 0));
    }));
    const brands = expected.map((name) => ({ name, offers: found.get(brandKey(name)) || 0, hit: found.has(brandKey(name)) }));
    g.found = found;
    g.brands = brands;
    g.expected = expected;
    g.bad = g.files.filter(priceBad).length;
    g.offers = g.files.reduce((sum, f) => sum + (f.offers || 0), 0);
    g.monitored = monBy.has(g.key);
    return g;
  });
  rows.sort((a, b) => {
    const ao = mon.findIndex((m) => supplierKey(m.supplier) === a.key);
    const bo = mon.findIndex((m) => supplierKey(m.supplier) === b.key);
    if (ao !== -1 || bo !== -1) return (ao === -1 ? 999 : ao) - (bo === -1 ? 999 : bo);
    return a.supplier.localeCompare(b.supplier, "ru");
  });
  const q = String(S.pxQuery || "").trim().toLowerCase();
  const hitFile = (f) => !q || [f.file, f.id, f.city, f.problem, f.folder, f.name_supplier, f.supplier,
    ...(f.competitor_brands || []).map((b) => b.brand)].join(" ").toLowerCase().includes(q);
  const visible = rows.map((g) => {
    const filesHit = g.files.filter((f) => hitFile(f) || (q && g.supplier.toLowerCase().includes(q)));
    const brandHit = !q || g.supplier.toLowerCase().includes(q) || g.brands.some((b) => b.name.toLowerCase().includes(q));
    return { ...g, files: brandHit && g.supplier.toLowerCase().includes(q) ? g.files : (filesHit.length ? filesHit : g.files), show: (brandHit || filesHit.length) && matchPxFilter(g, filter) };
  }).filter((g) => g.show);
  const gaps = mon.filter((m) => !groups.has(supplierKey(m.supplier))).filter((m) => {
    if (filter !== "all" && filter !== "gap") return false;
    if (!q) return true;
    return `${m.supplier} ${m.brands.join(" ")}`.toLowerCase().includes(q);
  });
  let covered = 0;
  let pairs = 0;
  mon.forEach((m) => {
    const g = groups.get(supplierKey(m.supplier));
    m.brands.forEach((brand) => {
      pairs += 1;
      if (g && g.found.has(brandKey(brand))) covered += 1;
    });
  });
  const badFiles = files.filter(priceBad).length;
  const recognized = files.filter((f) => f.status !== "none" && f.offers > 0).length;
  kpis.innerHTML = [
    ["Файлов", fmtInt(files.length), "в папке Прайсы"],
    ["Распознано", fmtInt(recognized), badFiles ? `${fmtInt(badFiles)} с проблемой` : "без проблем"],
    ["Поставщиков", fmtInt(rows.length), "по папке или имени файла"],
    ["Цен", fmtInt(files.reduce((s, f) => s + (f.offers || 0), 0)), "строк с артикулом и ценой"],
    ["Лист 02", pairs ? `${fmtInt(covered)} из ${fmtInt(pairs)}` : "—", "пар поставщик–бренд нашлись"],
  ].map(([l, v, s]) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");
  if (!files.length && !gaps.length) {
    list.innerHTML = `<div class="card"><div class="empty compact"><h3>Прайсов пока нет</h3>
      <p>Создайте в Данные/Прайсы папку с именем поставщика — как на листе 02 — и положите туда его файлы. Например: Данные/Прайсы/TISS/TissEKB.xls.</p></div></div>`;
    return;
  }
  if (!S.pxOpen) S.pxOpen = {};
  const cards = filter === "gap" ? "" : visible.map((g) => pxCard(g)).join("");
  const gapCard = gaps.length && (filter === "all" || filter === "gap") ? `<div class="card">
    <div class="card-head"><h2>В позиционировании, прайса нет</h2><span class="sub">${fmtInt(gaps.length)}</span></div>
    <div class="px-brands" style="margin-bottom:10px">Эти поставщики есть на листе 02, папки или файла с их именем программа не увидела.</div>
    ${gaps.map((m) => `<div style="margin-bottom:8px"><b>${esc(m.supplier)}</b>
      <span class="px-brands">${m.brands.map((b) => `<span class="tag miss">${esc(b)}</span>`).join(" ")}</span></div>`).join("")}
  </div>` : "";
  const emptyFilter = !cards && !gapCard
    ? `<div class="card"><div class="empty compact"><h3>По этому отбору пусто</h3><p>Смените фильтр или очистите поиск.</p></div></div>` : "";
  list.innerHTML = cards + gapCard + emptyFilter;
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

function pxCard(g) {
  const trouble = g.bad > 0 || g.brands.some((b) => !b.hit);
  const editingHere = !!(S.editWhere === "prices" && S.editing && g.files.some((f) => f.id === S.editing.id));
  const open = editingHere || (Object.prototype.hasOwnProperty.call(S.pxOpen, g.key) ? S.pxOpen[g.key] : (trouble || g.files.length <= 2));
  const brands = g.brands.length
    ? g.brands.map((b) => `<span class="tag ${b.hit ? "hit" : "miss"}">${esc(b.name)}${b.hit ? ` <b>${fmtInt(b.offers)}</b>` : ""}</span>`).join("")
    : `<span class="muted">этого поставщика нет на листе 02</span>`;
  const body = open ? `<div class="table-wrap"><table>
    <thead><tr><th>Файл</th><th>Город</th><th>Откуда поставщик</th><th>Статус</th><th class="num">Цен</th><th>Бренды в файле</th><th>Замечание</th><th></th></tr></thead>
    <tbody>${g.files.map(pxRow).join("")}</tbody></table></div>${editingHere ? editorHtml() : ""}` : "";
  return `<div class="card px-card">
    <button type="button" class="px-head" data-px="${esc(g.key)}" data-open="${open ? "1" : "0"}">
      <span>${open ? "▾" : "▸"}</span>
      <b>${esc(g.supplier)}</b>
      ${g.monitored ? `<span class="pill good">в мониторинге</span>` : `<span class="pill warn">нет в позиционировании</span>`}
      <span class="muted">${fmtInt(g.files.length)} ${plural(g.files.length, "файл", "файла", "файлов")} · ${fmtInt(g.offers)} цен</span>
      ${g.bad ? `<span class="pill bad">${fmtInt(g.bad)} с проблемой</span>` : ""}
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
  const where = f.shape || (f.header_row ? `шапка, строка ${f.header_row}${cols ? ` · ${cols}` : ""}` : (f.problem ? "" : "шапка не найдена"));
  const editing = S.editing && S.editWhere === "prices" && S.editing.id === f.id;
  return `<tr class="${priceBad(f) ? "px-bad" : ""}">
    <td><b>${esc(f.file)}</b><div class="muted">${esc([f.date, f.sheet].filter(Boolean).join(" · "))}</div></td>
    <td>${esc(f.city) || "—"}</td>
    <td>${esc(from)}${where ? `<div class="muted">${esc(where)}</div>` : ""}</td>
    <td><span class="pill ${kind}">${st}</span></td>
    <td class="num">${fmtInt(f.offers)}</td>
    <td>${brands || `<span class="muted">${f.offers ? "конкурентов из листа 02 нет" : "—"}</span>`}</td>
    <td>${notes.length ? esc(notes.join(". ")) : "—"}</td>
    <td><button type="button" class="btn small" data-px-edit="${esc(f.id)}">${editing ? "Закрыть" : "Править"}</button></td>
  </tr>`;
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
        <div class="ed-preview-head"><b>Что получится</b><span class="muted">${fmtInt(file.offers)} цен · ${fmtInt(file.brands)} брендов</span></div>
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
  if (resetCols) ed.draft.cols = {};
  clearTimeout(dryTimer);
  dryTimer = setTimeout(async () => {
    try {
      const r = await api("/api/prices/save", { folder: S.folder, id: ed.id, ...ed.draft, dry_run: true });
      if (S.editing !== ed) return;
      ed.file = r.file;
      if (resetCols) ed.draft.cols = { ...r.file.cols };
      renderPriceFiles();
    } catch (e) { toast(e.message, "bad"); }
  }, 150);
}

async function savePrice(body) {
  const heavy = !body.dry_run && (body.header_row != null || body.cols || body.reset || body.remember);
  if (heavy) showBoot();
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

async function loadDynamics() {
  const note = $("#px-dyn-note");
  if (!note || S.dynLoading) return;
  S.dynLoading = true;
  note.textContent = "Считаю динамику…";
  try {
    const body = { folder: S.folder };
    if (S.dynFrom) body.date_from = S.dynFrom;
    if (S.dynTo) body.date_to = S.dynTo;
    S.dynamics = await api("/api/prices/dynamics", body);
    S.dynFrom = S.dynamics.date_from || "";
    S.dynTo = S.dynamics.date_to || "";
    if (S.dynamics.price_files && !S.editing) {
      // Прайсы прошлых дат читаются только для динамики — теперь их видно и в списке.
      S.priceFiles = S.dynamics.price_files;
      if (S.scan) S.scan.price_files = S.dynamics.price_files;
      renderPriceAudit();
    }
    renderDynamics();
  } catch (e) {
    note.textContent = e.message;
  } finally {
    S.dynLoading = false;
  }
}

function renderDynamics() {
  const d = S.dynamics;
  if (!d) return;
  const cache = $("#px-dyn-cache");
  if (cache) cache.textContent = d.cached ? "из кэша" : "";
  const note = $("#px-dyn-note");
  if (!d.dates || d.dates.length < 2) {
    note.textContent = "Нужны две папки с датами: Данные/Прайсы/12.10.2026/TISS и Данные/Прайсы/13.10.2026/TISS. Файл «валюта» положите прямо в «Прайсы».";
    $("#px-dyn-body").innerHTML = "";
  } else if (!d.has_groups) {
    note.textContent = "Группы берутся из ПСТ: кросс «Да» связывается с категорией, группой и подгруппой нашего артикула. Пока ПСТ нет — показано общее изменение по брендам.";
  } else {
    note.textContent = `Только бренды из позиционирования. Артикулов в обеих датах: ${fmtInt(d.compared)}. Цена — медиана. Ниже — как она разошлась с курсом.`;
  }
  const options = (d.dates || []).map((dt) => `<option value="${esc(dt)}">${esc(dt)}</option>`).join("");
  const from = $("#px-dyn-from");
  const to = $("#px-dyn-to");
  if (from) { from.innerHTML = options; from.value = d.date_from || ""; }
  if (to) { to.innerHTML = options; to.value = d.date_to || ""; }
  const rateText = (r) => r.change == null ? "нет даты" : fmtPct(r.change, true);
  const rateNums = (r) => (r.old != null && r.new != null)
    ? `${Number(r.old).toLocaleString("ru-RU", { maximumFractionDigits: 4 })} → ${Number(r.new).toLocaleString("ru-RU", { maximumFractionDigits: 4 })}`
    : "";
  $("#px-dyn-rates").innerHTML = (d.rates || []).length
    ? d.rates.map((r) => `<span class="tag">${esc(r.code)} <b>${rateText(r)}</b> <span class="muted">${rateNums(r)}</span></span>`).join("")
    : `<span class="muted">Файла «валюта» в папке прайсов нет</span>`;
  const level = $("#px-dyn-level");
  if (level) level.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.v === (S.dynLevel || "all")));
  if (d.dates && d.dates.length >= 2) {
    paintDynSummary();
    paintDynTable();
  } else if ($("#px-dyn-summary")) $("#px-dyn-summary").innerHTML = "";
}

function paintDynSummary() {
  const d = S.dynamics;
  const host = $("#px-dyn-summary");
  if (!d || !host) return;
  const rows = (d.rows || []).filter((r) => r.brand && !r.category && !r.group && !r.subgroup);
  rows.sort((a, b) => Math.abs(b.change || 0) - Math.abs(a.change || 0));
  const rates = (d.rates || []).filter((r) => r.change != null);
  if (!rows.length) {
    host.innerHTML = `<div class="muted">В обеих датах нет цен брендов из позиционирования.</div>`;
    return;
  }
  const conclusions = (d.conclusions || []).map((line) => `<div>${esc(line)}</div>`).join("");
  host.innerHTML = (conclusions ? `<div class="dyn-brand">${conclusions}</div>` : "") + rows.map((r) => {
    const change = r.change || 0;
    const dir = change > 0.005 ? "подорожал" : change < -0.005 ? "подешевел" : "почти не изменился";
    const lines = rates.map((rate) => {
      const gap = change - rate.change;
      const word = gap > 0.005 ? "выше курса" : gap < -0.005 ? "ниже курса" : "рядом с курсом";
      return `${esc(rate.code)} ${fmtPct(rate.change, true)} · цена ${word} ${fmtPp(gap)}`;
    }).join(" · ");
    return `<div class="dyn-brand"><div><b>${esc(r.brand)}</b> ${dir} на ${fmtPct(Math.abs(change))} <span class="muted">${fmtInt(r.n)} арт.</span></div><div class="muted">${lines}</div></div>`;
  }).join("");
}

function paintDynTable() {
  const d = S.dynamics;
  const host = $("#px-dyn-body");
  if (!d || !host) return;
  const level = S.dynLevel || "all";
  const q = (($("#px-dyn-search") && $("#px-dyn-search").value) || "").trim().toLowerCase();
  const fit = (r) => {
    if (level === "category") return !r.group && !r.subgroup && !!r.category;
    if (level === "group") return !!r.group && !r.subgroup;
    if (level === "subgroup") return !!r.subgroup;
    return !r.category && !r.group && !r.subgroup;
  };
  const rows = (d.rows || []).filter((r) => fit(r) && (!q || `${r.category} ${r.group} ${r.subgroup} ${r.brand}`.toLowerCase().includes(q)));
  rows.sort((a, b) => Math.abs(b.change || 0) - Math.abs(a.change || 0));
  if (!rows.length) {
    host.innerHTML = `<div class="muted">На этом уровне строк нет.</div>`;
    return;
  }
  const rates = d.rates || [];
  const head = rates.map((r) => `<th class="num" title="Цена минус изменение курса">${esc(r.code)}</th>`).join("");
  const body = rows.slice(0, 300).map((r) => {
    const name = [r.category, r.group, r.subgroup].filter(Boolean).join(" / ") || "Все";
    const cells = rates.map((rate) => {
      if (r.change == null || rate.change == null) return `<td class="num muted">—</td>`;
      return `<td class="num">${fmtPp(r.change - rate.change)}</td>`;
    }).join("");
    return `<tr><td>${esc(name)}</td><td>${esc(r.brand || "все бренды")}</td><td class="num">${fmtInt(r.n)}</td><td class="num">${r.change == null ? "—" : fmtPct(r.change, true)}</td>${cells}</tr>`;
  }).join("");
  host.innerHTML = `<div class="table-wrap"><table><thead><tr><th>Срез</th><th>Бренд</th><th class="num">Артикулов</th><th class="num">Цена</th>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

async function writeDynamicsExport() {
  const btn = $("#px-dyn-export");
  if (btn) btn.disabled = true;
  showBoot();
  try {
    const r = await api("/api/prices/dynamics-export", { folder: S.folder, date_from: S.dynFrom, date_to: S.dynTo });
    S.pxFolder = r.folder;
    hideBoot();
    toast(`Записан ${r.name}: ${fmtInt(r.rows)} строк по артикулам`, "good");
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

function renderRecForm() {
  const r = S.opts.rec;
  $("#rec-enabled").checked = !!r.enabled;
  $("#rec-body").classList.toggle("off", !r.enabled);
  const targetButtons = [
    ["first", "Только первый", "Только первый в списке позиционирования. Для дисков это Brembo. Нет его цены — на остальных не ориентируемся."],
    ["main", "По списку", "Сначала первый в списке. Нет цены — следующий, со своим процентом."],
    ["all", "Все с ценой", "Среднее или самая низкая цель среди всех, у кого есть цена."],
  ];
  const group = r.target === "avg" || r.target === "min" ? "all" : (r.target === "first" ? "first" : "main");
  $("#rec-target").innerHTML = targetButtons.map(([k, t, tip]) =>
    `<button type="button" data-v="${k}" title="${esc(tip)}">${esc(t)}</button>`).join("");
  seg($("#rec-target"), group, (v) => {
    r.target = v === "all" ? (r.target === "min" ? "min" : "avg") : v;
    renderRecForm();
    analyze();
  });
  const all = $("#rec-all");
  all.classList.toggle("hidden", group !== "all");
  seg(all, r.target === "min" ? "min" : "avg", (v) => { r.target = v; analyze(); });
  const note = targetButtons.find(([k]) => k === group);
  $("#rec-target-note").textContent = note ? note[2] : "";
  seg($("#green-when"), S.opts.green_when, (v) => { S.opts.green_when = v; persist(); if (S.analysis) renderDash(); });
  const set = (id, key) => { $(id).value = r[key] ?? ""; };
  set("#rec-tolerance", "tolerance"); set("#rec-down", "down"); set("#rec-vr", "min_vr"); set("#rec-round", "rounding");
  set("#rec-top", "top_share"); set("#rec-top-k", "top_factor"); set("#rec-mid", "mid_share"); set("#rec-mid-k", "mid_factor");
  $("#rec-up").innerHTML = (S.init.segments || []).map((s) => `<label>${esc(s)}
      <input class="input num-input" data-seg="${esc(s)}" type="number" min="0" max="40" step="0.5" value="${esc(r.up[s] ?? "")}"></label>`).join("");
  $("#rec-enabled").onchange = (e) => {
    r.enabled = e.target.checked;
    $("#rec-body").classList.toggle("off", !r.enabled);
    analyze();
  };
  const bindNum = (id, apply) => {
    $(id).oninput = () => {
      const n = toNum($(id).value);
      if (n === "") return;
      apply(n);
      analyze();
    };
  };
  bindNum("#rec-tolerance", (n) => { r.tolerance = n; rememberGroupRec(); });
  bindNum("#rec-down", (n) => { r.down = n; rememberGroupRec(); });
  bindNum("#rec-vr", (n) => { r.min_vr = n; rememberGroupRec(); });
  bindNum("#rec-round", (n) => { r.rounding = n; rememberGroupRec(); });
  bindNum("#rec-top", (n) => { r.top_share = n; rememberGroupRec(); });
  bindNum("#rec-top-k", (n) => { r.top_factor = n; rememberGroupRec(); });
  bindNum("#rec-mid", (n) => { r.mid_share = n; rememberGroupRec(); });
  bindNum("#rec-mid-k", (n) => { r.mid_factor = n; rememberGroupRec(); });
  $("#rec-up").querySelectorAll("input").forEach((inp) => (inp.oninput = () => {
    const n = toNum(inp.value);
    if (n === "") return;
    r.up[inp.dataset.seg] = n;
    rememberGroupRec();
    analyze();
  }));
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

function renderDash() {
  const d = S.analysis && S.analysis.dashboard;
  const el = $("#dash");
  if (!d) { el.innerHTML = ""; return; }
  const recOn = !!d.rec_enabled;
  const k = d.kinds || {};
  const nowSlice = (d.slices || []).find((s) => s.kind === "all") || {};
  const recSlice = (d.slices || []).find((s) => s.kind === "rec") || {};
  const max = Math.max(1, ...d.buckets.map((b) => Math.max(b.now, b.rec || 0)));
  const bars = d.buckets.map((b) => {
    const nowW = Math.round((b.now / max) * 100);
    const recW = Math.round(((b.rec || 0) / max) * 100);
    return `<div class="bar-row">
      <span>${esc(b.label)}</span>
      ${recOn ? `<div class="bar-pair">
          <div class="bar-track"><i class="${tone(b.label)}" style="width:${nowW}%"></i></div>
          <div class="bar-track"><i class="after" style="width:${recW}%"></i></div>
        </div>
        <span class="num">${fmtInt(b.now)} → ${fmtInt(b.rec)}</span>`
      : `<div class="bar-track"><i class="${tone(b.label)}" style="width:${nowW}%"></i></div>
        <span class="num">${fmtInt(b.now)}</span>`}
    </div>`;
  }).join("");
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
  const comps = (d.slices || []).filter((s) => s.kind === "comp" && s.n);
  const compRows = comps.map((s) => `<tr>
      <td>${esc(String(s.title).replace(" (без учёта % позиц.)", ""))}</td>
      <td class="num">${fmtInt(s.n)}</td>
      <td class="num">${fmtPct(s.mean, true)}</td>
      ${recOn ? `<td class="num after">${fmtPct(s.mean_rec, true)}</td><td class="num delta ${deltaKind(s.mean, s.mean_rec, "closer")}">${s.mean != null && s.mean_rec != null ? fmtPp(s.mean_rec - s.mean) : "—"}</td>` : ""}
      <td class="num">${fmtPct(s.median, true)}</td>
    </tr>`).join("");
  const movers = (d.top_changes || []).slice(0, 8);
  const moverRows = movers.map((x) => `<tr>
      <td class="nowrap"><b>${esc(x.article)}</b></td>
      <td class="muted ellip" title="${esc(x.name)}">${esc(x.name)}</td>
      <td class="num">${fmtMoney(x.cat_a)}</td>
      <td class="num after">${fmtMoney(x.rec)}</td>
      <td class="num">${x.change == null ? "—" : fmtPct(x.change, true)}</td>
      <td class="num delta ${deltaKind(0, x.profit_delta, "up")}">${fmtSignedRub(x.profit_delta)}</td>
    </tr>`).join("");
  const compare = recOn ? `
    <div class="dash-title">Сейчас и после рекомендаций</div>
    <div class="table-wrap"><table class="cmp">
      <thead><tr><th>Показатель</th><th class="num">Сейчас</th><th class="num">После</th><th class="num">Изменение</th></tr></thead>
      <tbody>
        ${cmpRow("ВР группы", fmtPct(d.vr), fmtPct(d.vr_rec), d.vr != null && d.vr_rec != null ? fmtPp(d.vr_rec - d.vr) : "—", deltaKind(d.vr, d.vr_rec, "up"))}
        ${cmpRow("Выручка", fmtRub(d.sale), fmtRub(d.sale_rec), d.sale ? fmtPct(d.sale_rec / d.sale - 1, true) : "—", deltaKind(d.sale, d.sale_rec, "up"))}
        ${cmpRow("Прибыль", fmtRub(d.profit), fmtRub(d.profit_rec), d.profit_rec != null && d.profit != null ? fmtSignedRub(d.profit_rec - d.profit) : "—", deltaKind(d.profit, d.profit_rec, "up"))}
        ${cmpRow("Среднее отклонение", fmtPct(nowSlice.mean, true), fmtPct(recSlice.mean, true), nowSlice.mean != null && recSlice.mean != null ? fmtPp(recSlice.mean - nowSlice.mean) : "—", deltaKind(nowSlice.mean, recSlice.mean, "closer"))}
        ${cmpRow("Отклонение с весом продаж", fmtPct(nowSlice.weighted, true), fmtPct(recSlice.weighted, true), nowSlice.weighted != null && recSlice.weighted != null ? fmtPp(recSlice.weighted - nowSlice.weighted) : "—", deltaKind(nowSlice.weighted, recSlice.weighted, "closer"))}
        ${cmpRow("Медиана отклонения", fmtPct(nowSlice.median, true), fmtPct(recSlice.median, true), nowSlice.median != null && recSlice.median != null ? fmtPp(recSlice.median - nowSlice.median) : "—", deltaKind(nowSlice.median, recSlice.median, "closer"))}
        ${cmpRow("Среднее без выбросов", fmtPct(nowSlice.iqr, true), fmtPct(recSlice.iqr, true), nowSlice.iqr != null && recSlice.iqr != null ? fmtPp(recSlice.iqr - nowSlice.iqr) : "—", deltaKind(nowSlice.iqr, recSlice.iqr, "closer"))}
        ${cmpRow(`В пределах ±${d.tolerance}%`, d.with_target ? `${fmtInt(d.in_band)} из ${fmtInt(d.with_target)}` : "—", d.with_target && d.in_band_rec != null ? `${fmtInt(d.in_band_rec)} из ${fmtInt(d.with_target)}` : "—", d.with_target && d.in_band_rec != null ? fmtSignedInt(d.in_band_rec - d.in_band) : "—", d.with_target ? deltaKind(d.in_band, d.in_band_rec, "up") : "")}
        ${cmpRow(`Ниже ${fmtInt(d.min_vr)}% ВР`, fmtInt(d.low_vr), fmtInt(d.low_vr_rec), d.low_vr_rec == null ? "—" : fmtSignedInt(d.low_vr_rec - d.low_vr), deltaKind(d.low_vr, d.low_vr_rec, "down"))}
      </tbody>
    </table></div>
    <div class="cmp-note">Плюс в отклонении — мы дешевле цели, минус — дороже. Зелёное изменение отклонения — ближе к цели.</div>
    <div class="rec-pills">
      <span class="pill up">поднять ${fmtInt((k.up || 0) + (k.vr || 0))}</span>
      <span class="pill down">понизить ${fmtInt(k.down || 0)}</span>
      <span class="pill keep">не трогать ${fmtInt(k.keep || 0)}</span>
      ${k.vr ? `<span class="pill up">из них до мин. ВР ${fmtInt(k.vr)}</span>` : ""}
    </div>` : `
    <div class="kpis">
      <div class="kpi"><div class="l">Товаров / с ценой</div><div class="v">${fmtInt(d.items)}</div><div class="s">${fmtInt(d.priced)} · ${pct(d.priced, d.items)}%</div></div>
      <div class="kpi"><div class="l">ВР группы</div><div class="v">${fmtPct(d.vr)}</div><div class="s">ниже ${fmtInt(d.min_vr)}% — ${fmtInt(d.low_vr)} шт.</div></div>
      <div class="kpi"><div class="l">В позиционировании ±${esc(d.tolerance)}%</div><div class="v">${fmtInt(d.in_band)}</div><div class="s">из ${fmtInt(d.with_target)}</div></div>
      <div class="kpi"><div class="l">Среднее отклонение</div><div class="v">${fmtPct(nowSlice.mean, true)}</div><div class="s">медиана ${fmtPct(nowSlice.median, true)}</div></div>
    </div>
    <div class="cmp-note">Галочка «вставить рек. цены» включает сравнение «сейчас → после».</div>`;
  const sliceRows = (d.slices || []).filter((s) => s.kind === "all" || s.kind === "abc").map((s) => `<tr>
      <td>${esc(s.title)}</td><td class="num">${fmtInt(s.n)}</td>
      <td class="num">${fmtPct(s.mean, true)}</td><td class="num">${fmtPct(s.median, true)}</td>
      <td class="num">${fmtPct(s.iqr, true)}</td><td class="num">${fmtInt(s.outliers)}</td></tr>`).join("");
  const cleanNow = fmtPct(nowSlice.iqr, true);
  const cleanAfter = fmtPct(recSlice.iqr, true);
  const live = recOn
    ? `ВР ${fmtPct(d.vr)} → ${fmtPct(d.vr_rec)}${d.vr != null && d.vr_rec != null ? ` (${fmtPp(d.vr_rec - d.vr)})` : ""}. Отклонение без выбросов ${cleanNow} → ${cleanAfter}. Прибыль ${d.profit != null && d.profit_rec != null ? fmtSignedRub(d.profit_rec - d.profit) : "—"}.`
    : `ВР группы ${fmtPct(d.vr)}. Отклонение без выбросов ${cleanNow}. Включите галочку выше — цифры пересчитаются сразу.`;
  el.innerHTML = `
    <div class="live-line">${live}</div>
    ${compare}
    <div class="dash-cols">
      <div>
        <div class="dash-title">Отклонение от цели${recOn ? "" : ""}</div>
        ${recOn ? `<div class="legend"><span>верхняя полоска — сейчас</span><span><i class="after"></i>нижняя — после рекомендаций</span></div>` : ""}
        <div class="bars">${bars}</div>
      </div>
      <div>
        ${cutTable("ВР по ABC", d.abc || [], groupHead("ABC"), groupBody(d.abc || []))}
      </div>
    </div>
    ${recOn ? cutTable("По подгруппам", (d.subgroups || []).length > 1 ? d.subgroups : [], groupHead("Подгруппа"), groupBody(d.subgroups || [])) : ""}
    ${recOn ? cutTable("По сегментам автопарка", d.segments || [], groupHead("Сегмент"), groupBody(d.segments || [])) : ""}
    ${comps.length ? cutTable("Среднее отклонение от цены конкурента", comps, `<th>Конкурент</th><th class="num">Товаров</th><th class="num">Сейчас</th>${recOn ? `<th class="num">После</th><th class="num">Δ</th>` : ""}<th class="num">Медиана</th>`, compRows) : ""}
    ${recOn && movers.length ? cutTable("Что сильнее меняет прибыль", movers, `<th>Артикул</th><th>Наименование</th><th class="num">Кат. А</th><th class="num">Рек.</th><th class="num">Изм.</th><th class="num">Δ прибыли</th>`, moverRows) : ""}
    ${recOn ? "" : `<div class="table-wrap" style="margin-bottom:10px"><table class="mini"><thead><tr>
      <th>Отклонение от цели</th><th class="num">Товаров</th><th class="num">Среднее</th><th class="num">Медиана</th><th class="num">Без выбросов</th><th class="num">Выбросов</th>
    </tr></thead><tbody>${sliceRows}</tbody></table></div>`}
    <ul class="conclusions">${(d.conclusions || []).map((t) => `<li>${esc(t)}</li>`).join("")}</ul>`;
}

// ------------------------------------------------------------------ покрытие ценами конкурентов

const isoDate = (v) => String(v || "").split("-").reverse().join(".");

function renderCoverage() {
  const el = $("#coverage");
  if (!el) return;
  const a = S.analysis;
  const f = (a && a.funnel) || [];
  if (!a || !f.length) { el.innerHTML = ""; return; }
  const sc = a.screen || { enabled: true, ratio: 3 };
  const screened = Object.entries(a.screened || {}).map(([k, n]) => `${esc(k)} — ${fmtInt(n)}`).join("; ");
  const rows = f.map((c) => {
    const r = c.reliability || {};
    const lead = (r.leaders || []).map((x) => `${esc(x.supplier)} ${fmtInt(x.rows)}`).join(", ");
    return `<tr>
      <td><b>${esc(c.name)}</b></td>
      <td class="num">${fmtInt(c.items)}</td>
      <td class="num">${fmtInt(c.yes)}</td>
      <td class="num">${fmtInt(c.in_prices)}</td>
      <td class="num">${fmtInt(c.monitored)}</td>
      <td class="num">${c.excluded ? fmtInt(c.excluded) : "—"}</td>
      <td class="num"><b>${fmtInt(c.final)}</b>${c.by_no ? `<div class="muted">из них по «Нет» ${fmtInt(c.by_no)}</div>` : ""}</td>
      <td class="num">${c.sold_share == null ? "—" : fmtPct(c.sold_share)}</td>
      <td class="num">${r.priced ? fmtInt(r.single) : "—"}</td>
      <td class="num">${r.second_gap == null ? "—" : fmtPct(r.second_gap, true)}</td>
      <td class="muted">${lead || "—"}</td>
    </tr>`;
  }).join("");
  const actions = f.flatMap((c) => (c.actions || []).map((t) => `<li><b>${esc(c.name)}.</b> ${esc(t)}</li>`)).join("");
  const h = (a.history || []).slice().reverse();
  const history = h.length ? `
    <div class="dash-title">Как менялось по сборкам отчёта</div>
    <div class="table-wrap"><table class="mini"><thead><tr>
      <th>Дата</th><th class="num">Товаров</th><th class="num">С ценой</th><th class="num">Доля продаж с ценой</th>
      <th class="num">В позиционировании</th><th class="num">Медиана откл.</th><th class="num">Откл. с весом</th><th class="num">ВР</th>
    </tr></thead><tbody>${h.map((x) => `<tr>
      <td>${esc(isoDate(x.date))}</td><td class="num">${fmtInt(x.items)}</td><td class="num">${fmtInt(x.priced)}</td>
      <td class="num">${x.sold_share == null ? "—" : fmtPct(x.sold_share)}</td>
      <td class="num">${x.with_target ? `${fmtInt(x.in_band)} из ${fmtInt(x.with_target)}` : "—"}</td>
      <td class="num">${fmtPct(x.median, true)}</td><td class="num">${fmtPct(x.weighted, true)}</td><td class="num">${fmtPct(x.vr)}</td>
    </tr>`).join("")}</tbody></table></div>` : "";
  el.innerHTML = `
    <div class="dash-title">Покрытие ценами конкурентов</div>
    <div class="row tight" style="flex-wrap:wrap; gap:12px; margin-bottom:8px">
      <label><input type="checkbox" id="screen-on" ${sc.enabled ? "checked" : ""}> Отсеивать несопоставимое: комплекты и цены, отличающиеся от нашей в</label>
      <input class="input num-input" id="screen-ratio" type="number" min="1.5" step="0.5" value="${esc(sc.ratio)}" style="width:64px"><span class="muted">раза и больше</span>
      <button class="btn small" id="gaps-export" type="button">Выгрузить дыры в кроссах</button>
    </div>
    ${screened ? `<div class="muted" style="font-size:12.5px; margin-bottom:8px">Отсеяно пар: ${screened}. Список — в выгрузке дыр, лист «Несопоставимо».</div>` : ""}
    <div class="table-wrap"><table class="mini"><thead><tr>
      <th>Конкурент</th><th class="num">Товаров</th><th class="num">Кросс «Да»</th><th class="num">Есть в прайсах</th>
      <th class="num">У поставщиков листа 02</th><th class="num">Отсеяно</th><th class="num">С ценой</th><th class="num">Доля продаж с ценой</th>
      <th class="num">Один поставщик</th><th class="num">Второй дороже на</th><th>Кто чаще даёт минимум</th>
    </tr></thead><tbody>${rows}</tbody></table></div>
    ${actions ? `<ul class="conclusions">${actions}</ul>` : ""}
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
  if (!hints.length && !aliases.length) {
    el.classList.add("hidden");
    el.innerHTML = "";
    return;
  }
  el.classList.remove("hidden");
  el.innerHTML = `
    <div class="card-head"><h2>Написание брендов</h2>
      <span class="sub">Бренд в прайсе или в кроссах похож на конкурента, но не совпал. «Считать как» — и его цены пойдут в сравнение</span></div>
    ${hints.length ? `<div class="table-wrap"><table class="mini"><thead><tr>
      <th>Как написано</th><th class="num">Строк</th><th>Где</th><th>Похож на</th><th></th></tr></thead><tbody>
      ${hints.map((h, i) => `<tr><td><b>${esc(h.brand)}</b></td><td class="num">${fmtInt(h.rows)}</td>
        <td class="muted">${esc((h.where || []).join(", "))}</td><td>${esc(h.suggest)}</td>
        <td><button type="button" class="btn small" data-alias="${i}">Считать как ${esc(h.suggest)}</button></td></tr>`).join("")}
    </tbody></table></div>` : `<div class="muted">Похожих несовпавших брендов нет.</div>`}
    ${aliases.length ? `<div class="px-brands" style="margin-top:10px">Уже заданы:
      ${aliases.map((x) => `<span class="tag hit">${esc(x[0])} → ${esc(x[1])}
        <button type="button" class="btn small ghost" data-unalias="${esc(x[0])}">убрать</button></span>`).join(" ")}</div>` : ""}`;
  el.querySelectorAll("[data-alias]").forEach((b) => (b.onclick = () => {
    const h = hints[Number(b.dataset.alias)];
    if (h) saveAlias({ add: [h.brand, h.suggest] });
  }));
  el.querySelectorAll("[data-unalias]").forEach((b) => (b.onclick = () => saveAlias({ remove: b.dataset.unalias })));
}

async function saveAlias(body) {
  try {
    await api("/api/brand-aliases", body);
  } catch (e) {
    toast(e.message, "bad");
    return;
  }
  toast("Сохранено — перечитываю прайсы", "good");
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
function analyze() {
  persist();
  clearTimeout(analyzeTimer);
  const state = $("#analyze-state");
  if (state) state.textContent = "Фильтр выбран, пересчёт чуть позже…";
  analyzeTimer = setTimeout(runAnalyze, 450);
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
      <div class="muted" style="font-size:13px">${esc(r.kind === "quick" ? "Быстрый" : "Полный")} отчёт · ${fmtInt(r.rows)} товаров, ${fmtInt(r.priced)} с ценой конкурента, исключено ${fmtInt(r.excluded)} · листы: ${r.sheets.map(esc).join(", ")}</div>
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
        <div class="muted" style="font-size:13px">Файлы Excel — только группы с полным набором: ПСТ, СБС, Продажи и АЦПКД. В разборе — каждая группа из списка слева, даже если файл не создан. Вид — ${esc(S.opts.report_kind === "quick" ? "быстрый" : "полный")}.</div>
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
  seg($("#sk-kind"), p.report_kind === "quick" ? "quick" : "full", (v) => { p.report_kind = v; persistStock(); });
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
  const fallback = p.base === "acp" ? "acp_own" : "own_acp";
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
  renderRates(); renderGz(); renderDiscounts(); renderPositioning(); renderCars();
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
